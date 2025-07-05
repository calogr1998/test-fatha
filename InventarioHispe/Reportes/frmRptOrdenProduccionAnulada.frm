VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRptOrdenProduccionAnulada 
   Caption         =   "Reporte de Orden(es) de Producción Anuladas"
   ClientHeight    =   8745
   ClientLeft      =   330
   ClientTop       =   2040
   ClientWidth     =   13650
   Icon            =   "frmRptOrdenProduccionAnulada.frx":0000
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
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtCodProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         ToolTipText     =   "Seleccione Producto (F2)"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtProducto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Text            =   "Text1"
         ToolTipText     =   "Ingrese cadena a buscar"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   3855
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
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Filtrar"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   240
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
      Tools           =   "frmRptOrdenProduccionAnulada.frx":058A
      ToolBars        =   "frmRptOrdenProduccionAnulada.frx":A9D1
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgReporte 
      Height          =   7800
      Left            =   120
      OleObjectBlob   =   "frmRptOrdenProduccionAnulada.frx":AB02
      TabIndex        =   0
      Top             =   840
      Width           =   13410
   End
End
Attribute VB_Name = "frmRptOrdenProduccionAnulada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intOpcionResumen As Integer

Private Sub listarAnnos()
    Dim rstAnnoPA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "YEAR(DBO.FECHA(FECHA)) AS ANNO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "ANULADO = 1 "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "YEAR(DBO.FECHA(FECHA)) "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "YEAR(DBO.FECHA(FECHA))"
    
    If rstAnnoPA.State = 1 Then rstAnnoPA.Close
    
    rstAnnoPA.Open SqlCad, cnBdStudioModa, adOpenDynamic, adLockOptimistic
    
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
    SqlCad = SqlCad & "MONTH(DBO.FECHA(FECHA)) AS MES "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ORDENPRODUCCION "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "ANULADO = 1 AND "
    SqlCad = SqlCad & "YEAR(DBO.FECHA(FECHA)) = " & cmbAnno.Text & " "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "MONTH(DBO.FECHA(FECHA))"
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "MONTH(DBO.FECHA(FECHA))"
    
    If rstMesPA.State = 1 Then rstMesPA.Close
    
    rstMesPA.Open SqlCad, cnBdStudioModa, adOpenForwardOnly, adLockOptimistic
    
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

Private Sub limpiarCajas()
    txtBusqueda.Text = vbNullString
    
    txtCodProducto.Text = vbNullString
    txtProducto.Text = vbNullString
End Sub

Private Sub procesarConsulta()
    On Error GoTo errProcesarConsulta
    
    Screen.MousePointer = vbHourglass
    
    ModMilano.visualizarOrdenProduccionAnulada dbgReporte, cmbAnno.Text, right(cmbMes.Text, 2), Trim(txtBusqueda.Text)
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errProcesarConsulta:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbInformation + vbOKOnly, App.ProductName & " - ProcesarConsulta"
    
    Err.Clear
End Sub

Private Sub cmbAnno_Click()
    listarMeses
End Sub

Private Sub dbgReporte_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "CANTTOTAL", "CANTIDADTOTAL", "CANTIDAD"
            Text = Format(Text, "#0.00")
        Case "ESTADO"
            If UCase(Text) = "ANULADO" Then
                Color = vbRed
                FontColor = vbWhite
            Else
                Color = vbWhite
                FontColor = vbBlue
            End If
            
            Font.Bold = True
    End Select
End Sub

Private Sub Form_Load()
    limpiarCajas
    
    listarAnnos
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
