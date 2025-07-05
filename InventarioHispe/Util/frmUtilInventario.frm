VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilInventario 
   Caption         =   "Inventario de Producto"
   ClientHeight    =   9015
   ClientLeft      =   330
   ClientTop       =   1365
   ClientWidth     =   16410
   Icon            =   "frmUtilInventario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   16410
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkMostrarProductoDescontinuado 
      Caption         =   "Incluir Productos Descontinuados."
      Height          =   255
      Left            =   15720
      TabIndex        =   22
      Top             =   600
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   " Proveedor "
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
      Left            =   11160
      TabIndex        =   19
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCodProveedor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProveedor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Proceso "
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
      Left            =   1560
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   13335
      Begin MSComctlLib.ProgressBar pgbProceso1 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblProceso1 
         Caption         =   "Proceso 1"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   12735
      End
   End
   Begin VB.CheckBox chkMostrarNegativo 
      Caption         =   "Mostrar solo Productos con Stock en Negativo"
      Height          =   255
      Left            =   15720
      TabIndex        =   15
      Top             =   240
      Width           =   3735
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgInventario 
      Height          =   7680
      Left            =   120
      OleObjectBlob   =   "frmUtilInventario.frx":058A
      TabIndex        =   14
      Top             =   960
      Width           =   16170
   End
   Begin MSComDlg.CommonDialog cmdlgInventario 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFecha 
      Caption         =   " Hasta "
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
      Left            =   6480
      TabIndex        =   10
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cmbAnno 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   116916225
         CurrentDate     =   41939
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   " Ingresar cadena a buscar "
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   6015
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbInventario 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   3
      ToolsCount      =   15
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmUtilInventario.frx":19F6
      ToolBars        =   "frmUtilInventario.frx":CBEE
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgInventario1 
      Height          =   3690
      Left            =   2040
      OleObjectBlob   =   "frmUtilInventario.frx":CF58
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   8745
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmUtilInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolObviarCierre As Boolean
Private bolObviarTomaInv As Boolean

Private Sub listarAlmacen()
    Dim rstAlmacen As New ADODB.Recordset
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM MAESTROS.EF2ALMACENES ORDER BY F2CODALM", cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    Else
        rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    
    tlbInventario.Tools("Almacen").ComboBox.Clear
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        Do While Not rstAlmacen.EOF
            tlbInventario.Tools("Almacen").ComboBox.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            If tlbInventario.Tools("Almacen").ComboBox.ListCount > 0 Then
                tlbInventario.Tools("Almacen").ComboBox.ListIndex = 0
            End If
    End If
End Sub

Private Sub listarFamilia()
    objAyudaFamilia.listarFamilia tlbInventario.Tools.ITEM("Familia").ComboBox
    
    If tlbInventario.Tools.ITEM("Familia").ComboBox.ListCount > 1 Then
        tlbInventario.Tools.ITEM("Familia").ComboBox.ListIndex = 1
    End If
    
    If tlbInventario.Tools.ITEM("Familia").ComboBox.ListIndex = 0 Then
        tlbInventario.Tools.ITEM("SubFamilia").Enabled = False
    Else
        tlbInventario.Tools.ITEM("SubFamilia").Enabled = True
    End If
End Sub

Private Sub listarSubFamilia()
    With objAyudaSubFamilia
        .CodigoFamilia = Trim(right(tlbInventario.Tools.ITEM("Familia").ComboBox.Text, 4))
        
        .listarSubFamilia tlbInventario.Tools.ITEM("SubFamilia").ComboBox
        
        If .CodigoFamilia = vbNullString Then
            tlbInventario.Tools.ITEM("SubFamilia").Enabled = False
        Else
            tlbInventario.Tools.ITEM("SubFamilia").Enabled = True
        End If
    End With
End Sub

Private Sub listarAnnosVale()
    'objAyudaVale.listarAnnoValeSoloSeleccion cmbAnno
    
    objAyudaVale.listarAnnoValeSoloSeleccion cmbAnno
    
    If cmbAnno.ListCount > 0 Then
        cmbAnno.ListIndex = cmbAnno.ListCount - 1
    End If
End Sub

Private Sub listarMesesVale()
    'objAyudaVale.listarMesValeSoloSeleccion cmbMes, Trim(cmbAnno.Text)
    
    objAyudaVale.listarMesValeSoloSeleccion cmbMes, Trim(cmbAnno.Text)
    
    If cmbMes.ListCount > 0 Then
        cmbMes.ListIndex = cmbMes.ListCount - 1
    End If
End Sub

Private Sub listarInventario()

    Screen.MousePointer = vbHourglass
    
    dbgInventario.Dataset.Close
    
        With objAyudaVale
            .inicializarEntidades
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
            bolObviarCierre = True
            
            tlbInventario.Tools.ITEM("CerrarMes").State = IIf(.verificarCierreVale, ssChecked, ssUnchecked)
            
            'tlbInventario.Tools.ITEM("TomarInventario").Enabled = True
            tlbInventario.Tools.ITEM("TomarInventario").Enabled = Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State)
            
            If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
            Else
                tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
            End If
            
            bolObviarCierre = False
        End With
        
        With objAyudaTomaInventario
            .inicializarEntidades
            
            .AnnoTI = Trim(cmbAnno.Text)
            .MesTI = Format(right(cmbMes.Text, 2), "00")
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            bolObviarTomaInv = True
            
            tlbInventario.Tools.ITEM("TomarInventario").State = IIf(.verificarExistencia, ssChecked, ssUnchecked)
            
            If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
                tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
            Else
                tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
            End If
            
            .obtenerConfigTomaInventario
    '
    '        If CBool(tlbInventario.Tools.ITEM("TomarInventario").Enabled) Then
    '            tlbInventario.Tools.ITEM("TomarInventario").Enabled = Not .CierreInventario
    '        End If
            
            If .CierreInventario Then
                bolObviarTomaInv = True
                
                tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
                
                tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Toma Inventario Cerrada"
                
                tlbInventario.Tools.ITEM("TomarInventario").Enabled = False
                
                bolObviarTomaInv = False
            End If
            
            bolObviarTomaInv = False
        End With
        
        objAyudaVale.listarGrillaInventarioProductoV2 dbgInventario, _
                                                    txtBusqueda.Text, _
                                                    Trim(dtpHasta.value & ""), _
                                                    CBool(tlbInventario.Tools.ITEM("Valorizado").State), _
                                                    Trim(left(tlbInventario.Tools.ITEM("Moneda").ComboBox.Text, 1)), _
                                                    CBool(tlbInventario.Tools.ITEM("Agrupar").State), _
                                                    Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), _
                                                    Trim(right(tlbInventario.Tools.ITEM("Familia").ComboBox.Text, 4)), _
                                                    Trim(right(tlbInventario.Tools.ITEM("SubFamilia").ComboBox.Text, 4)), _
                                                    Nothing, _
                                                    CBool(chkMostrarNegativo.value), _
                                                    Trim(txtCodProveedor.Text), _
                                                    CBool(chkMostrarProductoDescontinuado.value)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cerrarMes()
    Dim strFechaCorteInicialDeValesParaCP As String
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    
    strFechaCorteInicialDeValesParaCP = ModUtilitario.sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "FechaCorteInicialDeValesParaCP", "l")
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With objAyudaVale
            If Val(Trim(cmbAnno.Text) & Format(right(cmbMes.Text, 2), "00")) > Val(Format(CDate(strFechaCorteInicialDeValesParaCP), "yyyymm")) Then
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                intAnnoCorte = Val(cmbAnno.Text) - IIf(Val(right(cmbMes.Text, 2)) > 1, 0, 1)
                intMesCorte = IIf(Val(right(cmbMes.Text, 2)) > 1, Val(right(cmbMes.Text, 2)) - 1, 12)
                
                .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
                
                .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
                .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0) 'dtpHasta.value
                
                If Not .verificarCierreVale Then
                    MsgBox "Imposible cerrar el Periodo del Almacen seleccionado; ya que el anterior aun se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    bolObviarCierre = True
                    
                    tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
                    
                    If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                    Else
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                    End If
                    
                    bolObviarCierre = False
                    
                    Exit Sub
                End If
            End If
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
            If .verificarCierreVale Then
                MsgBox "Mes del Almacen y Ejercicio seleccionado ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
                
                If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                Else
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
        
        
        With objAyudaVale
            If Val(Trim(cmbAnno.Text) & Format(right(cmbMes.Text, 2), "00")) > Val(Format(CDate(strFechaCorteInicialDeValesParaCP), "yyyymm")) Then
                .inicializarEntidades
                .inicializarEntidadesAdicionales
                
                intAnnoCorte = Val(cmbAnno.Text) - IIf(Val(right(cmbMes.Text, 2)) > 1, 0, 1)
                intMesCorte = IIf(Val(right(cmbMes.Text, 2)) > 1, Val(right(cmbMes.Text, 2)) - 1, 12)
                
                .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
                
                .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
                .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0) 'dtpHasta.value
                
                If Not .verificarCierreVale Then
                    MsgBox "Imposible cerrar el Periodo del Almacen seleccionado; ya que el anterior aun se encuentra abierto, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    bolObviarCierre = True
                    
                    tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
                    
                    If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                    Else
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                    End If
                    
                    bolObviarCierre = False
                    
                    Exit Sub
                End If
            End If
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
            If .verificarCierreVale Then
                MsgBox "Mes del Almacen y Ejercicio seleccionado ya se encuentra cerrado.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
                
                If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                Else
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
        
        With objAyudaTomaInventario
            .inicializarEntidades
            
            .AnnoTI = Trim(cmbAnno.Text)
            .MesTI = Format(right(cmbMes.Text, 2), "00")
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            If .verificarExistencia Then
            
                .obtenerConfigTomaInventario
                
                If Not .CierreInventario Then
                    MsgBox "Toma de Inventario no ha sido cerrada, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    bolObviarCierre = True
                    
                    tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
                    
                    If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                    Else
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                    End If
                    
                    bolObviarCierre = False
                    
                    Exit Sub
                End If
            End If
        End With
    End If
    
    With objAyudaVale
        If MsgBox("¿Desea cerrar el Periodo " & ModUtilitario.devuelveNombreMes(Format(right(cmbMes.Text, 2), "00")) & "-" & Trim(cmbAnno.Text) & "?" & vbNewLine & _
                    "RECOMENDACIÓN: Antes de proceder con el cierre del Periodo en mención, asegurese de haber ejecutado el recalculo del Costo.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .VB1 = True
            .VB1Usuario = wusuario
            .VB1Fecha = Format(Date, "Short Date")
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
'            If .generarCierreMensual(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)), Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))) Then
                If .cerrarVale Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        With objAyudaVale
                            .inicializarEntidades
                            .inicializarEntidadesAdicionales
                            
                            .VB1 = True
                            .VB1Usuario = wusuario
                            .VB1Fecha = Format(Date, "Short Date")
                            
                            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
                            
                            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
                            .FechaFinMes = dtpHasta.value
                            
                            If .cerrarVale Then
                                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                            End If
                            
                            .inicializarEntidades
                            .inicializarEntidadesAdicionales
                        End With
                    End If
                    
                    MsgBox "Periodo cerrado correctamente.", vbInformation + vbOKOnly, App.ProductName
                    
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
'                    End With
                Else
                    bolObviarCierre = True
                    
                    tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
                    
                    bolObviarCierre = False
                End If
'            Else
'                bolObviarCierre = True
'
'                tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
'
'                bolObviarCierre = False
'            End If
        Else
            bolObviarCierre = True
            
            tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
            
            bolObviarCierre = False
        End If
    End With
    
    tlbInventario.Tools.ITEM("TomarInventario").Enabled = Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State)
End Sub

Private Sub abrirMes()
    Dim intAnnoCorte As Integer
    Dim intMesCorte As Integer
    
    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
        With objAyudaVale
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            intAnnoCorte = Val(cmbAnno.Text) + IIf(Val(right(cmbMes.Text, 2)) < 12, 0, 1)
            intMesCorte = IIf(Val(right(cmbMes.Text, 2)) < 12, Val(right(cmbMes.Text, 2)) + 1, 1)
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
            .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0) 'dtpHasta.value
            
            If .verificarCierreVale Then
                MsgBox "Imposible abrir el Periodo del Almacen seleccionado; ya que el posterior se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
                
                If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                Else
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
            If Not .verificarCierreVale Then
                MsgBox "Mes del Almacen y Ejercicio seleccionado ya se encuentra abierto.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
                
                If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                Else
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
    Else
        With objAyudaVale
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            intAnnoCorte = Val(cmbAnno.Text) + IIf(Val(right(cmbMes.Text, 2)) < 12, 0, 1)
            intMesCorte = IIf(Val(right(cmbMes.Text, 2)) < 12, Val(right(cmbMes.Text, 2)) + 1, 1)
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(intAnnoCorte, intMesCorte + 0, 1)
            .FechaFinMes = DateSerial(intAnnoCorte, intMesCorte + 1, 0) 'dtpHasta.value
            
            If .verificarCierreVale Then
                MsgBox "Imposible abrir el Periodo del Almacen seleccionado; ya que el posterior se encuentra cerrado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
                
                If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                Else
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
            If Not .verificarCierreVale Then
                MsgBox "Mes del Almacen y Ejercicio seleccionado ya se encuentra abierto.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarCierre = True
                
                tlbInventario.Tools.ITEM("CerrarMes").State = ssUnchecked
                
                If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
                Else
                    tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Abrir &Mes"
                End If
                
                bolObviarCierre = False
                
                Exit Sub
            End If
        End With
    End If
    
    With objAyudaVale
        If MsgBox("¿Desea abrir el Periodo " & ModUtilitario.devuelveNombreMes(Format(right(cmbMes.Text, 2), "00")) & "-" & Trim(cmbAnno.Text) & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            
            .inicializarEntidades
            .inicializarEntidadesAdicionales
            
            .VB1 = False
            .VB1Usuario = wusuario
            .VB1Fecha = Format(Date, "Short Date")
            
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
            .FechaFinMes = dtpHasta.value
            
'            If .generarCierreMensual(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)), Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), True) Then
                If .cerrarVale Then
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        With objAyudaVale
                            .inicializarEntidades
                            .inicializarEntidadesAdicionales
                            
                            .VB1 = False
                            .VB1Usuario = wusuario
                            .VB1Fecha = Format(Date, "Short Date")
                            
                            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
                            
                            .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
                            .FechaFinMes = dtpHasta.value
                            
                            If .cerrarVale Then
                                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                            End If
                            
                            .inicializarEntidades
                            .inicializarEntidadesAdicionales
                        End With
                    End If
                    
                    MsgBox "Periodo abierto correctamente.", vbInformation + vbOKOnly, App.ProductName
                    
'                    With objAyudaVale
'                        .inicializarEntidades
'                        .inicializarEntidadesAdicionales
'
'                        .FechaInicioMes = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 0, 1)
'                        .FechaFinMes = dtpHasta.value
                        
                        'tlbInventario.Tools.ITEM("CerrarMes").Enabled = Not .verificarCierreVale
                        tlbInventario.Tools("CerrarMes").ChangeAll ssChangeAllName, "Cerrar &Mes"
'                    End With
                Else
                    bolObviarCierre = True
                    
                    tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
                    
                    bolObviarCierre = False
                End If
'            Else
'                bolObviarCierre = True
'
'                tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
'
'                bolObviarCierre = False
'            End If
        Else
            bolObviarCierre = True
            
            tlbInventario.Tools.ITEM("CerrarMes").State = ssChecked
            
            bolObviarCierre = False
        End If
    End With
    
    tlbInventario.Tools.ITEM("TomarInventario").Enabled = Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State)
End Sub

Private Sub tomarInventario()
    Screen.MousePointer = vbHourglass
        
'    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'        With objSqlAyudaTomaInventario
'            .inicializarEntidades
'            .inicializarEntidadesDetalle
'
'            .AnnoTI = Trim(cmbAnno.Text)
'            .MesTI = Format(right(cmbMes.Text, 2), "00")
'            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
'
'            If .verificarExistencia Then
'                MsgBox "Toma de Inventario ya realizada, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                bolObviarTomaInv = True
'
'                tlbInventario.Tools.ITEM("TomarInventario").State = ssChecked
'
'                If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
'                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
'                Else
'                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
'                End If
'
'                bolObviarTomaInv = False
'
'                Screen.MousePointer = vbDefault
'
'                Exit Sub
'            End If
'        End With
'    Else
        With objAyudaTomaInventario
            .inicializarEntidades
            .inicializarEntidadesDetalle
            
            .AnnoTI = Trim(cmbAnno.Text)
            .MesTI = Format(right(cmbMes.Text, 2), "00")
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            If .verificarExistencia Then
                MsgBox "Toma de Inventario ya realizada, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarTomaInv = True
                
                tlbInventario.Tools.ITEM("TomarInventario").State = ssChecked
                
                If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
                Else
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
                End If
                
                bolObviarTomaInv = False
                
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End With
'    End If
    
        
    If MsgBox("¿Desea generar la Toma de Inventario hasta el '" & dtpHasta.value & "'?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        bolObviarTomaInv = True
        
        tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
        
        If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
            tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
        Else
            tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
        End If
        
        bolObviarTomaInv = False
        
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    
    Dim rstTemporal As New ADODB.Recordset
    
    fraBusqueda.Enabled = False
    fraFecha.Enabled = False
    chkMostrarNegativo.Enabled = False
    tlbInventario.Enabled = False
    
    fraProceso.Visible = True
    pgbProceso1.value = 0
    lblProceso1.Caption = "Guardando Toma de Inventarios..."
    
'    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'        objAyudaVale.listarGrillaInventarioProductoV2 Nothing, _
'                                                            vbNullString, _
'                                                            Trim(dtpHasta.value & ""), _
'                                                            False, _
'                                                            ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l"), _
'                                                            False, _
'                                                            Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), _
'                                                            vbNullString, _
'                                                            vbNullString, _
'                                                            Nothing, _
'                                                            False, _
'                                                            vbNullString, _
'                                                            True, _
'                                                            "tmpCPInventario" & wusuario
'    Else
        objAyudaVale.listarGrillaInventarioProductoV2 Nothing, _
                                                        vbNullString, _
                                                        Trim(dtpHasta.value & ""), _
                                                        False, _
                                                        ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l"), _
                                                        False, _
                                                        Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2)), _
                                                        vbNullString, _
                                                        vbNullString, _
                                                        Nothing, _
                                                        False, _
                                                        vbNullString, _
                                                        True
'    End If
    
    With objAyudaTomaInventario
        .inicializarEntidades
        .inicializarEntidadesDetalle
        
        .Fecha = InputBox("Ingrese o Confirme la fecha de Toma de Inventario:", "Tomar Inventario", Format(dtpHasta.value, "Short Date"))
        
        If Not IsDate(.Fecha) Then
            .Fecha = Format(dtpHasta.value, "Short Date")
        End If
        
        .Observacion = UCase(InputBox("Ingrese alguna observación:", "Tomar Inventario", vbNullString))
        
        .FecReg = Format(Date, "Short Date")
        .UsuReg = wusuario
        
        If .guardarTomaInventario Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
'            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                With objSqlAyudaTomaInventario
'                    .inicializarEntidades
'                    .inicializarEntidadesDetalle
'
'                    .Fecha = objAyudaTomaInventario.Fecha
'                    .Observacion = objAyudaTomaInventario.Observacion
'                    .FecReg = objAyudaTomaInventario.FecReg
'                    .UsuReg = objAyudaTomaInventario.UsuReg
'
'                    If .guardarTomaInventario Then
'                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
'                    End If
'                End With
'            End If
            
            If rstTemporal.State = 1 Then rstTemporal.Close
            
            'rstTemporal.Open objAyudaVale.SQLSelectAlter, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            rstTemporal.Open objAyudaVale.SQLSelectAlter, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
            
            If Not rstTemporal.EOF Then
                rstTemporal.MoveFirst
                
                pgbProceso1.Max = ModUtilitario.devuelveCantRegistros(rstTemporal)
                pgbProceso1.value = 0
                lblProceso1.Caption = "Registrando Toma de Inventario..."
                
                Do While Not rstTemporal.EOF
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstTemporal!f5codpro & "")
                    .StockSistema = Val(rstTemporal!SALDO & "")
                    
                    .guardarTomaInvDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
'                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                        With objSqlAyudaTomaInventario
'                            .inicializarEntidadesDetalle
'
'                            .CodigoProducto = objAyudaTomaInventario.CodigoProducto
'                            .StockSistema = objAyudaTomaInventario.StockSistema
'
'                            .guardarTomaInvDetalleOneByOne
'
'                            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
'                        End With
'                    End If
                    
                    DoEvents
                                
                    fraProceso.Visible = True
                    pgbProceso1.value = pgbProceso1.value + 1
                    lblProceso1.Caption = "Registrando Toma de Inventario: " & Trim(rstTemporal!FAMILIA & "") & " / " & Trim(rstTemporal!SUBFAMILIA & "") & " / " & left(Trim(rstTemporal!F5NOMPRO & ""), 100) & " (" & Trim(rstTemporal!F7SIGMED & "") & ") " & "... " & FormatPercent(pgbProceso1.value / pgbProceso1.Max, 3)
                    
                    rstTemporal.MoveNext
                Loop
            End If
            
            MsgBox "Toma de Inventario finalizado correctamente." & vbNewLine & _
                    "Productos registrados: " & pgbProceso1.Max, vbInformation + vbOKOnly, App.ProductName
            
            bolObviarTomaInv = True
            
            tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
            
            bolObviarTomaInv = False
        Else
            bolObviarTomaInv = True
            
            tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
            
            bolObviarTomaInv = False
        End If
    End With
    
    fraBusqueda.Enabled = True
    fraFecha.Enabled = True
    chkMostrarNegativo.Enabled = True
    tlbInventario.Enabled = True
    
    fraProceso.Visible = False
    pgbProceso1.value = 0
    lblProceso1.Caption = vbNullString
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub eliminarInventario()
    Screen.MousePointer = vbHourglass
        

        With objAyudaTomaInventario
            .inicializarEntidades
            .inicializarEntidadesDetalle
            
            .AnnoTI = Trim(cmbAnno.Text)
            .MesTI = Format(right(cmbMes.Text, 2), "00")
            .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
            
            If Not .verificarExistencia Then
                MsgBox "Toma de Inventario no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarTomaInv = True
                
                tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
                
                If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
                Else
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
                End If
                
                bolObviarTomaInv = False
                
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
            
            .obtenerConfigTomaInventario
            
            If .CierreInventario Then
                MsgBox "Toma de Inventario cerrada, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarTomaInv = True
                
                tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
                
                If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Toma Inventario Cerrada"
                Else
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
                End If
                
                bolObviarTomaInv = False
                
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
            
            If .ValeIngreso <> vbNullString Or .ValeSalida <> vbNullString Then
                MsgBox "Imposible eliminar Toma de Inventario, cuenta con Vale(s) generado(s) a pesar de estar Abierto; verifique.", vbInformation + vbOKOnly, App.ProductName
                
                bolObviarTomaInv = True
                
                tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
                
                If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Toma Inventario Cerrada"
                Else
                    tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
                End If
                
                bolObviarTomaInv = False
                
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End With
'    End If
        
    If MsgBox("¿Desea eliminar Toma de Inventario del Periodo seleccionado?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        bolObviarTomaInv = True
        
        tlbInventario.Tools.ITEM("TomarInventario").State = ssChecked
        
        If Not CBool(tlbInventario.Tools.ITEM("TomarInventario").State) Then
            tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
        Else
            tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "Eliminar &Toma Inventario"
        End If
        
        bolObviarTomaInv = False
        
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    
    With objAyudaTomaInventario
        .inicializarEntidades
        .inicializarEntidadesDetalle
        
        .AnnoTI = Trim(cmbAnno.Text)
        .MesTI = Format(right(cmbMes.Text, 2), "00")
        .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
        
        If .eliminarTomaInventario Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
'            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                With objSqlAyudaTomaInventario
'                    .inicializarEntidades
'                    .inicializarEntidadesDetalle
'
'                    .AnnoTI = Trim(cmbAnno.Text)
'                    .MesTI = Format(right(cmbMes.Text, 2), "00")
'                    .CodigoAlmacen = Trim(right(tlbInventario.Tools.ITEM("Almacen").ComboBox.Text, 2))
'
'                    If .eliminarTomaInventario Then
'                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
'                    End If
'                End With
'            End If
            
            MsgBox "Toma de Inventario eliminado correctamente.", vbInformation + vbOKOnly, App.ProductName
            
            bolObviarTomaInv = True
            
            tlbInventario.Tools("TomarInventario").ChangeAll ssChangeAllName, "&Tomar Inventario"
            
            bolObviarTomaInv = False
        Else
            bolObviarTomaInv = True
            
            tlbInventario.Tools.ITEM("TomarInventario").State = ssUnchecked
            
            bolObviarTomaInv = False
        End If
    End With
    
    fraBusqueda.Enabled = True
    fraFecha.Enabled = True
    chkMostrarNegativo.Enabled = True
    tlbInventario.Enabled = True
    
    fraProceso.Visible = False
    pgbProceso1.value = 0
    lblProceso1.Caption = vbNullString
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkMostrarNegativo_Click()
    'listarInventario
End Sub

Private Sub chkMostrarProductoDescontinuado_Click()
    'listarInventario
End Sub

Private Sub cmbAnno_Click()
    listarMesesVale
End Sub

Private Sub cmbMes_Click()
    DoEvents
    
    dtpHasta.value = DateSerial(Val(cmbAnno.Text), Val(right(cmbMes.Text, 2)) + 1, 0)
    
    'listarInventario
End Sub

Private Sub dbgInventario_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "SALDO", "COSTOPROMEDIO", "TOTALACTUAL", "TOTALANTERIOR"
            If Val(Text) < 0 Then
                FontColor = vbRed
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
        Case "DIFERENCIA"
            If Val(Text) = 0 Then
                FontColor = RGB(230, 185, 184)
            ElseIf Val(Text) < 0 Then
                FontColor = vbRed
            ElseIf Val(Text) > 0 Then
                FontColor = vbBlue
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgInventario_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "SALDO", "COSTOPROMEDIO", "TOTALACTUAL", "TOTALANTERIOR", "DIFERENCIA"
            If Val(Text) < 0 Then
                FontColor = vbRed
            Else
                FontColor = vbBlue
            End If
            
            Select Case Trim(Mid(tlbInventario.Tools.ITEM("Moneda").ComboBox.Text, 1, 1))
                Case "S"
                    Color = vbWhite
                Case "D"
                    Color = &HC0FFC0
            End Select
            Font.Bold = True
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgInventario_OnCustomDrawFooterNode(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal FooterIndex As Integer, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "COSTOPROMEDIO", "TOTALACTUAL", "TOTALANTERIOR", "DIFERENCIA"
            If Val(Text) < 0 Then
                FontColor = vbRed
            End If
            
            Select Case Trim(Mid(tlbInventario.Tools.ITEM("Moneda").ComboBox.Text, 1, 1))
                Case "S"
                    Color = vbWhite
                Case "D"
                    Color = &HC0FFC0
            End Select
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dtpHasta_CloseUp()
    'listarInventario
End Sub

Private Sub Form_Load()
    txtBusqueda.Text = vbNullString
    txtCodProveedor.Text = vbNullString
    lblProveedor.Caption = vbNullString
    
    tlbInventario.Tools.ITEM("Moneda").Enabled = CBool(tlbInventario.Tools.ITEM("Valorizado").State)
    tlbInventario.Tools.ITEM("Agrupar").Enabled = CBool(tlbInventario.Tools.ITEM("Valorizado").State)
    
    listarAlmacen
    
    listarFamilia
    
    listarSubFamilia
    
    listarAnnosVale
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgInventario.Move 0, fraBusqueda.Height + 300, Me.ScaleWidth, (Me.ScaleHeight - (fraBusqueda.Height + 300))
End Sub

Private Sub tlbInventario_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Almacen"
            'listarInventario
        Case "Familia"
            listarSubFamilia
            
            'listarInventario
        Case "SubFamilia"
            'listarInventario
        Case "Moneda"
            'listarInventario
    End Select
End Sub

Private Sub tlbInventario_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Consultar"
            listarInventario
        Case "Valorizado"
            tlbInventario.Tools.ITEM("Moneda").Enabled = CBool(Tool.State)
            tlbInventario.Tools.ITEM("Agrupar").Enabled = CBool(Tool.State)
            
            If Not CBool(Tool.State) Then
                tlbInventario.Tools("Valorizado").ChangeAll ssChangeAllName, "Valo&rizado"
            Else
                tlbInventario.Tools("Valorizado").ChangeAll ssChangeAllName, "No Valo&rizado"
            End If
            
            'listarInventario
        Case "Agrupar"
            If Not CBool(Tool.State) Then
                tlbInventario.Tools("Agrupar").ChangeAll ssChangeAllName, "&Agrupar"
            Else
                tlbInventario.Tools("Agrupar").ChangeAll ssChangeAllName, "Des&agrupar"
            End If
            
            'listarInventario
        Case "Filtrar"
            If Not CBool(Tool.State) Then
                tlbInventario.Tools("Filtrar").ChangeAll ssChangeAllName, "&Filtrar"
            Else
                tlbInventario.Tools("Filtrar").ChangeAll ssChangeAllName, "Quitar &Filtro"
            End If
            
            dbgInventario.Filter.FilterActive = CBool(Tool.State)
        Case "ExportarExcel"
'            If Not CBool(tlbInventario.Tools.ITEM("CerrarMes").State) Then
'                MsgBox "Cierre el Mes antes de descargar el Reporte.", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
            
            Screen.MousePointer = vbHourglass
            
            With cmdlgInventario
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString

                .ShowSave

                If .FileName <> vbNullString Then
                    dbgInventario.m.ExportToXLS .FileName

                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "CerrarMes"
            If bolObviarCierre Then
                Exit Sub
            End If
            
            If CBool(Tool.State) Then
                cerrarMes
            Else
                abrirMes
            End If
        Case "TomarInventario"
            If bolObviarTomaInv Then
                Exit Sub
            End If
            
            If CBool(Tool.State) Then
                
                    tomarInventario
            Else
                    eliminarInventario
            End If
    End Select
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarInventario
    End Select
End Sub

Private Sub txtCodProveedor_DblClick()
    txtCodProveedor_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            Me.MousePointer = vbHourglass
            
            wcodcliprov = vbNullString
            
            With Ayuda_Proveedores
                .Show 1
            End With
            
            If wcodcliprov <> vbNullString Then
                txtCodProveedor.Text = wcodcliprov
                lblProveedor.Caption = wnomcliprov
                
                listarInventario
            Else
                MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                ModUtilitario.seleccionarTextoCaja txtCodProveedor
                
                Exit Sub
            End If
            
            Me.MousePointer = vbDefault
        Case vbKeyReturn
            If Trim(txtCodProveedor.Text) <> vbNullString Then
                'lblProveedor.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", Trim(txtCodProveedor.Text), "T")
                lblProveedor.Caption = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F2NOMPROV", "MAESTROS.EF2PROVEEDORES", "F2CODPROV", Trim(txtCodProveedor.Text), "T")
                
                If Trim(lblProveedor.Caption) = vbNullString Then
                    MsgBox "Proveedor no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    ModUtilitario.seleccionarTextoCaja txtCodProveedor
                    
                    Exit Sub
                Else
                    listarInventario
                End If
            End If
    End Select
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub


