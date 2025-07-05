VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmMantOrigen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Concepto de Movimiento de Almacen"
   ClientHeight    =   4470
   ClientLeft      =   5970
   ClientTop       =   1785
   ClientWidth     =   7680
   Icon            =   "frmMantOrigen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7680
   Begin ActiveToolBars.SSActiveToolBars tlbOrigen 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMantOrigen.frx":058A
      ToolBars        =   "frmMantOrigen.frx":382C
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7455
      Begin VB.ComboBox cmbCodAyudaProducto 
         Height          =   315
         ItemData        =   "frmMantOrigen.frx":390C
         Left            =   1440
         List            =   "frmMantOrigen.frx":3916
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3240
         Width           =   3615
      End
      Begin VB.CheckBox chkTieneAlmacenDestino 
         Caption         =   "Se debera asignar un Almacen de Destino."
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkRegistrarCosto 
         Caption         =   "Habilitar el registro de Costo."
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.ComboBox cmbTipoMovimiento 
         Height          =   315
         ItemData        =   "frmMantOrigen.frx":39C4
         Left            =   1440
         List            =   "frmMantOrigen.frx":39CE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtCodExterno 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkHabilitado 
         Caption         =   "Habilitado"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Para el Concepto mostrar la siguiente ayuda de Productos disponible:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Movimiento"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Externo"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMantOrigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolAyuda                As Boolean
Private bolNuevoOrigen       As Boolean
Private strCodigo               As String

Private objOrigen            As ClsOrigen


Public Property Let Ayuda(ByVal value As Boolean)
    bolAyuda = value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let NuevoOrigen(ByVal value As Boolean)
    bolNuevoOrigen = value
End Property

Public Property Get NuevoOrigen() As Boolean
    NuevoOrigen = bolNuevoOrigen
End Property

Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Private Sub limpiarCajas()
    txtCodigo.Text = vbNullString
    txtCodExterno.Text = vbNullString
    txtDescripcion.Text = vbNullString
    cmbTipoMovimiento.ListIndex = -1
    
    chkRegistrarCosto.value = vbUnchecked
    chkTieneAlmacenDestino.value = vbUnchecked
    chkHabilitado.value = vbChecked
    
    cmbCodAyudaProducto.ListIndex = 0
    
    txtCodigo.Locked = True
    txtCodigo.BackColor = DF
    cmbTipoMovimiento.Locked = False
    cmbTipoMovimiento.BackColor = HA
End Sub

Private Sub consultarOrigen()
    Set objOrigen = New ClsOrigen
    
    limpiarCajas
    
    With objOrigen
        .Codigo = strCodigo
        
        If .obtenerOrigen Then
            txtCodigo.Text = .Codigo
            txtCodExterno.Text = .CodigoExterno
            txtDescripcion.Text = .Descripcion
            
            cmbTipoMovimiento.ListIndex = ModUtilitario.seleccionarItem(cmbTipoMovimiento, .TipoMovimiento, "DER", 1)
            
            chkRegistrarCosto.value = IIf(.RegistrarCosto, vbChecked, vbUnchecked)
            chkTieneAlmacenDestino.value = IIf(.TieneAlmacenDestino, vbChecked, vbUnchecked)
            chkHabilitado.value = IIf(.Estado, vbChecked, vbUnchecked)
            
            cmbCodAyudaProducto.ListIndex = ModUtilitario.seleccionarItem(cmbCodAyudaProducto, .CodigoAyudaProducto, "DER", 1)
            
            cmbTipoMovimiento.Locked = True
            cmbTipoMovimiento.BackColor = DH
            
            If Not bolAyuda Then
                With frmListaOrigen
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        .listarOrigenSQL
                    Else
                        .listarOrigen
                    End If

                    
                End With
            End If
        End If
    End With
    
    Set objOrigen = Nothing
End Sub

Private Sub validarCajas()
    If Trim(txtCodigo.Text) = vbNullString Then
        MsgBox "El Campo Código es obligatorio.", vbCritical, App.ProductName
        
        txtCodigo.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtDescripcion.Text) = vbNullString Then
        MsgBox "El Campo Descripción es obligatorio.", vbCritical, App.ProductName
        
        txtCodigo.SetFocus
        
        Exit Sub
    End If
    
    If cmbTipoMovimiento.ListIndex = -1 Then
        MsgBox "El Campo Tipo de Movimiento es obligatorio.", vbCritical, App.ProductName
        
        cmbTipoMovimiento.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarOrigen
    End If
End Sub

Private Sub guardarOrigen()
    Set objOrigen = New ClsOrigen
    
    With objOrigen
        .Codigo = Trim(txtCodigo.Text)
        .CodigoExterno = Trim(txtCodExterno.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        
        .TipoMovimiento = right(cmbTipoMovimiento.Text, 1)
        
        .RegistrarCosto = IIf(chkRegistrarCosto.value = vbChecked, True, False)
        .TieneAlmacenDestino = IIf(chkTieneAlmacenDestino.value = vbChecked, True, False)
        .Estado = IIf(chkHabilitado.value = vbChecked, True, False)
        
        .CodigoAyudaProducto = right(cmbCodAyudaProducto.Text, 1)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarOrigen Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodigo = .Codigo
            
                consultarOrigen
            
            
            
            If Not bolAyuda Then
                MsgBox "Registro Actualizado.", vbInformation + vbOKOnly, App.ProductName
            Else
                objAyudaOrigen.Codigo = Trim(txtCodigo.Text)
                objAyudaOrigen.Descripcion = Trim(txtDescripcion.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objOrigen = Nothing
End Sub

Private Sub eliminarOrigen()
    Set objOrigen = New ClsOrigen
    
    With objOrigen
        .Codigo = Trim(txtCodigo.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Registro no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar el registro con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarOrigen Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodigo = .Codigo
                
                consultarOrigen
                
                MsgBox "Registro Eliminado.", _
                    vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objOrigen = Nothing
End Sub

Private Sub chkHabilitado_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub chkRegistrarCosto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub chkTieneAlmacenDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbCodAyudaProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbTipoMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub Form_Load()
    consultarOrigen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bolAyuda Then
        With frmListaOrigen
        
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                .listarOrigenSQL
            Else
                .listarOrigen
            End If
            
            
            '.Show
        End With
    End If
End Sub

Private Sub tlbOrigen_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            strCodigo = vbNullString
            
            consultarOrigen
        Case "Guardar"
            validarCajas
        Case "Eliminar"
            eliminarOrigen
        Case "Salir"
            objAyudaOrigen.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtCodExterno_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub





