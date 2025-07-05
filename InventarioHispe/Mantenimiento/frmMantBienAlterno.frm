VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmMantBienAlterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Producto Alterno"
   ClientHeight    =   2190
   ClientLeft      =   3555
   ClientTop       =   2670
   ClientWidth     =   9750
   Icon            =   "frmMantBienAlterno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   9750
   Begin ActiveToolBars.SSActiveToolBars tlbBienAlterno 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMantBienAlterno.frx":058A
      ToolBars        =   "frmMantBienAlterno.frx":382C
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txtCodAlterno 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkHabilitado 
         Caption         =   "Habilitado"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblAlterno 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label lblOriginal 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Producto Alterno"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Producto Original"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMantBienAlterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolAyuda                  As Boolean
Private bolNuevoBienAlterno       As Boolean
Private strCodigo                 As String
Private strCodigoAlterno          As String

Private objBienAlterno            As ClsBienAlterno


Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let NuevoBienAlterno(ByVal Value As Boolean)
    bolNuevoBienAlterno = Value
End Property

Public Property Get NuevoBienAlterno() As Boolean
    NuevoBienAlterno = bolNuevoBienAlterno
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let CodigoAlterno(ByVal Value As String)
    strCodigoAlterno = Value
End Property

Public Property Get CodigoAlterno() As String
    CodigoAlterno = strCodigoAlterno
End Property




Private Sub limpiarCajas()
    txtCodigo.Text = vbNullString
        lblOriginal.Caption = vbNullString
    txtCodAlterno.Text = vbNullString
        lblAlterno.Caption = vbNullString
    chkHabilitado.Value = vbChecked
    
    If strCodigo = vbNullString Then
        txtCodigo.Locked = True
        txtCodigo.BackColor = DF
    End If
End Sub

Private Sub consultarBienAlterno()
    Set objBienAlterno = New ClsBienAlterno
    
    limpiarCajas
    
    With objBienAlterno
        .CodigoBien = strCodigo
        .CodigoBienAlterno = strCodigoAlterno
        
        If .obtenerBienAlterno Then
            txtCodigo.Text = .CodigoBien
                lblOriginal.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", .CodigoBien, "T")
            txtCodAlterno.Text = .CodigoBienAlterno
                lblAlterno.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", .CodigoBienAlterno, "T")
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DF
            
            If Not bolAyuda Then
                With frmListaBienAlterno
                    .listarBienAlterno
                End With
            End If
        Else
            If bolAyuda Then
                txtCodigo.Text = strCodigo
                lblOriginal.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", strCodigo, "T")
                
                txtCodigo.Locked = True
                txtCodigo.BackColor = DF
            End If
        End If
    End With
    
    Set objBienAlterno = Nothing
End Sub

Private Sub validarCajas()
    If Trim(txtCodigo.Text) = vbNullString Then
        MsgBox "El Campo Producto Original es obligatorio.", vbCritical, App.ProductName
        
        txtCodigo.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtCodAlterno.Text) = vbNullString Then
        MsgBox "El Campo Producto Alterno es obligatorio.", vbCritical, App.ProductName
        
        txtCodAlterno.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarBienAlterno
    End If
End Sub

Private Sub guardarBienAlterno()
    Set objBienAlterno = New ClsBienAlterno
    
    With objBienAlterno
        .CodigoBien = Trim(txtCodigo.Text)
        .CodigoBienAlterno = Trim(txtCodAlterno.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarBienAlterno Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodigo = .CodigoBien
            strCodigoAlterno = .CodigoBienAlterno
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                guardarBienAlternoSql
            End If
            
            consultarBienAlterno
            
            If Not bolAyuda Then
                MsgBox "Color Actualizado.", _
                        vbInformation, App.ProductName
            Else
                objAyudaBienAlterno.CodigoBien = Trim(txtCodigo.Text)
                objAyudaBienAlterno.CodigoBienAlterno = Trim(txtCodAlterno.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objBienAlterno = Nothing
End Sub

Private Sub eliminarBienAlterno()
    Set objBienAlterno = New ClsBienAlterno
    
    With objBienAlterno
        .CodigoBien = Trim(txtCodigo.Text)
        .CodigoBienAlterno = Trim(txtCodAlterno.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Registro no existente.", vbInformation, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar la relación entre productos actual?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarBienAlterno Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodigo = .CodigoBien
                strCodigoAlterno = .CodigoBienAlterno
                
                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                    eliminarBienAlternoSql
                End If
                
                consultarBienAlterno
                
                MsgBox "Registro Eliminado.", _
                    vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objBienAlterno = Nothing
End Sub

Private Sub Form_Load()
    consultarBienAlterno
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bolAyuda Then
        With frmListaBienAlterno
            .listarBienAlterno
            
            '.Show
        End With
    End If
End Sub

Private Sub tlbBienAlterno_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            strCodigo = vbNullString
            strCodigoAlterno = vbNullString
            
            consultarBienAlterno
        Case "Guardar"
            validarCajas
        Case "Eliminar"
            eliminarBienAlterno
        Case "Salir"
            objAyudaBienAlterno.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtCodigo_DblClick()
    txtCodigo_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                
                .Ayuda = True
                .InsumoOP = False
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                .Show vbModal
                
                If objAyudaBien.Codigo <> vbNullString Then
                    txtCodigo.Text = objAyudaBien.Codigo
                    lblOriginal.Caption = objAyudaBien.Descripcion
                Else
                    txtCodigo.Text = vbNullString
                    lblOriginal.Caption = vbNullString
                End If
            End With
    End Select
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtCodAlterno_DblClick()
    txtCodAlterno_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodAlterno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ModUtilitario.validarFormAbierto("frmListaBien") Then
                Unload frmListaBien
            End If
            
            With frmListaBien
                objAyudaBien.inicializarEntidades
                
                '.Ayuda = True
                '.TieneMovimientoAlmacen = True
                
                .Ayuda = True
                .InsumoOP = False
                .ParaVenta = False
                .TieneMovimientoAlmacen = True
                .CadenaCorte = vbNullString
                .FiltroAdicional = vbNullString
                .TipoBienMostrar = "P"
                
                .Show vbModal
                
                If objAyudaBien.Codigo <> vbNullString Then
                    txtCodAlterno.Text = objAyudaBien.Codigo
                    lblAlterno.Caption = objAyudaBien.Descripcion
                Else
                    txtCodAlterno.Text = vbNullString
                    lblAlterno.Caption = vbNullString
                End If
            End With
    End Select
End Sub

Private Sub txtCodAlterno_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub




''--------------------------------------------------------------------------------------------------------------
''------------ SQL ---------------------------------------------------------------------------------------------
''--------------------------------------------------------------------------------------------------------------
Private Sub consultarBienAlternoSql()
    Set objSqlBienAlterno = New SqlClsBienAlterno
    
    limpiarCajas
    
    With objSqlBienAlterno
        .CodigoBien = strCodigo
        .CodigoBienAlterno = strCodigoAlterno
        
        If .obtenerBienAlterno Then
            txtCodigo.Text = .CodigoBien
                lblOriginal.Caption = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F5NOMPRO", "MAESTROS.IF5PLA", "F5CODPRO", .CodigoBien, "T")
            txtCodAlterno.Text = .CodigoBienAlterno
                lblAlterno.Caption = ModUtilitario.ObtenerCampoV2(cnBdCPlus, "F5NOMPRO", "MAESTROS.IF5PLA", "F5CODPRO", .CodigoBienAlterno, "T")
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DF
            
            If Not bolAyuda Then
                With frmListaBienAlterno
                    .listarBienAlterno
                End With
            End If
        Else
            If bolAyuda Then
                txtCodigo.Text = strCodigo
                lblOriginal.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", strCodigo, "T")
                
                txtCodigo.Locked = True
                txtCodigo.BackColor = DF
            End If
        End If
    End With
    
    Set objSqlBienAlterno = Nothing
End Sub

Private Sub guardarBienAlternoSql()
    Set objSqlBienAlterno = New SqlClsBienAlterno
    
    With objSqlBienAlterno
        .CodigoBien = strCodigo  'Trim(txtCodigo.Text)
        .CodigoBienAlterno = strCodigoAlterno 'Trim(txtCodAlterno.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarBienAlterno Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
'            strCodigo = .CodigoBien
'            strCodigoAlterno = .CodigoBienAlterno
'
'            consultarBienAlterno
'
'            If Not bolAyuda Then
'                MsgBox "Color Actualizado.", _
'                        vbInformation, App.ProductName
'            Else
'                objAyudaBienAlterno.CodigoBien = Trim(txtCodigo.Text)
'                objAyudaBienAlterno.CodigoBienAlterno = Trim(txtCodAlterno.Text)
'
'                Unload Me
'            End If
        End If
    End With
    
    Set objSqlBienAlterno = Nothing
End Sub

Private Sub eliminarBienAlternoSql()
    Set objSqlBienAlterno = New SqlClsBienAlterno
    
    With objSqlBienAlterno
        .CodigoBien = strCodigo  'Trim(txtCodigo.Text)
        .CodigoBienAlterno = strCodigoAlterno  'Trim(txtCodAlterno.Text)
'
'        If Not .verificarExistencia Then
'            MsgBox "Registro no existente.", vbInformation, App.ProductName
'
'            Exit Sub
'        End If
'
'        If MsgBox("¿Desea Eliminar la relación entre productos actual?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarBienAlterno Then
                Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
'
'                strCodigo = .CodigoBien
'                strCodigoAlterno = .CodigoBienAlterno
'
'                consultarBienAlterno
'
'                MsgBox "Registro Eliminado.", _
'                    vbInformation, App.ProductName
            End If
'        End If
    End With
    
    Set objSqlBienAlterno = Nothing
End Sub
