VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmMantTipoExistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipo de Existencia"
   ClientHeight    =   6015
   ClientLeft      =   2085
   ClientTop       =   2115
   ClientWidth     =   9375
   Icon            =   "frmMantTipoExistencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9375
   Begin VB.Frame Frame3 
      Caption         =   " Configuracion Contable - Ventas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   9135
      Begin VB.TextBox txtCtaContableVta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCtaContableImpVta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Text            =   "Text3"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtAnexoVta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Text            =   "Text3"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtAnexoImpVta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Text            =   "Text3"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblCtaContableVta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Contable"
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCtaContableImpVta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   3600
         TabIndex        =   29
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Cont. Importación"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Anexo"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Anexo Importación"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Configuracion Contable - Compras "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   9135
      Begin VB.TextBox txtAnexoImp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtAnexo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtCtaContableImp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtCtaContable 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Anexo Importación"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Anexo"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Cont. Importación"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblCtaContableImp 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Contable"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCtaContable 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   5295
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbTipoExistencia 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMantTipoExistencia.frx":058A
      ToolBars        =   "frmMantTipoExistencia.frx":382C
   End
   Begin VB.Frame Frame1 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtAbreviatura 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1440
         Width           =   1215
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
         Alignment       =   1  'Right Justify
         Caption         =   "Habilitado"
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Abreviatura"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Externo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblDescripcion 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMantTipoExistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolAyuda                As Boolean
Private bolNuevoTipoExistencia       As Boolean
Private strCodigo               As String

Private objTipoExistencia            As ClsTipoExistencia

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let NuevoTipoExistencia(ByVal Value As Boolean)
    bolNuevoTipoExistencia = Value
End Property

Public Property Get NuevoTipoExistencia() As Boolean
    NuevoTipoExistencia = bolNuevoTipoExistencia
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Private Sub limpiarCajas()
    txtCodigo.Text = vbNullString
    txtCodExterno.Text = vbNullString
    chkHabilitado.Value = vbChecked
    txtDescripcion.Text = vbNullString
    txtAbreviatura.Text = vbNullString
    
    txtCtaContable.Text = vbNullString
        lblCtaContable.Caption = vbNullString: lblCtaContable.BackColor = DF
    txtAnexo.Text = vbNullString
    txtCtaContableImp.Text = vbNullString
        lblCtaContableImp.Caption = vbNullString: lblCtaContableImp.BackColor = DF
    txtAnexoImp.Text = vbNullString
    
    
    txtCtaContableVta.Text = vbNullString
        lblCtaContableVta.Caption = vbNullString: lblCtaContableVta.BackColor = DF
    txtAnexoVta.Text = vbNullString
    txtCtaContableImpVta.Text = vbNullString
        lblCtaContableImpVta.Caption = vbNullString: lblCtaContableImpVta.BackColor = DF
    txtAnexoImpVta.Text = vbNullString
    
    
    txtCodigo.Locked = True
    txtCodigo.BackColor = DF
End Sub

Private Sub consultarTipoExistencia()
    Set objTipoExistencia = New ClsTipoExistencia
    
    limpiarCajas
    
    With objTipoExistencia
        .Codigo = strCodigo
        
        If .obtenerTipoExistencia Then
            txtCodigo.Text = .Codigo
            txtCodExterno.Text = .CodigoExterno
            txtDescripcion.Text = .Descripcion
            txtAbreviatura.Text = .Abreviatura
            
            txtCtaContable.Text = .CtaContable
                lblCtaContable.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContable, "T")
                txtAnexo.Text = .Anexo
            
            txtCtaContableImp.Text = .CtaContableImportacion
                lblCtaContableImp.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableImportacion, "T")
                txtAnexoImp.Text = .AnexoImportacion
                
            txtCtaContableVta.Text = .CtaContableVta
                lblCtaContableVta.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableVta, "T")
                txtAnexoVta.Text = .AnexoVta
            
            txtCtaContableImpVta.Text = .CtaContableImportacionVta
                lblCtaContableImpVta.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableImportacionVta, "T")
                txtAnexoImpVta.Text = .AnexoImportacionVta
            
            chkHabilitado.Value = IIf(.Estado, vbChecked, vbUnchecked)
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DH
            
            If Not bolAyuda Then
                With frmListaTipoExistencia
                    If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                        With .dbgTipoExistencia
                            .DefaultFields = False
                            .Dataset.ADODataset.ConnectionString = cnBdCPlus
                            '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
                        End With
                        
                        .listarTipoExistenciaSQL
                    Else
                        With .dbgTipoExistencia
                            .DefaultFields = False
                            .Dataset.ADODataset.ConnectionString = cnn_dbbancos
                        End With
                        
                        .listarTipoExistencia
                    End If
                End With
            End If
        Else
            txtCodigo.Text = .generarCodigoTipoExistencia
        End If
    End With
    
    Set objTipoExistencia = Nothing
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
    
    If Trim(txtAbreviatura.Text) = vbNullString Then
        MsgBox "El Campo Abreviatura es obligatorio.", vbCritical, App.ProductName
        
        txtAbreviatura.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarTipoExistencia
    End If
End Sub

Private Sub guardarTipoExistencia()
    Set objTipoExistencia = New ClsTipoExistencia
    
    With objTipoExistencia
        .Codigo = Trim(txtCodigo.Text)
        .CodigoExterno = Trim(txtCodExterno.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        .Abreviatura = Trim(txtAbreviatura.Text)
        
        .CtaContable = Trim(txtCtaContable.Text)
        .Anexo = Trim(txtAnexo.Text)
        .CtaContableImportacion = Trim(txtCtaContableImp.Text)
        .AnexoImportacion = Trim(txtAnexoImp.Text)
        
        .CtaContableVta = Trim(txtCtaContableVta.Text)
        .AnexoVta = Trim(txtAnexoVta.Text)
        .CtaContableImportacionVta = Trim(txtCtaContableImpVta.Text)
        .AnexoImportacionVta = Trim(txtAnexoImpVta.Text)
        
        .Estado = IIf(chkHabilitado.Value = vbChecked, True, False)
        
        .UsuarioReg = wusuario
        .FechaReg = Format(Date, "Short Date")
        
        .UsuarioMod = wusuario
        .FechaMod = Format(Date, "Short Date")
        
        If .guardarTipoExistencia Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodigo = .Codigo
            
            consultarTipoExistencia
            
            If Not bolAyuda Then
                MsgBox "Registro Actualizado.", vbInformation + vbOKOnly, App.ProductName
            Else
                objAyudaTipoExistencia.Codigo = Trim(txtCodigo.Text)
                objAyudaTipoExistencia.Descripcion = Trim(txtDescripcion.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objTipoExistencia = Nothing
End Sub

Private Sub eliminarTipoExistencia()
    Set objTipoExistencia = New ClsTipoExistencia
    
    With objTipoExistencia
        .Codigo = Trim(txtCodigo.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Registro no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar el registro con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarTipoExistencia Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodigo = .Codigo
                
                consultarTipoExistencia
                
                MsgBox "Registro Eliminado.", _
                    vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objTipoExistencia = Nothing
End Sub

Private Sub Form_Load()
    abrirCnContaTabla
    
    consultarTipoExistencia
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bolAyuda Then
        With frmListaTipoExistencia
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                With .dbgTipoExistencia
                    .DefaultFields = False
                    .Dataset.ADODataset.ConnectionString = cnBdCPlus
                    '.Dataset.ADODataset.ConnectionString = "Provider=sqloledb;Data Source=.;Initial Catalog=BdCPlus;User Id=sa;Password=XXXXXX"
                End With
                
                .listarTipoExistenciaSQL
            Else
                With .dbgTipoExistencia
                    .DefaultFields = False
                    .Dataset.ADODataset.ConnectionString = cnn_dbbancos
                End With
                
                .listarTipoExistencia
            End If
            
            '.Show
        End With
    End If
End Sub

Private Sub tlbTipoExistencia_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            strCodigo = vbNullString
            
            consultarTipoExistencia
        Case "Guardar"
            validarCajas
        Case "Eliminar"
            eliminarTipoExistencia
        Case "Salir"
            objAyudaTipoExistencia.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtAbreviatura_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCtaContable_DblClick()
    txtCtaContable_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCtaContable_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_PlanCta
                wctacont = vbNullString
                
                .Show 1
                
                If wctacont <> vbNullString Then
                    txtCtaContable.Text = wctacont
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCtaContable_LostFocus()
    If Trim(txtCtaContable.Text) <> vbNullString Then
        lblCtaContable.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(txtCtaContable.Text), "T")
    Else
        txtCtaContable.Text = vbNullString
        lblCtaContable.Caption = vbNullString
    End If
End Sub

Private Sub txtCtaContable_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCtaContableImp_DblClick()
    txtCtaContableImp_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCtaContableImp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_PlanCta
                wctacont = vbNullString
                
                .Show 1
                
                If wctacont <> vbNullString Then
                    txtCtaContableImp.Text = wctacont
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCtaContableImpImp_LostFocus()
    If Trim(txtCtaContableImp.Text) <> vbNullString Then
        lblCtaContableImp.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(txtCtaContableImp.Text), "T")
    Else
        txtCtaContableImp.Text = vbNullString
        lblCtaContableImp.Caption = vbNullString
    End If
End Sub

Private Sub txtCtaContableImp_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub




Private Sub txtCtaContableVta_DblClick()
    txtCtaContableVta_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCtaContableVta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_PlanCta
                wctacont = vbNullString
                
                .Show 1
                
                If wctacont <> vbNullString Then
                    txtCtaContableVta.Text = wctacont
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCtaContableVta_LostFocus()
    If Trim(txtCtaContableVta.Text) <> vbNullString Then
        lblCtaContableVta.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(txtCtaContableVta.Text), "T")
    Else
        txtCtaContableVta.Text = vbNullString
        lblCtaContableVta.Caption = vbNullString
    End If
End Sub

Private Sub txtCtaContableVta_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCtaContableImpVta_DblClick()
    txtCtaContableImpVta_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCtaContableImpVta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_PlanCta
                wctacont = vbNullString
                
                .Show 1
                
                If wctacont <> vbNullString Then
                    txtCtaContableImpVta.Text = wctacont
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCtaContableImpImpVta_LostFocus()
    If Trim(txtCtaContableImpVta.Text) <> vbNullString Then
        lblCtaContableImpVta.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(txtCtaContableImpVta.Text), "T")
    Else
        txtCtaContableImpVta.Text = vbNullString
        lblCtaContableImpVta.Caption = vbNullString
    End If
End Sub

Private Sub txtCtaContableImpVta_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub




Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub



