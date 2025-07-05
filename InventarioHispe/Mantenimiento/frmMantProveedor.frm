VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Proveedor"
   ClientHeight    =   8625
   ClientLeft      =   225
   ClientTop       =   1740
   ClientWidth     =   7935
   ControlBox      =   0   'False
   Icon            =   "frmMantProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   7935
   Begin VB.Frame fraPredeteminada 
      Caption         =   " Configuraciones Predeterminadas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   7695
      Begin VB.TextBox txtCuentaContable 
         Height          =   285
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox txtGrupoResultados 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtGrupoFlujo 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2160
         Width           =   885
      End
      Begin VB.CheckBox chkEsAptoParaOrden 
         Alignment       =   1  'Right Justify
         Caption         =   "Orden de Compra / Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   1800
         Width           =   2775
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         ItemData        =   "frmMantProveedor.frx":058A
         Left            =   1560
         List            =   "frmMantProveedor.frx":0597
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox cmbTipoComprobante 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1080
         Width           =   3015
      End
      Begin VB.ComboBox cmbCategoria 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtCodigoGasto 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   14
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   4920
         TabIndex        =   50
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGrupoResultados 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   2520
         TabIndex        =   48
         Top             =   2520
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo Resultados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   47
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblGrupoFlujo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   2520
         TabIndex        =   46
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo Flujo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   45
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblCodigoGasto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   2520
         TabIndex        =   44
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo de Gasto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Comprobante"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Frame fraAdicional 
      Caption         =   " Información Adicional "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   7695
      Begin VB.TextBox txtCtaAbono 
         Height          =   285
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   12
         Top             =   1080
         Width           =   5925
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   5925
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   2805
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   10
         Top             =   360
         Width           =   1965
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   9
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cuenta de Abono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   195
         TabIndex        =   37
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefono(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   " Información Principal "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   7695
      Begin VB.ComboBox cmbDistrito 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtCodigoExterno 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1440
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1440
      End
      Begin VB.TextBox txtDireccionFiscal 
         Height          =   645
         Left            =   1560
         MaxLength       =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1800
         Width           =   5925
      End
      Begin VB.ComboBox cmbTipoDocumento 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbPersona 
         Height          =   315
         ItemData        =   "frmMantProveedor.frx":05B2
         Left            =   2760
         List            =   "frmMantProveedor.frx":05BC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNroDocumento 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox txtRazonSocial 
         Height          =   285
         Left            =   1560
         MaxLength       =   150
         TabIndex        =   5
         Top             =   1440
         Width           =   5925
      End
      Begin VB.ComboBox cmbOrigen 
         Height          =   315
         ItemData        =   "frmMantProveedor.frx":069D
         Left            =   360
         List            =   "frmMantProveedor.frx":06A7
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Externo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   4680
         TabIndex        =   51
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Interno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Postal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   30
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Razón Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "No.RUC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Tipo de Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   26
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label 
         Caption         =   "Persona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   25
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label 
         Caption         =   "Origen de Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComctlLib.Toolbar tlbProveedor 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
      ButtonWidth     =   2117
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imglstCliente"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            Object.ToolTipText     =   "Nuevo (Ctrl + N)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
            Object.ToolTipText     =   "Guardar (Ctrl + G)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
            Object.ToolTipText     =   "Eliminar  (Ctrl + E)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "C&ontacto"
            Object.ToolTipText     =   "Contacto (Ctrl + O)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Salir (Ctrl + S)"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList imglstCliente 
         Left            =   9600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantProveedor.frx":078B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantProveedor.frx":0D25
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantProveedor.frx":12BF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantProveedor.frx":1859
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantProveedor.frx":1DF3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMantProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda        As Boolean
Private strCodProveedor          As String

Private objProveedor             As ClsProveedor
'Private objSqlProveedor          As SqlClsProveedor


Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodProveedor = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodProveedor
End Property

Private Sub listarTipoDocumento()
    If cmbOrigen.ListIndex <> -1 And cmbPersona.ListIndex <> -1 Then
        With objAyudaTipoDocID
            .Origen = right(cmbOrigen.Text, 1)
            .Persona = right(cmbPersona.Text, 1)
            
            .listarTipoDocumento cmbTipoDocumento
        End With
        
        If cmbTipoDocumento.ListCount > 0 Then
            cmbTipoDocumento.ListIndex = -1
        End If
    End If
End Sub

Private Sub listarDistritos()
    objAyudaDistrito.listarDistrito cmbDistrito, True
    
    If cmbDistrito.ListCount > 0 Then
        cmbDistrito.ListIndex = ModUtilitario.seleccionarItem(cmbDistrito, "01", "DER", 2)
    End If
End Sub

Private Sub listarCategoria()
    objAyudaCategoria.listarCategoria cmbCategoria
    
    If cmbCategoria.ListCount > 0 Then
        cmbCategoria.ListIndex = 0
    End If
End Sub

Private Sub listarTipoComprobante()
    objAyudaComprobante.listarTipoComprobante cmbTipoComprobante, "'P', 'A'"
    
    If cmbTipoComprobante.ListCount > 0 Then
        cmbTipoComprobante.ListIndex = ModUtilitario.seleccionarItem(cmbTipoComprobante, "01", "DER", 2)
    End If
End Sub

Private Sub listarFormaPago()
    objAyudaFormaPago.listarFormaPago cmbFormaPago
    
    If cmbFormaPago.ListIndex > 0 Then
        cmbFormaPago.ListIndex = ModUtilitario.seleccionarItem(cmbFormaPago, "001", "DER", 3)
    End If
End Sub

Private Sub limpiarCajas()
    cmbOrigen.ListIndex = 0
    cmbPersona.ListIndex = 1
    If cmbTipoDocumento.ListCount > 0 Then cmbTipoDocumento.ListIndex = -1
    
    txtNroDocumento.Text = vbNullString
    txtCodigo.Text = vbNullString
    txtRazonSocial.Text = vbNullString
    txtDireccionFiscal.Text = vbNullString
    If cmbDistrito.ListCount > 0 Then cmbDistrito.ListIndex = ModUtilitario.seleccionarItem(cmbDistrito, "01", "DER", 2)
    txtCodigoExterno.Text = vbNullString
    
    txtTelefono.Text = vbNullString
    txtFax.Text = vbNullString
    txtEmail.Text = vbNullString
    txtCtaAbono.Text = vbNullString
    txtContacto.Text = vbNullString
    
    cmbMoneda.ListIndex = ModUtilitario.seleccionarItem(cmbMoneda, ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l"), "IZQ", 1)
    
    If cmbTipoComprobante.ListCount > 0 Then cmbTipoComprobante.ListIndex = -1
    
    If cmbFormaPago.ListCount > 0 Then cmbFormaPago.ListIndex = -1
    
    If cmbCategoria.ListCount > 0 Then cmbCategoria.ListIndex = -1
    
    txtCodigoGasto.Text = vbNullString: lblCodigoGasto.Caption = vbNullString
    txtGrupoFlujo.Text = vbNullString: lblGrupoFlujo.Caption = vbNullString
    txtGrupoResultados.Text = vbNullString: lblGrupoResultados.Caption = vbNullString
    
    
    txtCodigo.Locked = True
    txtCodigo.BackColor = DF
    txtNroDocumento.Locked = False
    
    tlbProveedor.buttons(4).Enabled = False
    tlbProveedor.buttons(4).Visible = False
    tlbProveedor.buttons(6).Enabled = False
    tlbProveedor.buttons(6).Visible = False
End Sub

Private Sub consultarProveedor()
    Set objProveedor = New ClsProveedor
    
    limpiarCajas
    
    With objProveedor
        .inicializarEntidades
        
        .Codigo = strCodProveedor
        
        If .obtenerProveedor Then
            cmbOrigen.ListIndex = ModUtilitario.seleccionarItem(cmbOrigen, .OrigenProveedor, "DER", 1)
            cmbPersona.ListIndex = ModUtilitario.seleccionarItem(cmbPersona, .ClaseProveedor, "DER", 1)
            cmbTipoDocumento.ListIndex = ModUtilitario.seleccionarItem(cmbTipoDocumento, .CodigoTipoDocumento, "DER", 1)
            
            txtNroDocumento.Text = .NumeroDocumento
            txtCodigo.Text = .Codigo
            txtRazonSocial.Text = .NombreProveedor
            txtDireccionFiscal.Text = .DireccionProveedor
            cmbDistrito.ListIndex = ModUtilitario.seleccionarItem(cmbDistrito, .CodigoDistrito, "DER", 2)
            txtCodigoExterno.Text = .CodigoExterno
            
            txtTelefono.Text = .Telefono
            txtFax.Text = .Fax
            txtEmail.Text = .Email
            txtCtaAbono.Text = .CuentaAbono
            txtContacto.Text = .Contacto
            
            cmbMoneda.ListIndex = ModUtilitario.seleccionarItem(cmbMoneda, .CodigoMoneda, "IZQ", 1)
            cmbTipoComprobante.ListIndex = ModUtilitario.seleccionarItem(cmbTipoComprobante, .CodTipoComprobante, "DER", 2)
            cmbFormaPago.ListIndex = ModUtilitario.seleccionarItem(cmbFormaPago, .CodigoFormaPago, "DER", 3)
            cmbCategoria.ListIndex = ModUtilitario.seleccionarItem(cmbCategoria, "*" & .CodigoCategoria, "DER", Len(Trim(.CodigoCategoria & "")) + 1)
            
            txtCodigoGasto.Text = .CodigoGasto
                lblCodigoGasto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "BF9GIN", "CODIGO", .CodigoGasto, "T")
            
            txtGrupoFlujo.Text = .GrupoFlujo
                lblGrupoFlujo.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "GRUPOS_FLUJO", "CODIGO", .GrupoFlujo, "T")
                
            txtGrupoResultados.Text = .GrupoFlujo
                lblGrupoResultados.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "GRUPORESULTADO", "CODIGO", .GrupoFlujo, "T")
                
            chkEsAptoParaOrden.Value = IIf(.EsAptoParaOrden, vbChecked, vbUnchecked)
            
            txtNroDocumento.Locked = False
            
            If Not bolAyuda Then
                tlbProveedor.buttons(4).Enabled = True
                tlbProveedor.buttons(4).Visible = True
                tlbProveedor.buttons(6).Enabled = True
                tlbProveedor.buttons(6).Visible = True
            End If
        End If
    End With
    
    Set objProveedor = Nothing
End Sub

Private Sub validarCajas()
    If cmbOrigen.ListIndex = -1 Then
        MsgBox "El Campo Origen es obligatorio.", vbCritical + vbOKOnly, App.ProductName
            
        cmbOrigen.SetFocus
        
        Exit Sub
    End If
    
    If cmbPersona.ListIndex = -1 Then
        MsgBox "El Campo Persona es obligatorio.", vbCritical + vbOKOnly, App.ProductName
            
        cmbPersona.SetFocus
        
        Exit Sub
    End If
    
    If cmbTipoDocumento.ListIndex = -1 Then
        MsgBox "El Campo Tipo de Documento de Identidad es obligatorio.", vbCritical + vbOKOnly, App.ProductName
            
        cmbTipoDocumento.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtNroDocumento.Text) = vbNullString Then
        MsgBox "El Campo Numero de Documento de Identidad es obligatorio.", vbCritical + vbOKOnly, App.ProductName
            
        txtNroDocumento.SetFocus
        
        Exit Sub
    Else
        With objAyudaTipoDocID
            .inicializarEntidades
            
            .Codigo = right(cmbTipoDocumento.Text, 1)
            
            .obtenerConfigTipoDocumento
            
            If .TieneLargoFijo Then
                If Len(Trim(txtNroDocumento.Text)) < txtNroDocumento.MaxLength Then
                    MsgBox "La longitud del Numero de Documento de Identidad es incorrecto, verifique.", vbCritical + vbOKOnly, App.ProductName
                    
                    txtNroDocumento.SetFocus
                    
                    Exit Sub
                End If
            End If
        End With
    End If
    
    If Not txtCodigo.Locked Then
        If Trim(txtCodigo.Text) = vbNullString Then
            MsgBox "El Campo Código es obligatorio.", vbCritical + vbOKOnly, App.ProductName
            
            txtCodigo.SetFocus
            
            Exit Sub
        End If
    End If
    
    If Trim(txtRazonSocial.Text) = vbNullString Then
        MsgBox "El Campo Razon Social es obligatorio.", vbCritical + vbOKOnly, App.ProductName
        
        txtRazonSocial.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtDireccionFiscal.Text) = vbNullString Then
        MsgBox "El Campo Dirección Fiscal es obligatorio.", vbCritical + vbOKOnly, App.ProductName
        
        txtDireccionFiscal.SetFocus
        
        Exit Sub
    End If
    
    If cmbTipoComprobante.ListIndex = -1 Then
        MsgBox "El Campo Tipo de Comprobante es obligatorio.", vbCritical + vbOKOnly, App.ProductName
        
        cmbTipoComprobante.SetFocus
        
        Exit Sub
    End If
    
    If cmbFormaPago.ListIndex = -1 Then
        MsgBox "El Campo Forma de Pago es obligatorio.", vbCritical + vbOKOnly, App.ProductName
        
        cmbFormaPago.SetFocus
        
        Exit Sub
    End If
    
    If cmbCategoria.ListIndex = -1 Then
        MsgBox "El Campo Categoria es obligatorio.", vbCritical + vbOKOnly, App.ProductName
        
        cmbCategoria.SetFocus
        
        Exit Sub
    End If
    
    If txtCodigoGasto.Text = vbNullString Then
        If MsgBox("Se sugiere consignar el Codigo de Gasto." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        
            txtCodigoGasto.SetFocus
            
            Exit Sub
        End If
    End If
    
    If CBool(chkEsAptoParaOrden.Value) And Trim(txtEmail.Text) = vbNullString Then
        If MsgBox("Si se marca al Proveedor como apto para Orden, se sugiere consignar el E-mail." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        
            txtEmail.SetFocus
            
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Desea guardar los datos del Proveedor?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarProveedor
    End If
End Sub

Private Sub guardarProveedor()
    Set objProveedor = New ClsProveedor
    
    With objProveedor
        .inicializarEntidades
        
        .OrigenProveedor = right(cmbOrigen.Text, 1)
        .ClaseProveedor = right(cmbPersona.Text, 1)
        .CodigoTipoDocumento = right(cmbTipoDocumento.Text, 1)
        
        .NumeroDocumento = Trim(txtNroDocumento.Text)
        .Codigo = Trim(txtCodigo.Text)
        .NombreProveedor = Replace(Trim(txtRazonSocial.Text), "'", "' & Chr(39) & '", 1)
        .DireccionProveedor = Replace(Trim(txtDireccionFiscal.Text), "'", "' & Chr(39) & '", 1)
        .CodigoDistrito = right(cmbDistrito.Text, 2)
        .CodigoExterno = Trim(txtCodigoExterno.Text)
        
        .Telefono = Trim(txtTelefono.Text)
        .Fax = Trim(txtFax.Text)
        .Email = Trim(txtEmail.Text)
        .CuentaAbono = Trim(txtCtaAbono.Text)
        .Contacto = Replace(Trim(txtContacto.Text), "'", "' & Chr(39) & '", 1)
        
        .CodigoMoneda = UCase(left(cmbMoneda.Text, 1))
        .CodTipoComprobante = right(cmbTipoComprobante.Text, 2)
        .CodigoFormaPago = right(cmbFormaPago.Text, 3)
        .CodigoCategoria = Mid(cmbCategoria.Text, InStr(1, cmbCategoria.Text, "*") + 1)
        
        .CuentaContable = Trim(txtCuentaContable.Text)
        .CodigoGasto = Trim(txtCodigoGasto.Text)
        .GrupoFlujo = Trim(txtGrupoFlujo.Text)
        .GrupoResultado = Trim(txtGrupoResultados.Text)
        
        .EsAptoParaOrden = CBool(chkEsAptoParaOrden.Value)
        
        If Not txtNroDocumento.Locked And Trim(txtCodigo.Text) = vbNullString Then
            If .verificarExistenciaPorNroDocumento Then
                MsgBox "Numero de Documento ingresado ya existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                txtNroDocumento.SetFocus
                
                ModUtilitario.seleccionarTextoCaja txtNroDocumento
            End If
        End If
        
        .FechaReg = Format(Date, "Short Date")
        .UsuarioReg = wusuario
        .FechaMod = Format(Date, "Short Date")
        .UsuarioMod = wusuario
        
        If .guardarProveedor Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
'            If ModMilano.exportarProveedorAserverSQL(.Codigo) Then
'
'            End If
            
            strCodProveedor = .Codigo
            
'            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                guardarProveedorSql
'            End If
            
'            If ConsultarProveedorMySQL(Trim("" & txtNroDocumento.Text)) = True Then
'                If AddProveedorMySQL(Trim("" & txtNroDocumento.Text), Trim("" & txtRazonSocial.Text), Trim("" & txtNroDocumento.Text)) = True Then
'                    'MsgBox "El proveedor ha sido registrado", vbInformation, wnomcia
'                End If
'            End If
            
'            If AddProveedorMySQL(Trim("" & txtNroDocumento.Text), Trim("" & txtRazonSocial.Text), Trim("" & txtNroDocumento.Text)) = True Then
'                'MsgBox "El proveedor ha sido registrado", vbInformation, wnomcia
'            End If
    
            consultarProveedor
            
            If Not bolAyuda Then
                MsgBox "Proveedor Actualizado.", _
                        vbInformation, App.ProductName
            Else
                objAyudaProveedor.Codigo = Trim(txtCodigo.Text)
                objAyudaProveedor.NombreProveedor = Trim(txtRazonSocial.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objProveedor = Nothing
End Sub

Private Sub eliminarProveedor()
    Set objProveedor = New ClsProveedor
    
    With objProveedor
        .Codigo = Trim(txtCodigo.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Proveedor no existente.", vbInformation, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar la Proveedor con Codigo '" & .Codigo & "'?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarProveedor Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodProveedor = .Codigo
                
'                If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
'                    eliminarProveedorSql
'                End If
                
                consultarProveedor
                
                MsgBox "Proveedor Eliminado.", _
                        vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objProveedor = Nothing
End Sub

Private Function validarShortCut(ByVal Key As Integer) As Boolean
    validarShortCut = True
    
    Select Case Key
        Case 14 'Ctrl + N (Nuevo)
            tlbProveedor_ButtonClick tlbProveedor.buttons(2)
        Case 7 'Ctrl + G (Guardar)
            tlbProveedor_ButtonClick tlbProveedor.buttons(3)
        Case 5 'Ctrl + E (Eliminar)
            tlbProveedor_ButtonClick tlbProveedor.buttons(4)
        Case 15 'Ctrl + O (Contacto)
            tlbProveedor_ButtonClick tlbProveedor.buttons(6)
        Case 19 'Ctrl + S (Salir)
            tlbProveedor_ButtonClick tlbProveedor.buttons(8)
        Case Else
            validarShortCut = False
    End Select
End Function

Private Sub Form_Load()
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    'Llenar Distritos
    listarDistritos
    'Llenar Tipos de Comprobantes
    listarTipoComprobante
    'Llenar Formas de Pago
    listarFormaPago
    'Llenar Categoria de Proveedor
    listarCategoria
    
    consultarProveedor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If Not bolAyuda Then
'        With frmListaProveedor
'            .listarProveedor
'        End With
'    End If
End Sub

Private Sub tlbProveedor_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Trim(Button.Caption)
        Case "&Nuevo"
            If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                strCodProveedor = vbNullString
                
                consultarProveedor
            End If
        Case "&Guardar"
            validarCajas
        Case "&Eliminar"
            eliminarProveedor
        Case "&Salir"
            Unload Me
    End Select
End Sub

Private Sub cmbOrigen_Click()
    listarTipoDocumento
End Sub

Private Sub cmbOrigen_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbPersona_Click()
    listarTipoDocumento
End Sub

Private Sub cmbPersona_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbTipoDocumento_Click()
    With objAyudaTipoDocID
        .inicializarEntidades
        
        .Codigo = right(cmbTipoDocumento.Text, 1)
        
        .obtenerConfigTipoDocumento

        txtNroDocumento.MaxLength = .Longitud
        
        If Len(Trim(txtNroDocumento.Text)) > .Longitud Then
            txtNroDocumento.Text = vbNullString
        End If
    End With
End Sub

Private Sub cmbTipoDocumento_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbTipoDocumento_LostFocus()
    cmbTipoDocumento_Click
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    'If Not txtNroDocumento.Locked Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyReturn
            If cmbTipoDocumento.ListIndex = -1 Then MsgBox "Seleccione el Tipo de Documento.", vbInformation + vbOKOnly, App.ProductName
            
            If right(cmbTipoDocumento.Text, 1) = "6" Then
                If MsgBox("¿Desea verificar el Numero de Documento en Linea?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                    Me.MousePointer = vbHourglass
                    
                    With objAyudaProveedor
                        .inicializarEntidades
                        

                        
                        If ModUtilitario.validarRUCenFactiliza(Trim(txtNroDocumento.Text), True) Then
                            If srazonSocial <> vbNullString Then
                                txtRazonSocial.Text = srazonSocial 'ModUtilitario.limpiarCaracteresEnCadena(objAyudaCliente.NombreCliente)
                                
                                If sdireccion <> vbNullString Then
                                    txtDireccionFiscal.Text = sdireccion 'ModUtilitario.limpiarCaracteresEnCadena(objAyudaCliente.DireccionCliente)
                                End If
                                
                                If objAyudaCliente.Telefono <> vbNullString Then
                                    txtTelefono.Text = stelefono 'objAyudaCliente.Telefono
                                End If
                                
                                If objAyudaCliente.Fax <> vbNullString Then
                                    txtFax.Text = "" 'objAyudaCliente.Fax
                                End If
                                
                                MsgBox "Verificación de Proveedor exitosa." & vbNewLine & _
                                        "Condición: " & scondicion, vbInformation + vbOKOnly, App.ProductName
                            
                                .inicializarEntidades
    
                                ModUtilitario.pulsarTecla vbKeyTab
                            Else
                                MsgBox "Número de Documento inválido o Proveedor no existe.", vbInformation + vbOKOnly, App.ProductName
                                
                                .inicializarEntidades
                            End If
                        End If
                    End With
                    
                    Me.MousePointer = vbDefault
                Else
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End If
    End Select
End Sub

Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
    'If Not txtNroDocumento.Locked Then
        If Not validarShortCut(KeyAscii) Then
            'Validar Cadena segun el Tipo de Documento Seleccionado
            With objAyudaTipoDocID
                .inicializarEntidades
                
                .Codigo = right(cmbTipoDocumento.Text, 1)
                
                .obtenerConfigTipoDocumento
                
                Select Case .TipoCadena
                    Case "N"
                        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
                    Case "A"
                        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
                End Select
            End With
        End If
    'End If
End Sub

Private Sub txtNroDocumento_LostFocus()
    On Error GoTo errNroDocumento
    
    If Not txtNroDocumento.Locked And Trim(txtCodigo.Text) = vbNullString Then
        With objAyudaProveedor
            .NumeroDocumento = Trim(txtNroDocumento.Text)
            
            If .verificarExistenciaPorNroDocumento Then
                MsgBox "Numero de Documento ingresado ya existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                txtNroDocumento.SetFocus
                
                ModUtilitario.seleccionarTextoCaja txtNroDocumento
            Else
                If right(cmbTipoDocumento.Text, 1) = "6" Then
'                    .inicializarEntidades
'
'                    If ModUtilitario.validarRUCenSunat(Trim(txtNroDocumento.Text), True) Then
'                        txtRazonSocial.Text = ModUtilitario.limpiarCaracteresEnCadena(.NombreProveedor)
'                        txtDireccionFiscal.Text = ModUtilitario.limpiarCaracteresEnCadena(.DireccionProveedor)
'                        txtTelefono.Text = .Telefono
'
'                        MsgBox "Verificación de Proveedor exitosa." & vbNewLine & _
'                                "Estado: " & .CodigoVendedor & vbNewLine & _
'                                "Situación: " & .CodigoCobrador, vbInformation + vbOKOnly, App.ProductName
'
'                        .inicializarEntidades
'                    End If
                End If
            End If
        End With
    End If
    
    Exit Sub
errNroDocumento:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            MsgBox "No.: " & Err.Number & vbNewLine & "Descripcion: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    End Select
    
    Err.Clear
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtDireccionFiscal_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub cmbDistrito_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub txtCodigoExterno_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaTextoSinMayus(KeyAscii)
    End If
End Sub

Private Sub txtweb_keypress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaTextoSinMayus(KeyAscii)
    End If
End Sub

Private Sub txtContacto_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub cmbmoneda_keypress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbTipoComprobante_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbFormaPago_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbCategoria_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub txtCodigoGasto_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtCodigoGasto_DblClick()
    txtCodigoGasto_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodigoGasto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_gastos
                wdestino = "E"
                wgastos = vbNullString
                
                .TipoConcepto = wdestino
                
                .Show 1
                
                If wgastos <> vbNullString Then
                    txtCodigoGasto.Text = wgastos
                    lblCodigoGasto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "BF9GIN", "CODIGO", wgastos, "T")
                    txtCuentaContable.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CUENTA", "BF9GIN", "CODIGO", wgastos, "T")
                    
                    With objAyudaCuentaContable
                        .inicializarEntidades
                        
                        .Codigo = Trim(txtCuentaContable.Text)
                        
                        .obtenerConfigCuentaContable
                        
                        If .CodMoneda <> vbNullString Then
                            cmbMoneda.ListIndex = ModUtilitario.seleccionarItem(cmbMoneda, .CodMoneda, "IZQ", 1)
                            
                            cmbMoneda.Enabled = False
                        Else
                            cmbMoneda.ListIndex = -1
                            
                            cmbMoneda.Enabled = True
                        End If
                    End With
                    
                    Unload ayuda_gastos
                End If
            End With
    End Select
End Sub

Private Sub txtCodigoGasto_LostFocus()
    If Trim(txtCodigoGasto.Text) <> vbNullString Then
        txtCodigoGasto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CODIGO", Trim(txtCodigoGasto.Text), "T")
        lblCodigoGasto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "BF9GIN", "CODIGO", Trim(txtCodigoGasto.Text), "T")
        txtCuentaContable.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CUENTA", "BF9GIN", "CODIGO", Trim(txtCodigoGasto.Text), "T")
        
        With objAyudaCuentaContable
            .inicializarEntidades
            
            .Codigo = Trim(txtCuentaContable.Text)
            
            .obtenerConfigCuentaContable
            
            If .CodMoneda <> vbNullString Then
                cmbMoneda.ListIndex = ModUtilitario.seleccionarItem(cmbMoneda, .CodMoneda, "IZQ", 1)
                
                cmbMoneda.Enabled = False
            Else
                cmbMoneda.ListIndex = -1
                
                cmbMoneda.Enabled = True
            End If
        End With
    Else
        lblCodigoGasto.Caption = vbNullString
        txtCuentaContable.Text = vbNullString
        cmbMoneda.ListIndex = -1
                
        cmbMoneda.Enabled = True
    End If
End Sub

Private Sub txtGrupoFlujo_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtGrupoFlujo_DblClick()
    txtGrupoFlujo_KeyDown vbKeyF2, 0
End Sub

Private Sub txtGrupoFlujo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_grupoflujo
                sw_ayuda = True
                wdestino = "P"
                wcodgrupo = vbNullString
                
                .Show 1
                
                sw_ayuda = False
                
                If wcodgrupo <> vbNullString Then
                    txtGrupoFlujo.Text = wcodgrupo
                    lblGrupoFlujo.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "GRUPOS_FLUJO", "CODIGO", wcodgrupo, "T")
                    
                    Unload ayuda_gastos
                End If
            End With
    End Select
End Sub

Private Sub txtGrupoFlujo_LostFocus()
    If Trim(txtGrupoFlujo.Text) <> vbNullString Then
        txtGrupoFlujo.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "GRUPOS_FLUJO", "CODIGO", Trim(txtGrupoFlujo.Text), "T")
        lblGrupoFlujo.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "GRUPOS_FLUJO", "CODIGO", Trim(txtGrupoFlujo.Text), "T")
    End If
End Sub

Private Sub txtGrupoResultados_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End If
End Sub





'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: SQL :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

