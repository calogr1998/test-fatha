VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Cliente"
   ClientHeight    =   7815
   ClientLeft      =   2655
   ClientTop       =   2475
   ClientWidth     =   15720
   ControlBox      =   0   'False
   Icon            =   "frmMantCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   15720
   Begin VB.PictureBox ptCaptcha 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   78
      Top             =   9240
      Width           =   855
   End
   Begin VB.CommandButton cmdActualizarCaptcha 
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   7080
      TabIndex        =   77
      Top             =   9720
      Width           =   255
   End
   Begin VB.TextBox txtCaptcha 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   8040
      TabIndex        =   76
      Top             =   9240
      Width           =   975
   End
   Begin VB.TextBox txtEmpAbrev 
      Height          =   285
      Left            =   9240
      MaxLength       =   10
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraContacto 
      Caption         =   " Datos Contacto "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   7920
      TabIndex        =   62
      Top             =   1920
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   735
         Left            =   5520
         Picture         =   "frmMantCliente.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtCCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   72
         Top             =   360
         Width           =   1440
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   735
         Left            =   4560
         Picture         =   "frmMantCliente.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   735
         Left            =   6480
         Picture         =   "frmMantCliente.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   735
         Left            =   4560
         Picture         =   "frmMantCliente.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox txtCObservaciones 
         Height          =   765
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   3960
         Width           =   5685
      End
      Begin VB.TextBox txtCEmail 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3600
         Width           =   5685
      End
      Begin VB.TextBox txtCMovil 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2880
         Width           =   2805
      End
      Begin VB.TextBox txtCTelefono2 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   28
         Top             =   2520
         Width           =   2805
      End
      Begin VB.TextBox txtCTelefono1 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   26
         Top             =   2160
         Width           =   2805
      End
      Begin VB.TextBox txtCAnexo 
         Height          =   285
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   27
         Top             =   2160
         Width           =   1725
      End
      Begin VB.TextBox txtCApePat 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   23
         Top             =   720
         Width           =   3285
      End
      Begin VB.TextBox txtCApeMat 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   24
         Top             =   1080
         Width           =   3285
      End
      Begin VB.TextBox txtCNombre 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   25
         Top             =   1440
         Width           =   3285
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
         Index           =   31
         Left            =   360
         TabIndex        =   73
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Observaciones"
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
         Index           =   30
         Left            =   360
         TabIndex        =   71
         Top             =   3960
         Width           =   1215
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
         Index           =   29
         Left            =   360
         TabIndex        =   70
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Telf. Movil"
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
         Index           =   28
         Left            =   360
         TabIndex        =   69
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefono (2)"
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
         Index           =   27
         Left            =   360
         TabIndex        =   68
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefono (1)"
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
         Index           =   26
         Left            =   360
         TabIndex        =   67
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Anexo"
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
         Index           =   25
         Left            =   4680
         TabIndex        =   66
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido Paterno"
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
         Index           =   24
         Left            =   360
         TabIndex        =   65
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido Materno"
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
         Index           =   23
         Left            =   360
         TabIndex        =   64
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombres"
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
         Index           =   22
         Left            =   360
         TabIndex        =   63
         Top             =   1440
         Width           =   1215
      End
   End
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
      Height          =   1455
      Left            =   7920
      TabIndex        =   59
      Top             =   360
      Width           =   7695
      Begin VB.TextBox txtObservacion 
         Height          =   645
         Left            =   1560
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   720
         Width           =   5925
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Observaciones"
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
         Left            =   240
         TabIndex        =   61
         Top             =   720
         Width           =   1215
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
         TabIndex        =   60
         Top             =   360
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
      Height          =   4335
      Left            =   120
      TabIndex        =   47
      Top             =   3360
      Width           =   7695
      Begin VB.TextBox txtWeb 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3960
         Width           =   5925
      End
      Begin VB.TextBox txtEmailPropuesta 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3600
         Width           =   5325
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   17
         Top             =   3240
         Width           =   5925
      End
      Begin VB.TextBox txtMovil 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   16
         Top             =   2880
         Width           =   2805
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2520
         Width           =   1965
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2520
         Width           =   2805
      End
      Begin VB.TextBox txtDirCobranza 
         Height          =   645
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1800
         Width           =   5925
      End
      Begin VB.TextBox txtDirRecepcion 
         Height          =   645
         Left            =   1560
         MaxLength       =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1080
         Width           =   5925
      End
      Begin VB.ComboBox cmbSector 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtAbreviatura 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   9
         Top             =   360
         Width           =   3045
      End
      Begin MSComCtl2.DTPicker dtpFechaIngreso 
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   65142785
         CurrentDate     =   41846
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Pagina Web"
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
         Index           =   19
         Left            =   240
         TabIndex        =   58
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "E-mail de Propuestas"
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
         Left            =   240
         TabIndex        =   57
         Top             =   3600
         Width           =   1815
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
         TabIndex        =   56
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Telf. Movil"
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
         TabIndex        =   55
         Top             =   2880
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
         TabIndex        =   54
         Top             =   2520
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
         TabIndex        =   53
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir. de Cobranza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   240
         TabIndex        =   52
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir. Recepción o Entregas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Ingreso"
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
         Left            =   4680
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Sector Emp."
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
         TabIndex        =   49
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "R. Social Abrev."
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
         Left            =   240
         TabIndex        =   48
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
      Height          =   3015
      Left            =   120
      TabIndex        =   37
      Top             =   360
      Width           =   7695
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   1440
      End
      Begin VB.ComboBox cmbDistrito 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Width           =   3015
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
         ItemData        =   "frmMantCliente.frx":1BB2
         Left            =   2760
         List            =   "frmMantCliente.frx":1BBC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtPais 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2520
         Width           =   2280
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
         ItemData        =   "frmMantCliente.frx":1C9D
         Left            =   360
         List            =   "frmMantCliente.frx":1CA7
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2295
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
         TabIndex        =   46
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "País"
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
         Left            =   4680
         TabIndex        =   45
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Distrito"
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Documento"
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label 
         Caption         =   "Origen de Cliente"
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
         TabIndex        =   38
         Top             =   360
         Width           =   2295
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgContacto1 
      Height          =   5775
      Left            =   16560
      OleObjectBlob   =   "frmMantCliente.frx":1D8B
      TabIndex        =   22
      Top             =   1920
      Width           =   7695
   End
   Begin MSComctlLib.Toolbar tlbCliente 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   635
      ButtonWidth     =   2275
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
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "frmMantCliente.frx":3DCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantCliente.frx":4365
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantCliente.frx":48FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantCliente.frx":4E99
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantCliente.frx":5433
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgContacto 
      Height          =   5775
      Left            =   7920
      OleObjectBlob   =   "frmMantCliente.frx":59CD
      TabIndex        =   74
      Top             =   1920
      Width           =   7695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1695
      Left            =   120
      TabIndex        =   79
      Top             =   8160
      Visible         =   0   'False
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   2990
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   975
      Left            =   7080
      TabIndex        =   80
      Top             =   8160
      Width           =   1935
      ExtentX         =   3413
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label21 
      Caption         =   "Abreviatura"
      Height          =   255
      Left            =   9240
      TabIndex        =   81
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmMantCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim nroCaptcha

Private bolAyuda        As Boolean
Private strCodCliente          As String

Private objCliente             As ClsCliente



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''OCR''''''''''''''''''''''''''''''''''''''
Const RGN_DIFF = 4
Const BaseSet = "mxnc5bvalskdj1fhgqpwoe2irtuy3VCBNX4MZGFHD8JKSAL9YOWP0QIERU6T7z !@#$%^&*()-=\_+|[];" & "'" & """" & ":,./?{}`~"

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim combined_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal HWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Const WM_SYSCOMMAND = &H112
''''''''''''''''''''FIN OCR'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodCliente = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodCliente
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
    objAyudaDistrito.listarDistrito cmbDistrito
    
    If cmbDistrito.ListCount > 0 Then
        cmbDistrito.ListIndex = ModUtilitario.seleccionarItem(cmbDistrito, "01", "DER", 2)
    End If
End Sub

Private Sub listarSector()
    objAyudaSector.listarSector cmbSector
    
    If cmbSector.ListCount > 0 Then
        cmbSector.ListIndex = ModUtilitario.seleccionarItem(cmbSector, "01", "DER", 2)
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
    txtPais.Text = vbNullString
    
    txtAbreviatura.Text = vbNullString
    dtpFechaIngreso.Value = Date
    If cmbSector.ListCount > 0 Then cmbSector.ListIndex = -1
    txtDirRecepcion.Text = vbNullString
    txtDirCobranza.Text = vbNullString
    txtTelefono.Text = vbNullString
    txtFax.Text = vbNullString
    txtMovil.Text = vbNullString
    txtEmail.Text = vbNullString
    txtEmailPropuesta.Text = vbNullString
    txtWeb.Text = vbNullString
    
    If cmbFormaPago.ListCount > 0 Then cmbFormaPago.ListIndex = -1
    txtObservacion.Text = vbNullString
    
    txtCodigo.Locked = True
    txtCodigo.BackColor = DF
    txtNroDocumento.Locked = False
    
    tlbCliente.buttons(4).Enabled = False
    tlbCliente.buttons(4).Visible = False
    tlbCliente.buttons(6).Enabled = False
    tlbCliente.buttons(6).Visible = False
End Sub

Private Sub consultarCliente()
    Set objCliente = New ClsCliente
    
    limpiarCajas
    
    With objCliente
        .inicializarEntidades
        
        .Codigo = strCodCliente
        
        If .obtenerCliente Then
            cmbOrigen.ListIndex = ModUtilitario.seleccionarItem(cmbOrigen, .OrigenCliente, "DER", 1)
            cmbPersona.ListIndex = ModUtilitario.seleccionarItem(cmbPersona, .ClaseCliente, "DER", 1)
            cmbTipoDocumento.ListIndex = ModUtilitario.seleccionarItem(cmbTipoDocumento, .CodigoTipoDocumento, "DER", 1)
            
            txtNroDocumento.Text = .NumeroDocumento
            txtCodigo.Text = .Codigo
            txtRazonSocial.Text = .NombreCliente
            txtDireccionFiscal.Text = .DireccionCliente
            cmbDistrito.ListIndex = ModUtilitario.seleccionarItem(cmbDistrito, .CodigoDistrito, "DER", 2)
            txtPais.Text = .Pais
            
            txtAbreviatura.Text = .NombreAbreviado
            dtpFechaIngreso.Value = IIf(.FechaReg <> vbNullString, .FechaReg, Date)
            cmbSector.ListIndex = ModUtilitario.seleccionarItem(cmbSector, .CodigoSector, "DER", 2)
            txtDirRecepcion.Text = .DireccionRecepcion
            txtDirCobranza.Text = .DireccionCobranza
            txtTelefono.Text = .Telefono
            txtFax.Text = .Fax
            txtMovil.Text = .Movil
            txtEmail.Text = .Email
            txtEmailPropuesta.Text = .EmailCotizacion
            txtWeb.Text = .Web
            
            cmbFormaPago.ListIndex = ModUtilitario.seleccionarItem(cmbFormaPago, .CodigoFormaPago, "DER", 3)
            txtObservacion.Text = .Observacion
            
            txtNroDocumento.Locked = False
            
            If Not bolAyuda Then
                tlbCliente.buttons(4).Enabled = True
                tlbCliente.buttons(4).Visible = True
                tlbCliente.buttons(6).Enabled = True
                tlbCliente.buttons(6).Visible = True
            End If
        End If
        
        .vistaClienteContacto dbgContacto
    End With
    
    Set objCliente = Nothing
End Sub

Private Sub validarCajas()
    If cmbOrigen.ListIndex = -1 Then
        MsgBox "El Campo Origen es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
        cmbOrigen.SetFocus
        
        Exit Sub
    End If
    
    If cmbPersona.ListIndex = -1 Then
        MsgBox "El Campo Persona es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
        cmbPersona.SetFocus
        
        Exit Sub
    End If
    
    If cmbTipoDocumento.ListIndex = -1 Then
        MsgBox "El Campo Tipo de Documento de Identidad es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
        cmbTipoDocumento.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtNroDocumento.Text) = vbNullString Then
        MsgBox "El Campo Numero de Documento de Identidad es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
        txtNroDocumento.SetFocus
        
        Exit Sub
    Else
        With objAyudaTipoDocID
            .inicializarEntidades
            
            .Codigo = right(cmbTipoDocumento.Text, 1)
            
            .obtenerConfigTipoDocumento
            
            If .TieneLargoFijo Then
                If Len(Trim(txtNroDocumento.Text)) < txtNroDocumento.MaxLength Then
                    MsgBox "La longitud del Numero de Documento de Identidad es incorrecto, verifique.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
                    
                    txtNroDocumento.SetFocus
                    
                    Exit Sub
                End If
            End If
        End With
    End If
    
    If Not txtCodigo.Locked Then
        If Trim(txtCodigo.Text) = vbNullString Then
            MsgBox "El Campo Código es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
            txtCodigo.SetFocus
            
            Exit Sub
        End If
    End If
    
    If Trim(txtRazonSocial.Text) = vbNullString Then
        MsgBox "El Campo Razon Social es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
        
        txtRazonSocial.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtDireccionFiscal.Text) = vbNullString Then
        MsgBox "El Campo Dirección Fiscal es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
        
        txtDireccionFiscal.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar los datos del Cliente?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarCliente
    End If
End Sub

Private Sub guardarCliente()
    Set objCliente = New ClsCliente
    
    With objCliente
        .inicializarEntidades
        
        .OrigenCliente = right(cmbOrigen.Text, 1)
        .ClaseCliente = right(cmbPersona.Text, 1)
        .CodigoTipoDocumento = right(cmbTipoDocumento.Text, 1)
        
        .NumeroDocumento = Trim(txtNroDocumento.Text)
        .Codigo = Trim(txtCodigo.Text)
        .NombreCliente = Replace(Trim(txtRazonSocial.Text), "'", "' & Chr(39) & '", 1)
        .DireccionCliente = Replace(Trim(txtDireccionFiscal.Text), "'", "' & Chr(39) & '", 1)
        .CodigoDistrito = right(cmbDistrito.Text, 2)
        .Pais = Trim(txtPais.Text)
        
        .NombreAbreviado = Trim(txtAbreviatura.Text)
        .FechaReg = Trim(dtpFechaIngreso.Value)
        .CodigoSector = right(cmbSector.Text, 2)
        .DireccionRecepcion = Trim(txtDirRecepcion.Text)
        .DireccionCobranza = Trim(txtDirCobranza.Text)
        .Telefono = Trim(txtTelefono.Text)
        .Fax = Trim(txtFax.Text)
        .Movil = Trim(txtMovil.Text)
        .Email = Trim(txtEmail.Text)
        .EmailCotizacion = Trim(txtEmailPropuesta.Text)
        .Web = Trim(txtWeb.Text)
        
        .CodigoFormaPago = right(cmbFormaPago.Text, 3)
        .Observacion = Trim(txtObservacion.Text)
        
        If Not txtNroDocumento.Locked And Trim(txtCodigo.Text) = vbNullString Then
            If .verificarExistenciaPorNroDocumento Then
                MsgBox "Numero de Documento ingresado ya existe, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                txtNroDocumento.SetFocus
                
                ModUtilitario.seleccionarTextoCaja txtNroDocumento
            End If
        End If
        
        If .guardarCliente Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            strCodCliente = .Codigo
            
            consultarCliente
            
            If Not bolAyuda Then
                MsgBox "Cliente Actualizado.", _
                        vbInformation, App.ProductName
            Else
                objAyudaCliente.Codigo = Trim(txtCodigo.Text)
                objAyudaCliente.NombreCliente = Trim(txtRazonSocial.Text)
                
                Unload Me
            End If
        End If
    End With
    
    Set objCliente = Nothing
End Sub

Private Sub eliminarCliente()
    Set objCliente = New ClsCliente
    
    With objCliente
        .Codigo = Trim(txtCodigo.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Cliente no existente.", vbInformation, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar la Cliente con Codigo '" & .Codigo & "'?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarCliente Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strCodCliente = .Codigo
                
                consultarCliente
                
                MsgBox "Cliente Eliminado.", _
                        vbInformation, App.ProductName
            End If
        End If
    End With
    
    Set objCliente = Nothing
End Sub

Private Function validarShortCut(ByVal Key As Integer) As Boolean
    validarShortCut = True
    
    Select Case Key
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(2)
        Case 7 'Ctrl + G (Guardar)
            tlbCliente_ButtonClick tlbCliente.buttons(3)
        Case 5 'Ctrl + E (Eliminar)
            tlbCliente_ButtonClick tlbCliente.buttons(4)
        Case 15 'Ctrl + O (Contacto)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 19 'Ctrl + S (Salir)
            tlbCliente_ButtonClick tlbCliente.buttons(8)
        Case Else
            validarShortCut = False
    End Select
End Function

Private Sub Form_Load()
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    'Llenar Distritos
    listarDistritos
    'Llenar Sector de Cliente
    listarSector
    'Llenar Formas de Pago
    listarFormaPago
    
    consultarCliente
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If Not bolAyuda Then
'        With frmListaCliente
'            .listarCliente
'        End With
'    End If
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Trim(Button.Caption)
        Case "&Nuevo"
            If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                strCodCliente = vbNullString
                
                consultarCliente
            End If
        Case "&Guardar"
            validarCajas
        Case "&Eliminar"
            eliminarCliente
        Case "C&ontacto"
            If MsgBox("¿Desea registrar un nuevo Contacto?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                consultarContacto vbNullString
            End If
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

Private Sub txtDireccionFiscal_LostFocus()
    If Trim(txtDirRecepcion.Text) = vbNullString Then
        txtDirRecepcion.Text = Trim(txtDireccionFiscal.Text)
    End If
    
    If Trim(txtDirCobranza.Text) = vbNullString Then
        txtDirCobranza.Text = Trim(txtDireccionFiscal.Text)
    End If
End Sub

Private Sub txtNroDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    'If Not txtNroDocumento.Locked Then Exit Sub
    
'''''    Select Case KeyCode
'''''        Case vbKeyReturn
'''''            If cmbTipoDocumento.ListIndex = -1 Then
'''''                MsgBox "Seleccione el Tipo de Documento.", vbInformation + vbOKOnly, App.ProductName
'''''
'''''                Exit Sub
'''''            End If
'''''
'''''            If right(cmbTipoDocumento.Text, 1) = "6" Then
'''''                If Len(txtNroDocumento.Text) < 11 Then
'''''                    MsgBox "Largo de Numero de RUC incorrecto, verifique.", vbInformation + vbOKOnly, App.ProductName
'''''
'''''                    Exit Sub
'''''                End If
'''''
'''''                If MsgBox("¿Desea verificar el Numero de Documento en SUNAT?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'''''                    Me.MousePointer = vbHourglass
'''''
'''''                    With objAyudaCliente
'''''                        .inicializarEntidades
'''''
''''''                        If ModUtilitario.validarRUCenSunat(Trim(txtNroDocumento.Text), True) Then
''''''                            txtRazonSocial.Text = ModUtilitario.limpiarCaracteresEnCadena(.NombreCliente)
''''''                            txtDireccionFiscal.Text = ModUtilitario.limpiarCaracteresEnCadena(.DireccionCliente)
''''''                            txtTelefono.Text = .Telefono
''''''
''''''                            MsgBox "Verificación de Cliente exitosa." & vbNewLine & _
''''''                                    "Estado: " & .CodigoVendedor & vbNewLine & _
''''''                                    "Situación: " & .CodigoCobrador, vbInformation + vbOKOnly, App.ProductName
''''''
''''''                            .inicializarEntidades
''''''
''''''                            ModUtilitario.pulsarTecla vbKeyTab
''''''                        End If
'''''
'''''                        .inicializarEntidades
'''''
'''''                        Validar_Ruc txtNroDocumento
'''''                    End With
'''''
'''''                    Me.MousePointer = vbDefault
'''''
'''''                    ModUtilitario.pulsarTecla vbKeyTab
'''''                Else
'''''                    ModUtilitario.pulsarTecla vbKeyTab
'''''                End If
'''''            End If
'''''    End Select
    
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
    If Not txtNroDocumento.Locked And Trim(txtCodigo.Text) = vbNullString Then
        With objAyudaCliente
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
'                        txtRazonSocial.Text = ModUtilitario.limpiarCaracteresEnCadena(.NombreCliente)
'                        txtDireccionFiscal.Text = ModUtilitario.limpiarCaracteresEnCadena(.DireccionCliente)
'                        txtTelefono.Text = .Telefono
'
'                        MsgBox "Verificación de Cliente exitosa." & vbNewLine & _
'                                "Estado: " & .CodigoVendedor & vbNewLine & _
'                                "Situación: " & .CodigoCobrador, vbInformation + vbOKOnly, App.ProductName
'
'                        .inicializarEntidades
'                    End If
                End If
            End If
        End With
    End If
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

Private Sub txtPais_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
    End If
End Sub

Private Sub txtAbreviatura_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub dtpFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub dtpFechaIngreso_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub cmbSector_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub txtDirRecepcion_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtDirCobranza_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtMovil_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtEmailPropuesta_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub txtweb_keypress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Private Sub cmbFormaPago_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarSoloTeclaEnter(KeyAscii)
    End If
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If Not validarShortCut(KeyAscii) Then
        KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End If
End Sub

Rem Mantenimiento de Contactos de Cliente
Private Sub limpiarCajasContacto()
    txtCCodigo.Text = vbNullString
    txtCApePat.Text = vbNullString
    txtCApeMat.Text = vbNullString
    txtCNombre.Text = vbNullString
    
    txtCTelefono1.Text = vbNullString
    txtCAnexo.Text = vbNullString
    txtCTelefono2.Text = vbNullString
    txtCMovil.Text = vbNullString
    
    txtCEmail.Text = vbNullString
    txtCObservaciones.Text = vbNullString
    
    txtCCodigo.Locked = True
    txtCCodigo.BackColor = DF
    
    cmdAgregar.Enabled = False
    cmdAgregar.Visible = False
    cmdActualizar.Enabled = False
    cmdActualizar.Visible = False
End Sub

Private Sub consultarContacto(ByVal strCodigoContacto As String)
    limpiarCajasContacto
    
    dbgContacto.Dataset.Close
    
    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'") <> vbNullString Then
        txtCCodigo.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCApePat.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "APEPAT", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCApeMat.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "APEMAT", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCNombre.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NOMBRE", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        
        txtCTelefono1.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TELEFONO1", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCAnexo.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "ANEXO", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCTelefono2.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TELEFONO2", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCMovil.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CELULAR", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        
        txtCEmail.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "EMAIL", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        txtCObservaciones.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "OBSERVACION", "CONTACTOS", "CODIGO", strCodigoContacto, "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'")
        
        cmdActualizar.Enabled = True
        cmdActualizar.Visible = True
    Else
        cmdAgregar.Enabled = True
        cmdAgregar.Visible = True
    End If
    
    fraDatos.Enabled = False
    fraAdicional.Enabled = False
    fraPredeteminada.Enabled = False
    
    fraContacto.Enabled = True
    fraContacto.Visible = True
    
    ModUtilitario.pulsarTecla vbKeyTab
End Sub

Private Sub validarCajasContacto()
    If Trim(txtCApePat.Text) = vbNullString Then
        MsgBox "El Campo Apellido Paterno es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
        txtCApePat.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtCNombre.Text) = vbNullString Then
        MsgBox "El Campo Nombres es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
            
        txtCNombre.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtCTelefono1.Text) = vbNullString Then
        MsgBox "El Campo Telefono(1) es obligatorio.", vbInformation + vbOKOnly + vbOKOnly, App.ProductName
        
        txtCTelefono1.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar los datos del Contacto?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarContacto
    End If
End Sub

Private Sub guardarContacto()
    Dim strSQL As String
    
    strSQL = vbNullString
    
    dbgContacto.Dataset.Close
    
    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "CONTACTOS", "CODIGO", Trim(txtCCodigo.Text), "T", "AND CODCLI = '" & Trim(txtCodigo.Text) & "'") = vbNullString Then
        txtCCodigo.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "CONTACTOS", "CODCLI", Trim(txtCodigo.Text), "T", "ORDER BY CODIGO DESC")) + 1, "0000")
        
        strSQL = strSQL & "INSERT INTO CONTACTOS("
        strSQL = strSQL & "CODCLI, CODIGO, APEPAT, APEMAT, NOMBRE, "
        strSQL = strSQL & "TELEFONO1, ANEXO, TELEFONO2, CELULAR, "
        strSQL = strSQL & "EMAIL, OBSERVACION) "
        strSQL = strSQL & "VALUES("
        strSQL = strSQL & "'" & Trim(txtCodigo.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCCodigo.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCApePat.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCApeMat.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCNombre.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCTelefono1.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCAnexo.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCTelefono2.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCMovil.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCEmail.Text) & "', "
        strSQL = strSQL & "'" & Trim(txtCObservaciones.Text) & "')"
    Else
        strSQL = strSQL & "UPDATE "
        strSQL = strSQL & "CONTACTOS "
        strSQL = strSQL & "SET "
        strSQL = strSQL & "APEPAT = '" & Trim(txtCApePat.Text) & "', "
        strSQL = strSQL & "APEMAT = '" & Trim(txtCApeMat.Text) & "', "
        strSQL = strSQL & "NOMBRE = '" & Trim(txtCNombre.Text) & "', "
        strSQL = strSQL & "TELEFONO1 = '" & Trim(txtCTelefono1.Text) & "', "
        strSQL = strSQL & "ANEXO = '" & Trim(txtCAnexo.Text) & "', "
        strSQL = strSQL & "TELEFONO2 = '" & Trim(txtCTelefono2.Text) & "', "
        strSQL = strSQL & "CELULAR = '" & Trim(txtCMovil.Text) & "', "
        strSQL = strSQL & "EMAIL = '" & Trim(txtCEmail.Text) & "', "
        strSQL = strSQL & "OBSERVACION = '" & Trim(txtCObservaciones.Text) & "' "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "CODCLI = '" & Trim(txtCodigo.Text) & "' AND "
        strSQL = strSQL & "CODIGO = '" & Trim(txtCCodigo.Text) & "'"
    End If
    
    cnn_dbbancos.Execute strSQL
    Actualiza_Log strSQL, StrConexDbBancos
    
    cmdSalir_Click
End Sub

Private Sub eliminarContacto()
    Dim strSQL As String
    
    strSQL = vbNullString
    
    If MsgBox("¿Desea eliminar el Contacto?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    dbgContacto.Dataset.Close
    
    strSQL = strSQL & "DELETE FROM "
    strSQL = strSQL & "CONTACTOS "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CODCLI = '" & Trim(txtCodigo.Text) & "' AND "
    strSQL = strSQL & "CODIGO = '" & Trim(txtCCodigo.Text) & "'"
    
    cnn_dbbancos.Execute strSQL
    Actualiza_Log strSQL, StrConexDbBancos
    
    cmdSalir_Click
End Sub

Private Sub cmdAgregar_Click()
    validarCajasContacto
End Sub

Private Sub cmdActualizar_Click()
    validarCajasContacto
End Sub

Private Sub cmdEliminar_Click()
    eliminarContacto
End Sub

Private Sub cmdSalir_Click()
    fraDatos.Enabled = True
    fraAdicional.Enabled = True
    fraPredeteminada.Enabled = True
    
    fraContacto.Enabled = False
    fraContacto.Visible = False
    
    With objAyudaCliente
        .inicializarEntidades
        
        .Codigo = Trim(txtCodigo.Text)
        
        .vistaClienteContacto dbgContacto
    End With
End Sub

Private Sub dbgContacto_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgContacto_OnDblClick
    End Select
End Sub

Private Sub dbgContacto_OnDblClick()
    consultarContacto Trim(dbgContacto.Columns.ColumnByFieldName("CODIGO").Value & "")
End Sub

Private Sub dbgContacto_OnKeyPress(Key As Integer)
    Select Case Key
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            'cmdSalir_Click
            tlbCliente_ButtonClick tlbCliente.buttons(8)
    End Select
End Sub

Private Sub txtCApePat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
    End Select
End Sub

Private Sub txtCApeMat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
    End Select
End Sub

Private Sub txtCNombre_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
    End Select
End Sub

Private Sub txtCTelefono1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End Select
End Sub

Private Sub txtCAnexo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End Select
End Sub

Private Sub txtCTelefono2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End Select
End Sub

Private Sub txtCMovil_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
    End Select
End Sub

Private Sub txtCEmail_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End Select
End Sub

Private Sub txtCObservaciones_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.buttons(6)
        Case 7 'Ctrl + G (Guardar)
            validarCajasContacto
        Case 5 'Ctrl + E (Eliminar)
            'tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            cmdSalir_Click
        Case Else
            KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End Select
End Sub









Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim ORange As Object
    Dim i As Integer
    
    On Error GoTo abcd
    
'    If bolConexionDBContiSoftNoDisponible Then
'        Exit Sub
'    End If
    
'    If strCodCliente <> vbNullString Then
'        Exit Sub
'    End If
    
    If Trim(txtNroDocumento.Text) = vbNullString Then
        Exit Sub
    End If
    
    Set ORange = WebBrowser1.document.body.createControlRange()

    For i = 0 To WebBrowser1.document.Images.Length - 1

            Call ORange.Add(WebBrowser1.document.Images.ITEM(i))
            Call ORange.execCommand("Copy")

            Set ptCaptcha.Picture = Clipboard.GetData

            Exit For
    Next
    
    'wrutatemp & "\TESERAC
    
    Set ptCaptcha.Picture = ptCaptcha.Image
    'Guardar Imagen del Picture 2
    'SavePicture ptCaptcha.Picture, App.Path & "\Captcha\imagen.bmp"
    SavePicture ptCaptcha.Picture, wrutatemp & "\TESERAC\Captcha\imagen.bmp"
    'Transformar Imagen a Caracteres
    cmdCaptcha_Click
    nroCaptcha = txtCaptcha.Text
    
    validarRUCenSunat txtNroDocumento, txtCaptcha
    
    txtRazonSocial.Text = objAyudaCliente.NombreCliente
    txtDireccionFiscal.Text = objAyudaCliente.DireccionCliente
    
    If txtDireccionFiscal.Text <> "" Then
    
        Dim oTratamientoDireccion As String
        Dim oTratamientoDepartamento As String
        Dim oTratamientoDistrito As String
        
        Dim oCodigoDepartamento As String
        Dim oCodigoDistrito As String
        
        oTratamientoDireccion = objAyudaCliente.DireccionCliente
        oTratamientoDistrito = Trim(Mid(oTratamientoDireccion, (InStrRev(objAyudaCliente.DireccionCliente, "-")) + 1))
        oTratamientoDireccion = oTratamientoDireccion '(Replace(oTratamientoDireccion, "- " & oTratamientoDistrito, ""))
        'oTratamientoDepartamento = Trim(Mid(oTratamientoDireccion, (InStrRev(oTratamientoDireccion, "-")) + 1))
        'oTratamientoDireccion = (Replace(oTratamientoDireccion, "- " & oTratamientoDepartamento, ""))
        
        'oCodigoDepartamento = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "DEPARTAMENTO", "DESCRIPCION", oTratamientoDepartamento, "T")  'ObtenerCampoWhere("DEPARTAMENTO", "CODIGO", "DESCRIPCION", oTratamientoDepartamento, "T", cnDB, "")
        'oCodigoDistrito = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "DISTRITO", "DISTRITO", oTratamientoDistrito, "T") 'ObtenerCampoWhere("DISTRITO", "CODIGO", "DISTRITO", oTratamientoDistrito, "T", cnDB, "")
        
        'If oCodigoDistrito <> 0 Then
            txtDireccionFiscal.Text = oTratamientoDireccion
        '    cmbDistrito.Text = oTratamientoDistrito
        'End If
        
    End If
    
    Exit Sub
abcd:
    If Err.Number = 91 Then Exit Sub
End Sub

Private Function Validar_Ruc(nroRuc As String)

    Dim oEstadoSunat            As String
    Dim oCondicionSunat         As String
    Dim oErrorRuc               As String
    
    WebBrowser1.Navigate ("http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image")
    
    
'    Command1_Click
'    cmdCaptcha_Click
'
'    onroCaptcha = txtCaptcha.Text
'
'    validarRUCenSunat txtnuevo, txtCaptcha

    '''''''''''''''''''''''''''''''cmdCaptcha_Click
    '''''''''''''''''''''''''''''''nroCaptcha = txtCaptcha.Text

'   If MsgBox("Consulta Terminada - SUNAT", vbOKOnly, "CONTROL Plus!") = vbOK Then
'        fn_DatosSunat
'    Else
'        Exit Function
'    End If

End Function


Public Function validarRUCenSunat(ByVal strRucConsulta As String, ByVal strRucCaptcha As String, _
                                    Optional ByVal bolMostrarMensaje As Boolean) As Boolean


'On Error GoTo errValidarRUCenSunat
'
'    validarRUCenSunat = False
'
'    Dim strTexto As String
'    Dim strPrincipio As String
'    Dim strFinal As String
'    Dim XML
'    Dim dblPosicion1 As Double
'    Dim dblPosicion2 As Double
'    Dim dblPosicion3 As Double
'    Dim dblPosicion4 As Double
'    Dim StrPosicion1 As String
'    Dim StrPosicion2 As String
'    Dim StrPosicion3 As String
'    Dim StrPosicion4 As String
'
'
'    strTexto = vbNullString
'
'    'Instanciar un Objeto XML de Microsoft
'    Set XML = CreateObject("Microsoft.XMLHTTP")
'    'Setear los parametros de Apertura del Objeto, para Consultar la URL
'    XML.Open "POST", "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRuc&nroRuc=" & strRucConsulta & "&" & "codigo=" & UCase(strRucCaptcha), False
'
'
'    'Ejecutamos la Pagina en el Objeto
'    XML.send
'    'Capturamos el Texto de Respuesta albergado en el Objeto
'    strTexto = XML.responseText
'
'    ''Enviar a un bloc de notas
'    Dim NombreArchivo As String
'    NombreArchivo = "ConsultaSunat"
'    'nombrearchivo = App.Path & "\" & nombrearchivo
'    NombreArchivo = wrutatemp & "\TESERAC\" & NombreArchivo
'
'    If Not right(NombreArchivo, 3) = "txt" Then
'        NombreArchivo = NombreArchivo & ".txt"
'    End If
'
'    Open NombreArchivo For Output As #1
'    Print #1, strTexto
'    Close #1
'    '''''''''''''''
'    If strTexto <> vbNullString Then
'        With objAyudaCliente
'            .inicializarEntidades
'
'            dblPosicion1 = InStr(strTexto, "<TITLE>")
'            If dblPosicion1 = 0 Then dblPosicion1 = InStr(strTexto, "<title>")
'            dblPosicion2 = InStr(strTexto, "</TITLE>")
'            If dblPosicion2 = 0 Then dblPosicion2 = InStr(strTexto, "</title>")
'
'            StrPosicion1 = Mid(strTexto, dblPosicion1, (dblPosicion2 - dblPosicion1))
'            .ConexionIE = Mid(StrPosicion1, InStr(StrPosicion1, ">") + 1, dblPosicion2)
'
'            If .ConexionIE = ".:: Pagina de Mensajes ::." Then
'                'Validar_Ruc strRucConsulta
'            End If
'
'            If .ConexionIE = ".:: Pagina de Error ::." Then
'                Validar_Ruc strRucConsulta
'            End If
'
'            'No Existe Ruc
'            If .ConexionIE = "Documento sin t&iacute;tulo" Then
'                dblPosicion1 = InStr(strTexto, strRucConsulta)
'                dblPosicion2 = InStr(strTexto, "</td>")
'                StrPosicion1 = Mid(strTexto, dblPosicion1, (dblPosicion1 - dblPosicion2) - 6)
'                StrPosicion2 = Mid(StrPosicion1, 1, InStr(StrPosicion1, "</strong>") - 2)
'
'                .NombreCliente = Mid(StrPosicion2, 13, InStr(StrPosicion1, "</strong>") - 2)
'                Exit Function
'            End If
'            'Ruc
'            dblPosicion1 = InStr(strTexto, strRucConsulta)
'            dblPosicion2 = InStr(strTexto, "</td>")
'
'            StrPosicion1 = Mid(strTexto, dblPosicion1, (dblPosicion1 - dblPosicion2) - 6)
'
'            .NumeroDocumento = Mid(StrPosicion1, 1, InStr(StrPosicion1, "-") - 2)
'
'            If .ConexionIE = "Documento sin t&iacute;tulo" Then
'                .NombreCliente = Mid(StrPosicion1, 1, InStr(StrPosicion1, "</strong>") - 2)
'            End If
'
'            'Razon Social
'            dblPosicion1 = InStr(strTexto, strRucConsulta)
'            dblPosicion2 = InStr(strTexto, ">Tipo")
'
'            StrPosicion1 = Replace(Mid(strTexto, dblPosicion1, (dblPosicion2 - dblPosicion1) - 6), "<", "")
'
'            .NombreCliente = Mid(StrPosicion1, InStr(StrPosicion1, "-") + 2, Len(StrPosicion1))
'            .NombreCliente = Replace(left(.NombreCliente, InStr(.NombreCliente, "/td") - 1), "'", "")
'
'            'Rus
'            dblPosicion1 = InStr(strTexto, "Afecto al Nuevo RUS: ")
'            dblPosicion2 = InStr(strTexto, "Fecha ")
'
'            StrPosicion1 = Replace(Mid(strTexto, dblPosicion1 + 69, (dblPosicion2 - dblPosicion1) - 168), "<", "")
'
'            .NumeroDocumento = StrPosicion1
'            If Not .RusCliente = "SI" Then .RusCliente = "NO"
'
'            'Direccion Fiscal
'            dblPosicion1 = InStr(strTexto, "Direcci")
'            dblPosicion2 = InStr(strTexto, "-->")
'            '''''Dirección + Departamento
'            StrPosicion1 = Mid(strTexto, dblPosicion1 + 84, (dblPosicion2 - dblPosicion1) - 174)
'            StrPosicion2 = Mid(StrPosicion1, 1, InStr(StrPosicion1, "                "))
'            '''''Provincia
'            dblPosicion3 = InStr(Mid(strTexto, dblPosicion1 + 84, (dblPosicion2 - dblPosicion1) - 202), "-")
'            StrPosicion3 = Mid(Mid(StrPosicion1, dblPosicion3, Len(StrPosicion1)), 1, InStr(Mid(StrPosicion1, dblPosicion3, Len(StrPosicion1)), "          "))
'            '''''Distrito
'            StrPosicion4 = Mid(Mid(StrPosicion1, dblPosicion3, Len(StrPosicion1)), InStr(Mid(StrPosicion1, dblPosicion3, Len(StrPosicion1)), "          "))
'            dblPosicion4 = InStr(StrPosicion4, "-")
'
'            .DireccionCliente = StrPosicion2 & StrPosicion3 & Mid(StrPosicion4, dblPosicion4, 50)
'
'            'Estado
'            dblPosicion1 = InStr(strTexto, "Estado del Contribuyente: ")
'            dblPosicion2 = InStr(strTexto, "Condici")
'
'            .CodigoVendedor = Mid(strTexto, dblPosicion1 + 72, (dblPosicion2 - dblPosicion1) - 254)
'            .CodigoVendedor = IIf(left(.CodigoVendedor, 1) = "A", "ACTIVO", left(.CodigoVendedor, 8))
'
'            'Situacion
'            dblPosicion1 = InStr(strTexto, "Condici")
'            dblPosicion2 = InStr(strTexto, "Direcci")
'
''            .CodigoCobrador = Trim(Mid(strTexto, dblPosicion1 + 18, (dblPosicion2 - dblPosicion1) - 53))
'            .CodigoCobrador = Mid(strTexto, dblPosicion1 + 113, (dblPosicion2 - dblPosicion1) - 240)
'
'            If Not .CodigoCobrador = "HABIDO" Then .CodigoCobrador = "NO HABIDO"
''            'Telefono
''            dblPosicion1 = InStr(strTexto, "Tel")
''            dblPosicion2 = InStr(strTexto, "Dependencia")
''
''            .Telefono = left(Trim(Mid(strTexto, dblPosicion1 + 26, (dblPosicion2 - dblPosicion1) - 53)), 7)
''            If Not IsNumeric(.Telefono) Then
''                .Telefono = vbNullString
''            End If
'        End With
'
'        validarRUCenSunat = True
'
'        If objAyudaCliente.NombreCliente = " " Then
'            Validar_Ruc objAyudaCliente.NumeroDocumento
'        End If
'    Else
'        If bolMostrarMensaje Then
'            MsgBox "Conexión a Web de SUNAT fállida." & vbNewLine & _
'                    "Pagina de SUNAT no disponible por el momento o " & vbNewLine & _
'                    "problemas de conexión a Internet." & vbNewLine & vbNewLine & _
'                    "Verificación Fallida.", vbInformation + vbOKOnly, App.ProductName
'        End If
'
'        validarRUCenSunat = False
'    End If
'
'    Set XML = Nothing
'
'
'
'    Exit Function
'    Resume
'errValidarRUCenSunat:
'    Select Case Err.Number
'        Case 3704, 3709
''            modCPlus.abrirCn 'cnDB.Open cnDB.ConnectionString 'strCadenaConexionDB
'
'            Resume
'        Case Else
'            If bolMostrarMensaje Then
'                MsgBox "Error No.: " & Err.Number & vbNewLine & _
'                        "Descripción: " & Err.Description & vbNewLine & _
'                        "Conexión a Web de SUNAT fállida." & vbNewLine & _
'                        "Pagina de SUNAT no disponible por el momento o problemas de conexión a Internet." & vbNewLine & vbNewLine & _
'                        "Numero de Documento Invalido o Verificación Fallida.", vbExclamation + vbOKOnly, _
'                        App.ProductName & " - modUtilitario: ValidarRUCenSunat"
'            End If
'    End Select
'
'    validarRUCenSunat = False
'
'    Err.Clear
End Function

Private Sub Command1_Click()
    Dim ORange As Object
    Dim i As Integer

    WebBrowser1.Navigate ("http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image")
    
    Set ORange = WebBrowser1.document.body.createControlRange()

    For i = 0 To WebBrowser1.document.Images.Length - 1

        Call ORange.Add(WebBrowser1.document.Images.ITEM(i))
        Call ORange.execCommand("Copy")

        Set ptCaptcha.Picture = Clipboard.GetData

        Exit For
    Next
    
    Set ptCaptcha.Picture = ptCaptcha.Image
    'Guardar Imagen del Picture 2
    'SavePicture ptCaptcha.Picture, App.Path & "\Captcha\imagen.bmp"
    SavePicture ptCaptcha.Picture, wrutatemp & "\TESERAC\Captcha\imagen.bmp"
    
    'Transformar Imagen a Caracteres
    cmdCaptcha_Click
End Sub


Private Sub cmdCaptcha_Click()
    
'    On Error GoTo errorc
'
'    Dim hWn As Long
'    Dim sTemp As String
'    Dim nroCatpcha As String
''    Set oFso = New FileSystemObject
''    If oFso.FolderExists("C:\Captcha\") = False Then MkDir ("C:\Captcha\")
'
'    'SetAttr App.Path & "\cap.exe", 38
'    SetAttr wrutatemp & "\TESERAC\cap.exe", 38
'    'SetAttr App.Path & "\tessdata", 38
'    SetAttr wrutatemp & "\TESERAC\tessdata", 38
'    'SetAttr App.Path & "\Captcha", 38
'    SetAttr wrutatemp & "\TESERAC\Captcha", 38
'
'    'wrutatemp & "\TESERAC
'
''    hWn = ShellAndWaitForTermination(App.Path & "\cap " & App.Path & "\Captcha\imagen.bmp " & App.Path & "\output", vbHide)
'    'hWn = ShellAndWaitForTermination(Chr(34) & App.Path & "\cap" & Chr(34) & " " & Chr(34) & App.Path & "\Captcha\imagen.bmp" & Chr(34) & " " & Chr(34) & App.Path & "\output" & Chr(34), vbHide)
'    hWn = ShellAndWaitForTermination(Chr(34) & wrutatemp & "\TESERAC\cap" & Chr(34) & " " & Chr(34) & wrutatemp & "\TESERAC\Captcha\imagen.bmp" & Chr(34) & " " & Chr(34) & wrutatemp & "\TESERAC\output" & Chr(34), vbHide)
'    Dim tmpString As String, strTemp As Byte, i As Integer
'    'Open App.Path & "\output.txt" For Binary As #1
'    Open wrutatemp & "\TESERAC\output.txt" For Binary As #1
'        For i = 1 To LOF(1)
'            Get 1, i, strTemp
'            If InStr(1, BaseSet, Chr(strTemp)) Then
'                If Chr(strTemp) = "5" Then
'                    tmpString = tmpString & "S"
'                ElseIf Chr(strTemp) = "3" Then
'                    tmpString = tmpString & "B"
'                ElseIf Chr(strTemp) = "1" Then
'                    tmpString = tmpString & "I"
'                ElseIf Chr(strTemp) = "/" Or Chr(strTemp) = "," Or Chr(strTemp) = "_" Or Chr(strTemp) = "'" Or Chr(strTemp) = "-" Then
'                    Close #1
'                    cmdActualizarCaptcha_Click
'                Else
'                    tmpString = tmpString & Chr(strTemp)
'                End If
'            End If
'        Next i
'    Close #1
'    txtCaptcha.Text = UCase(Replace(tmpString, " ", ""))
'    'If txtCaptcha.Text = "" Or Len(Trim(txtCaptcha.Text)) <> 4 Then
'        'cmdActualizarCaptcha_Click
'    'End If
'
'    Exit Sub
'errorc:
'    If Err.Number = 52 Then
'        'cmdActualizarCaptcha_Click
'    End If
End Sub

Private Sub cmdActualizarCaptcha_Click()
    WebBrowser1.Navigate ("http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image")
End Sub
