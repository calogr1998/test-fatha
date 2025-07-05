VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUtilProcesoInicioSM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilitarios de Administrador"
   ClientHeight    =   9240
   ClientLeft      =   2100
   ClientTop       =   1140
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   15270
   Begin VB.CommandButton cmdProceso 
      Cancel          =   -1  'True
      Caption         =   "Actualizar Categoria y OP en Historial de Cambios de OP"
      Height          =   255
      Index           =   55
      Left            =   5160
      TabIndex        =   93
      Top             =   4440
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Reiniciar Registro de Compra de Vale de Ingreso x Compra"
      Height          =   255
      Index           =   54
      Left            =   5160
      TabIndex        =   92
      Top             =   4080
      Width           =   4935
   End
   Begin VB.Frame fraReinicioRcDeVale 
      Caption         =   " Reinicio de Registro de Compra de Vale de Ingreso por Compra "
      Height          =   735
      Left            =   10200
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtReinicioTipoVale 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   90
         Text            =   "I"
         ToolTipText     =   "Tipo de Vale: I / S"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtReinicioNumeroVale 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   88
         Text            =   "Text1"
         ToolTipText     =   "Numero de Vale: I-XXXXXXX / S-XXXXXXX"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdReinicioReiniciar 
         Caption         =   "Reiniciar"
         Height          =   255
         Left            =   3720
         TabIndex        =   89
         ToolTipText     =   "Exportar Vale CP a SQL"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtReinicioCodAlmacen 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   87
         Text            =   "Text1"
         ToolTipText     =   "Codigo de Almacen"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Vale"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraCambioAlmacen 
      Caption         =   " Cambio de Almacen "
      Height          =   1935
      Left            =   10200
      TabIndex        =   73
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtCambioCodAlmacenNuevo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtCambioTipo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   77
         Text            =   "Text1"
         ToolTipText     =   "Tipo: I / S"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCambioIdNumero 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   78
         Text            =   "Text1"
         ToolTipText     =   "ID: XXXXXXX"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtCambioTipoVale 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         MaxLength       =   1
         TabIndex        =   74
         Text            =   "Text1"
         ToolTipText     =   "Tipo de Vale: I / S"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtCambioNumeroVale 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   76
         Text            =   "Text1"
         ToolTipText     =   "Numero de Vale: I-XXXXXXX / S-XXXXXXX"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdCambioCambiar 
         Caption         =   "Cambiar"
         Height          =   255
         Left            =   3960
         TabIndex        =   80
         ToolTipText     =   "Exportar Vale CP a SQL"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCambioCodAlmacen 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   75
         Text            =   "Text1"
         ToolTipText     =   "Codigo de Almacen"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Nuevo Almacen"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblCambioAlmacen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   255
         Left            =   1920
         TabIndex        =   84
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "ID Externo"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Ó"
         Height          =   255
         Left            =   2040
         TabIndex        =   82
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Vale"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Cambiar Almacen de Vale de Almacen"
      Height          =   255
      Index           =   53
      Left            =   5160
      TabIndex        =   72
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Frame fraRegistroEspecifico 
      Caption         =   " Exportar Registro Especifico "
      Height          =   1335
      Left            =   10200
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtCodAlmacen 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   68
         Text            =   "Text1"
         ToolTipText     =   "Codigo de Almacen"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdExportarVale 
         Caption         =   "Exportar"
         Height          =   255
         Left            =   3720
         TabIndex        =   70
         ToolTipText     =   "Exportar Vale CP a SQL"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtNumeroVale 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   69
         Text            =   "Text1"
         ToolTipText     =   "Numero de Vale: I-XXXXXXX / S-XXXXXXX"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtTipoVale 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   67
         Text            =   "Text1"
         ToolTipText     =   "Tipo de Vale: I / S"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdExportarOrden 
         Caption         =   "Exportar"
         Height          =   255
         Left            =   3720
         TabIndex        =   65
         ToolTipText     =   "Exportar Orden CP a SQL"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtNumeroOrden 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   64
         Text            =   "Text1"
         ToolTipText     =   "Numero de Orden: OCXXXXXX / OSXXXXXX"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtTipoOrden 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   63
         Text            =   "Text1"
         ToolTipText     =   "Tipo de Orcen: OC / OS"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Vale"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Orden"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraSQLvsAccess 
      Caption         =   " Cantidad Registros: SQL vs Access"
      Height          =   5535
      Left            =   10200
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ListBox lstAccess 
         Height          =   4740
         Left            =   2520
         TabIndex        =   58
         Top             =   720
         Width           =   2175
      End
      Begin VB.ListBox lstSQL 
         Height          =   4740
         Left            =   240
         TabIndex        =   57
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cmbTipoComparacion 
         Height          =   315
         ItemData        =   "frmUtilProcesoInicioSM.frx":0000
         Left            =   1200
         List            =   "frmUtilProcesoInicioSM.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Vista"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Eliminar Registro de Compromisos Pendientes"
      Height          =   255
      Index           =   52
      Left            =   5160
      TabIndex        =   71
      Top             =   3360
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL Registro Especifico"
      Height          =   255
      Index           =   51
      Left            =   5160
      TabIndex        =   60
      Top             =   3000
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Cantidad de Registros: SQL vs Acess"
      Height          =   255
      Index           =   50
      Left            =   16680
      TabIndex        =   59
      Top             =   8520
      Width           =   4935
   End
   Begin VB.Frame fraProceso2 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   120
      TabIndex        =   52
      Top             =   8520
      Visible         =   0   'False
      Width           =   15015
      Begin ComctlLib.ProgressBar pgbProceso2 
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Analisis y Correccion de Stock Libre Negativo"
      Height          =   255
      Index           =   49
      Left            =   16680
      TabIndex        =   51
      Top             =   8160
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Recuperar OCs desde BD Backup"
      Height          =   255
      Index           =   48
      Left            =   16680
      TabIndex        =   50
      Top             =   7800
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL las Tareas de Usuario de CP"
      Height          =   255
      Index           =   47
      Left            =   5160
      TabIndex        =   49
      Top             =   2280
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Clientes de CP"
      Height          =   255
      Index           =   45
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Tipo de Cambio de CP"
      Height          =   255
      Index           =   43
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   840
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL las Formas de Pago de CP"
      Height          =   255
      Index           =   42
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Tipos de Comprobantes de CP"
      Height          =   255
      Index           =   41
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Tipos de Clientes de CP"
      Height          =   255
      Index           =   40
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7320
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL las Zonas de CP"
      Height          =   255
      Index           =   39
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6960
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Tipo de Doc. Identidad de CP"
      Height          =   255
      Index           =   38
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6600
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL las Categorias de CP"
      Height          =   255
      Index           =   37
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6240
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL las Medi-Ventas de CP"
      Enabled         =   0   'False
      Height          =   255
      Index           =   36
      Left            =   120
      TabIndex        =   38
      Top             =   5880
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Productos de CP"
      Height          =   255
      Index           =   35
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5520
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Niveles 02 de CP"
      Height          =   255
      Index           =   34
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5160
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Niveles 01 de CP"
      Height          =   255
      Index           =   33
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4800
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Colores de Bien de CP"
      Height          =   255
      Index           =   32
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4440
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL las Medidas de CP"
      Height          =   255
      Index           =   31
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4080
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Marcas de CP"
      Height          =   255
      Index           =   30
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3720
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Centros de CP"
      Height          =   255
      Index           =   29
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3360
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Tipos Existencia de CP"
      Height          =   255
      Index           =   28
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3000
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Almacenes de CP"
      Height          =   255
      Index           =   27
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Origenes de CP"
      Height          =   255
      Index           =   26
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2280
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL los Cierres Mensuales de CP"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   27
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL las Tomas de Inventario de CP"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL las Ordenes de Compra de CP"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL los Requerimientos de CP"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Corregir Items Duplicados en Detalle de OC's"
      Enabled         =   0   'False
      Height          =   255
      Index           =   20
      Left            =   16680
      TabIndex        =   22
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Descartar Reposicion de Compromisos Automaticamente"
      Enabled         =   0   'False
      Height          =   255
      Index           =   19
      Left            =   16680
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Actualizar Costo Promedio en Vales de Ajuste Dic-2014 - 2"
      Enabled         =   0   'False
      Height          =   255
      Index           =   18
      Left            =   16680
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Ajustes en Vales de Ajuste Dic-2014 en base Inventario Ene-15"
      Enabled         =   0   'False
      Height          =   255
      Index           =   17
      Left            =   16680
      TabIndex        =   19
      Top             =   6360
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Actualizar Costo Promedio en Vales de Ajuste Dic-2014"
      Enabled         =   0   'False
      Height          =   255
      Index           =   16
      Left            =   16680
      TabIndex        =   18
      Top             =   6000
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Analisis y Correccion de Stock Compromiso Actualmente"
      Enabled         =   0   'False
      Height          =   255
      Index           =   15
      Left            =   16680
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Corregir Stock Comprometido Negativos"
      Enabled         =   0   'False
      Height          =   255
      Index           =   14
      Left            =   16680
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar Toma de Inventarios"
      Enabled         =   0   'False
      Height          =   255
      Index           =   13
      Left            =   16680
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Importar Toma de Inventarios"
      Enabled         =   0   'False
      Height          =   255
      Index           =   12
      Left            =   16680
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Actualizar Estado de Descarga de O/Ps"
      Enabled         =   0   'False
      Height          =   255
      Index           =   11
      Left            =   16680
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Verificar Consistencia de Stock de Producto con Movimiento"
      Enabled         =   0   'False
      Height          =   255
      Index           =   10
      Left            =   16680
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Generar Registro de Compras en base Vales de Ing x Compra"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   16680
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar Proveedores de CP al Sistema Integrado"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   16680
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Verificar Atención de Ordenes de Compra"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   16680
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   7800
      Visible         =   0   'False
      Width           =   15015
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Generar Compromiso Automatico a parti de OPs Pendientes"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   16680
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Importar Ingresos y Salidas del Mes Actual"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   16680
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Generar Vales Iniciales a partir de Cierre del Mes Anterior"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   16680
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Importar Productos de Proveedor"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   16680
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Importar Personas (Cliente/Proveedor)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   16680
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Importar Colores"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   16680
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Importar Insumos"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   16680
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Exportar a SQL los Vales de Ingreso/Salida de CP"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Proveedores de CP"
      Height          =   255
      Index           =   44
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1200
      Width           =   4935
   End
   Begin VB.CommandButton cmdProceso 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar a SQL los Usuarios de CP"
      Height          =   255
      Index           =   46
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1920
      Width           =   4935
   End
End
Attribute VB_Name = "frmUtilProcesoInicioSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objValeProcesoIni   As ClsVale

Private Sub cmbTipoComparacion_Click()
    cantidadRegistrosSQLvsAccess
End Sub

Private Sub cmdCambioCambiar_Click()
    cambioDeAlmacenVale
End Sub

Private Sub cmdExportarOrden_Click()
    exportarSqlOrdenCompraEspecificaCP
End Sub

Private Sub cmdExportarVale_Click()
    exportarSqlValeEspecificoCP
End Sub

Private Sub cmdProceso_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0 'Importar Insumos
            ModMilano.importarInsumoServidorExterno fraProceso, pgbProceso
        Case 1 'Importar Colores
            ModMilano.importarColorServidorExterno fraProceso, pgbProceso
        Case 2 'Importar Personas (Clientes / Proveedores)
            ModMilano.importarPersonasServidorExterno fraProceso, pgbProceso
        Case 3 'Importar Productos de Proveedor
            ModMilano.importarProductoPorProveedorServidorExterno fraProceso, pgbProceso
        Case 4 'Generar Vales Iniciales a partir de Cierre
            ModMilano.importarCierreMesValesInicialesServidorExterno fraProceso, pgbProceso
        Case 5 'Importar Ingresos y Salidas
            ModMilano.importarValesServidorExterno "I", fraProceso, pgbProceso
            
            ModMilano.importarValesServidorExterno "S", fraProceso, pgbProceso
        Case 6 'Generar Compromiso Automatico a parti de OPs Pendientes
            ModMilano.importarOpPendientesParaCompromisoInicialServidorExterno fraProceso, pgbProceso
        Case 7 'Verificar Atención de Ordenes de Compra
            verificarAtencionOrden
        Case 8 'Exportar Proveedores de CP al Sistema Integrado
            exportarProveedoresCPaIntegrado
        Case 9 'Generar Registro de Compra en Base a Vales de Ingreso por Compra
            generarRegistroCompraDeValeIngPorCompra
        Case 10 'Verificar Consistencia de Stock de Producto con Movimiento
            verificarConsistenciaStockProducto
        Case 11 'Actualizar Estado de Descarga de O/Ps
            actualizarEstadoDescargaOPs
        Case 12 'Importar Toma de Inventarios
            ModMilano.importarTomaInventarioServidorExterno fraProceso, pgbProceso
        Case 13 'Exportar Toma de Inventarios
            ModMilano.exportarTomaInventarioServidorExterno fraProceso, pgbProceso
        Case 14 'Corregir Stock Comprometido Negativos
            'correccionStockComprometidoEnNegativo
            
            correccionStockComprometidoEnNegativoV2
        Case 15 'Analisis y Correccion de Stock Compromiso Actualmente
            analisisYCorreccionStockComprometidoActual
        Case 16 'Actualizar Costo Promedio en Vales de Ajuste Dic-2014
            correccionCostoPromedioEnValesAjusteDic14
        Case 17 'Ajustes en Vales de Ajuste Dic-2014 en base Inventario Ene-15
'            correccionCantidadEnValesAjusteDic14
        Case 18 'Actualizar Costo Promedio en Vales de Ajuste Dic-2014 - 2
            correccionCostoPromedioEnValesAjusteDic14paso2
        Case 19 'Descartar Reposicion de Compromisos Automaticamente
            descarteAutomaticoDeReposicionCompromiso
        Case 20 'Corregir Items Duplicados en Detalle de OC's
            corregirDuplicadosEnDetalleOC
        Case 21 'Exportar a SQL los Requerimientos de CP
            exportarSqlRequerimientoCP
        Case 22 'Exportar a SQL las Ordenes de Compra de CP
            exportarSqlOrdenCompraCP
        Case 23 'Exportar a SQL los Vales de Ingreso/Salida de CP
            exportarSqlValeCP
        Case 24 'Exportar a SQL las Tomas de Inventario de CP
            exportarSqlTomaInventarioCP
        Case 25 'Exportar a SQL los Cierres Mensuales de CP
            exportarSqlCierreValeCP
        
        
        Case 26 'Exportar a SQL los Origenes de CP
            exportarSqlOrigenesCP
        Case 27 'Exportar a SQL los Almacenes de CP
            exportarSqlAlmacenesCP
        Case 28 'Exportar a SQL los Tipos Existencia de CP
            exportarSqlTiposExistenciasCP
        Case 29 'Exportar a SQL los Centros de CP
            exportarSqlCentrosCP
        Case 30 'Exportar a SQL los Marcas de CP
            exportarSqlMarcasCP
        Case 31 'Exportar a SQL las Medidas de CP
            exportarSqlMedidasCP
        Case 32 'Exportar a SQL los Colores de Bien de CP
            exportarSqlBienColorCP
        Case 33 'Exportar a SQL los Niveles 01 de CP
            exportarSqlNivel1CP
        Case 34 'Exportar a SQL los Niveles 02 de CP
            exportarSqlNivel2CP
        Case 35 'Exportar a SQL los Productos de CP
            exportarSqlBienCP
        Case 37 'Exportar a SQL las Categorias de CP
            exportarSqlCategoriaCP
        Case 38 'Exportar a SQL los Tipo de Doc. Identidad de CP
            exportarSqlTipoDocumentoIDCP
        Case 39 'Exportar a SQL las Zonas de CP
            exportarSqlDistritosCP
        Case 40 'Exportar a SQL los Tipos de Clientes de CP
            exportarSqlTipoClienteCP
        Case 41 'Exportar a SQL los Tipos de Comprobantes de CP
            exportarSqlTiposComprobantesCP
        Case 42 'Exportar a SQL las Formas de Pago de CP
            exportarSqlFormaPagoCP
        Case 43 'Exportar a SQL los Tipo de Cambio de CP
'            exportarSqlTipoCambioCP
        Case 44 'Exportar a SQL los Proveedores de CP
            exportarSqlProveedoresCP
        Case 45 'Exportar a SQL los Clientes de CP
            exportarSqlClientesCP
        Case 46 'Exportar a SQL los Usuarios de CP
            exportarSqlUsuariosCP
        Case 47 'Exportar a SQL las Tareas de Usuario de CP
            exportarSqlTareasCP
            
        
        
        Case 48 'Recuperar OCs desde BD Backup
            recuperarOrdenCompraBackupCP
        Case 49 'Analisis y Correccion de Stock Libre Negativo
            analisisYCorreccionStockLibreActual
        Case 50 'Cantidad de Registros: SQL vs Acess
            lstSQL.Clear
            lstAccess.Clear
            cmbTipoComparacion.ListIndex = -1
            
            fraSQLvsAccess.Visible = Not CBool(fraSQLvsAccess.Visible)
            fraRegistroEspecifico.Visible = False
            fraCambioAlmacen.Visible = False
            fraReinicioRcDeVale.Visible = False
        Case 51 'Exportar a SQL Registro Especifico
            txtTipoOrden.Text = vbNullString
            txtNumeroOrden.Text = vbNullString
            
            txtTipoVale.Text = vbNullString
            txtCodAlmacen.Text = vbNullString
            txtNumeroVale.Text = vbNullString
            
            fraSQLvsAccess.Visible = False
            fraRegistroEspecifico.Visible = Not CBool(fraRegistroEspecifico.Visible)
            fraCambioAlmacen.Visible = False
            fraReinicioRcDeVale.Visible = False
        Case 52 'Eliminar Registro de Compromisos Pendientes
            Dim StrFecha As String
            Dim dblCantReg As Double
            
            StrFecha = InputBox("Ingresa Fecha a Eliminar:", "Limpieza", Date)
            dblCantReg = 0
            
            If IsDate(StrFecha) Then
                If MsgBox("¿Desea eliminar el Registro de Compromisos de la Fecha '" & StrFecha & "'?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                    cnn_dbbancos.Execute "DELETE FROM SF3COMPROMISOAUTOMATICO WHERE CVDATE(FECHAEJECUCION) = CVDATE('" & StrFecha & "')"
                    
                    cnn_dbbancos.Execute "DELETE FROM SF4COMPROMISOAUTOMATICO WHERE CVDATE(FECHAEJECUCION) = CVDATE('" & StrFecha & "')", dblCantReg
                    
                    MsgBox dblCantReg & " Registro de Compromiso Eliminado.", vbInformation + vbOKOnly, App.ProductName
                End If
            End If
            
            StrFecha = 0
            dblCantReg = 0
        Case 53 'Cambiar Almacen de Vale de Almacen
            txtCambioTipoVale.Text = vbNullString
            txtCambioCodAlmacen.Text = vbNullString
            txtCambioNumeroVale.Text = vbNullString
            
            txtCambioTipo.Text = vbNullString
            txtCambioIdNumero.Text = vbNullString
            
            txtCambioCodAlmacenNuevo.Text = vbNullString
            lblCambioAlmacen.Caption = vbNullString
            
            fraSQLvsAccess.Visible = False
            fraRegistroEspecifico.Visible = False
            fraCambioAlmacen.Visible = Not CBool(fraCambioAlmacen.Visible)
            fraReinicioRcDeVale.Visible = False
        Case 54 'Reiniciar Registro de Compra de Vale de Ingreso x Compra
            txtReinicioTipoVale.Text = "I"
            txtReinicioCodAlmacen.Text = vbNullString
            txtReinicioNumeroVale.Text = vbNullString
            
            fraSQLvsAccess.Visible = False
            fraRegistroEspecifico.Visible = False
            fraCambioAlmacen.Visible = False
            fraReinicioRcDeVale.Visible = Not CBool(fraReinicioRcDeVale.Visible)
        Case 55 'Actualizar Categoria y OP en Historial de Cambios de OP
            actualizarCatYopHistorialCambiosOP
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub verificarAtencionOrden()
    Dim rstOrdenDet As New ADODB.Recordset
    
    If rstOrdenDet.State = 1 Then rstOrdenDet.Close
    
    rstOrdenDet.Open "SELECT F4NUMORD FROM IF4ORDEN WHERE F4LOCAL = 'OC' ORDER BY F4NUMORD DESC", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstOrdenDet.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstOrdenDet)
        pgbProceso.Value = 0
        fraProceso.Caption = "Verificación de Atención de O/C's..."
        
        Do While Not rstOrdenDet.EOF
            With objAyudaOrden
                .inicializarEntidades
                
                .TipoOrden = "OC"
                .NumeroOrden = Trim(rstOrdenDet!F4NUMORD & "")
                
                If .obtenerOrden Then
                    If .Estado <> 7 And .Estado <> 8 Then
                        .atencionOrden
                    End If
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Verificación de Atención de O/C's... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstOrdenDet.MoveNext
        Loop
            
            MsgBox "Verificación de O/C's finalizado.", vbInformation + vbOKOnly, App.ProductName
    End If
End Sub

Private Sub exportarProveedoresCPaIntegrado()
    Dim rstProveedor As New ADODB.Recordset
    
    If rstProveedor.State = 1 Then rstProveedor.Close
    
    rstProveedor.Open "SELECT F2CODPROV FROM EF2PROVEEDORES WHERE TRIM(F2CODPROVEXTERNO & '') = '' ORDER BY F2CODPROV", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProveedor.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstProveedor)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando Proveedores..."
        
        Do While Not rstProveedor.EOF
            If ModMilano.exportarProveedorAserverSQL(Trim(rstProveedor!F2CODPROV & "")) Then
                
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando Proveedores... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstProveedor.MoveNext
        Loop
    End If
End Sub

Private Sub generarRegistroCompraDeValeIngPorCompra()
    On Error GoTo errGenerarRegistroCompraDeValeIngPorCompra
    
    Dim rstVales As New ADODB.Recordset
    
    Dim rstVale As ADODB.Recordset
    Dim rstValeDet As ADODB.Recordset
    Dim dblItem As Double
    Dim dblMontoCancelado As Double
    
    abrirCnContaTabla
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F4TIPOVALE, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F4NUMVAL, "
    SqlCad = SqlCad & "F4FECVAL, "
    SqlCad = SqlCad & "F1CODORI, "
    SqlCad = SqlCad & "F2CODPROV, "
    SqlCad = SqlCad & "F4REFERE, "
    SqlCad = SqlCad & "F4SERGUIA, "
    SqlCad = SqlCad & "F4NUMGUIA, "
    SqlCad = SqlCad & "F4TIPDOC, "
    SqlCad = SqlCad & "F4SERDOC, "
    SqlCad = SqlCad & "F4NUMDOC, "
    SqlCad = SqlCad & "F4FECULT, "
    SqlCad = SqlCad & "F4MONEDA, "
    SqlCad = SqlCad & "F4TIPCAM, "
    SqlCad = SqlCad & "F4OBSERVA, "
    SqlCad = SqlCad & "F4REGCOM "
    SqlCad = SqlCad & "EXPORTARVALE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF4VALES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "VAL(MID(F4NUMVAL, 3,2)) >= 11 AND "
    SqlCad = SqlCad & "YEAR(F4FECVAL) = 2015 AND "
    SqlCad = SqlCad & "MONTH(F4FECVAL) = 7 AND "
    SqlCad = SqlCad & "F1CODORI = 'XC0' AND "
    SqlCad = SqlCad & "TRIM(F2CODPROV & '') <> '' AND "
    SqlCad = SqlCad & "TRIM(F4TIPDOC & '') <> '' AND "
    SqlCad = SqlCad & "TRIM(F4NUMDOC & '') <> '' "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4FECVAL, "
    SqlCad = SqlCad & "F4NUMVAL"
    
    If rstVales.State = 1 Then rstVales.Close
    
    rstVales.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstVales.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstVales)
        pgbProceso.Value = 0
        fraProceso.Caption = "Generando Registro de Compras..."
        
        Do While Not rstVales.EOF
            With objAyudaComprobante
                .inicializarEntidades
           
                .Codigo = Trim(rstVales!F4TIPDOC & "")
            
                .obtenerConfigComprobante
            End With
            
            If objAyudaComprobante.EsOficial Then
                With objAyudaVale
                    .inicializarEntidades
                    
                    .CodigoAlmacen = Trim(rstVales!f2codalm & "")
                    .NumeroVale = Trim(rstVales!F4NUMVAL & "")
                    
                    .obtenerConfigVale
                    
                    'If .RegistroCompra = vbNullString Then
                        
                        .TipoVale = Trim(rstVales!F4TIPOVALE & "")
                        .CodigoProveedor = Trim(rstVales!F2CODPROV & "")
                        .CodTipoComprobante = Trim(rstVales!F4TIPDOC & "")
                        .SerieDocumento = Trim(rstVales!F4SERDOC & "")
                        .NumeroDocumento = Trim(rstVales!F4NUMDOC & "")
                        
                        Set rstVale = .obtenerRstValeCompraPorProvYdocumento
                        Set rstValeDet = .obtenerRstValeDetalleCompraPorProvYdocumento
                        
                        If Not rstVale.EOF Then
                            rstVale.MoveFirst
                            
                            With objAyudaCompra
                                .inicializarEntidades
                                
                                If Trim(rstVale!F4REGCOM & "") <> vbNullString Then
                                    .MesMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), 1, InStr(1, Trim(rstVale!F4REGCOM & ""), "-") - 1)
                                    .NumeroMovimiento = Mid(Trim(rstVale!F4REGCOM & ""), InStr(1, Trim(rstVale!F4REGCOM & ""), "-") + 1)
                                Else
                                    If Trim(rstVales!F4FECULT & "") <> vbNullString Then
                                        .MesMovimiento = Year(CDate(Trim(rstVales!F4FECULT & ""))) & Format(Month(CDate(Trim(rstVales!F4FECULT & ""))), "00")
                                    Else
                                        .MesMovimiento = Year(CDate(Trim(rstVales!F4FECVAL & ""))) & Format(Month(CDate(Trim(rstVales!F4FECVAL & ""))), "00")
                                    End If
                                    
                                    .NumeroMovimiento = vbNullString
                                End If
                                
                                .CodProveedor = Trim(rstVale!F2CODPROV & "")
                                .NomProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2CODPROV", .CodProveedor, "T")
                                .DireccionProveedor = left(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2CODPROV", .CodProveedor, "T"), 100)
                                .TelefonoProveedor = left(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TELPROV", "EF2PROVEEDORES", "F2CODPROV", .CodProveedor, "T"), 30)
                                .TipoDocAuxiliar = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "TIPPROV", "EF2PROVEEDORES", "F2CODPROV", .CodProveedor, "T")
                                .RucProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", .CodProveedor, "T")
                                
                                    If .TipoDocAuxiliar = vbNullString Then
                                        Select Case Len(.RucProveedor)
                                            Case 8
                                                .TipoDocAuxiliar = "1"
                                            Case 11
                                                .TipoDocAuxiliar = "6"
                                            Case Else
                                                .TipoDocAuxiliar = "0"
                                        End Select
                                    End If
                                    
                                .TipoDocumento = Trim(rstVale!F4TIPDOC & "")
                                .SerieDocumento = Trim(rstVale!F4SERDOC & "")
                                .NumeroDocumento = Trim(rstVale!F4NUMDOC & "")
                                
                                .CodigoCategoria = 1
                                
                                .FechaRegistro = Format(Date, "Short Date")
                                
                                If Trim(rstVales!F4FECULT & "") <> vbNullString Then
                                    .FechaDocumento = Format(Trim(rstVale!F4FECULT & ""), "Short Date")
                                Else
                                    .FechaDocumento = Format(Trim(rstVale!F4FECVAL & ""), "Short Date")
                                End If
                                
                                .CodMoneda = Trim(rstVale!F4MONEDA & "")
                                .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .FechaDocumento, "F"))
                                .ConceptoCompra = left(Trim(rstVale!F4OBSERVA & ""), 100)
                                
                                .CodFormaPago = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2PROVEEDORES", "F2CODPROV", Trim(rstVale!F2CODPROV & ""), "T")
                                
                                With objAyudaFormaPago
                                    .inicializarEntidades
                                    
                                    .Codigo = objAyudaCompra.CodFormaPago
                                    
                                    .obtenerConfigFormaPago
                                End With
                                
                                If Val(objAyudaFormaPago.Dias) > 0 Then
                                    .FechaVencimiento = CDate(.FechaDocumento) + Val(objAyudaFormaPago.Dias)
                                Else
                                    .FechaVencimiento = .FechaDocumento
                                End If
                                
                                Select Case .CodMoneda
                                    Case "S"
                                        .CodigoGasto = "PRO"
                                        .CuentaContable = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CUENTA", "BF9GIN", "CODIGO", .CodigoGasto, "T")
                                    Case Else
                                        .CodigoGasto = "PROD"
                                        .CuentaContable = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CUENTA", "BF9GIN", "CODIGO", .CodigoGasto, "T")
                                End Select
                                
                                .PorcentajeIGV = wIgv
                                
                                .FechaReg = Format(Date, "Short Date")
                                .UsuarioReg = wusuario
                                .FechaMod = Format(Date, "Short Date")
                                .UsuarioMod = wusuario
                                
                                If .guardarCompra(True) Then
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    .SQLSelectAlter = "DELETE FROM REGISMOV WHERE F4MESMOV = '" & .MesMovimiento & "' AND F4NUMMOV = '" & .NumeroMovimiento & "'"
                                    
                                    cnn_dbbancos.Execute .SQLSelectAlter
                                    
                                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    
                                    If Not rstValeDet.EOF Then
                                        rstValeDet.MoveFirst
                                        
                                        dblItem = 0
                                        
                                        Do While Not rstValeDet.EOF
                                            .inicializarEntidadesDetalle
                                            
                                            dblItem = dblItem + 1
                                            
                                            .ITEM = dblItem
                                            
                                            .CtaContableDet = Trim(rstValeDet!CUENTA & "") 'ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CTACON", "IF5PLA", "F5CODPRO", Trim(rstValeDet!CodProducto & ""), "T")
                                            
                                            If .CtaContableDet <> vbNullString Then
                                                If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T") = vbNullString Then
                                                    With objAyudaGasto
                                                        .inicializarEntidades
                                                        
                                                        .Codigo = vbNullString
                                                        .Base = "G"
                                                        .CuentaContable = objAyudaCompra.CtaContableDet
                                                        .Descripcion = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", objAyudaCompra.CtaContableDet, "T")
                                                        .TipoGasto = "P"
                                                        '.Moneda = left(cmbMoneda.Text, 1)
                                                        .GrupoFlujo = vbNullString
                                                        
                                                        If .guardarGasto Then
                                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                                        End If
                                                        
                                                        .inicializarEntidades
                                                    End With
                                                End If
                                            End If
                                            
                                            .CodigoGastoDet = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", .CtaContableDet, "T")
                                            
                                            .NumeroOrden = Trim(rstValeDet!NROOC & "")
                                            '.CodigoProducto = Trim(rstValeDet!CodProducto & "")
                                            .ConceptoDet = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableDet, "T") 'ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", .CodigoProducto, "T")
                                            
                                            .Cantidad = 1 'Val(rstValeDet!Cantidad & "")
                                            .PrecioUnitario = Val(rstValeDet!SUBTOTAL & "") 'Val(rstValeDet!costo & "")
                                            .SubTotalDet = Val(rstValeDet!SUBTOTAL & "") 'Val(rstValeDet!TOTAL & "")
                                            .Afecto = IIf(Trim(rstValeDet!F5AFECTO & "") = "*", True, False) 'IIf(Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5AFECTO", "IF5PLA", "F5CODPRO", .CodigoProducto, "T")) = "*", True, False)
                                            
                                            .DebHab = "D"
                                            
                                            'Acumular
                                            .BaseImponible = .BaseImponible + (.SubTotalDet * IIf(.Afecto, 1, 0))
                                            .MontoInafecto = .MontoInafecto + (.SubTotalDet * IIf(.Afecto, 0, 1))
                                            .TotalIGV = .TotalIGV + (Val(rstValeDet!IGV & "") * IIf(.Afecto, 1, 0))
                                            .Descuento = .Descuento + Val(rstValeDet!DSCTO & "")
                                            .TotalFacturado = .TotalFacturado + ((.BaseImponible + .MontoInafecto + .TotalIGV) - .Descuento)
                                            
                                            'Concatenar
                                            If InStr(1, .OrdenCompra, .NumeroOrden) = 0 Then
                                                If .OrdenCompra = vbNullString Then
                                                    .OrdenCompra = .NumeroOrden
                                                Else
                                                    .OrdenCompra = .OrdenCompra & "," & .NumeroOrden
                                                End If
                                            End If
                                            
                                            .guardarCompraDetalleOneByOne
                                            
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                            
                                            rstValeDet.MoveNext
                                        Loop
                                            'ACTUALIZAR POSTERIOR A LA GRABACION
                                            .SQLSelectAlter = vbNullString
                                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4BASIMP = " & .BaseImponible & ", "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4MONINA = " & .MontoInafecto & ", "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4IGV = " & .TotalIGV & ", "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4DCTO = " & .Descuento & " "
                                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                                            
                                            cnn_dbbancos.Execute .SQLSelectAlter
                                            
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                            
                                            .SQLSelectAlter = vbNullString
                                            .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                                            .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                                            .SQLSelectAlter = .SQLSelectAlter & "SET "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4OCOMPRA = '" & left(.OrdenCompra, 255) & "', "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = (VAL(F4BASIMP & '') + VAL(F4MONINA & '') + VAL(F4IGV & '') + VAL(F4OTRIMP & '') + VAL(F4REDSUMA & '')) - (VAL(F4FONAVI & '') + VAL(F4DCTO & '') + VAL(F4MONTORET & '') + VAL(F4REDRESTA & '')) "
                                            '.SQLSelectAlter = .SQLSelectAlter & "F4TOTAL = (VAL(F4BASIMP & '') + VAL(F4MONINA & '') + VAL(F4IGV & '') + VAL(F4OTRIMP & '') + VAL(F4REDSUMA & '')) - (VAL(F4FONAVI & '') + VAL(F4DCTO & '') + VAL(F4MONTORET & '') + VAL(F4REDRESTA & '')) "
                                            .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                                            .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                                            
                                            cnn_dbbancos.Execute .SQLSelectAlter
                                            
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    End If
                                    
                                End If
                                
                                If .obtenerCompra Then
                                    
                                    With objAyudaPagDcto
                                        .Correlativo = IIf(objAyudaCompra.Correlativo = 0, -1, objAyudaCompra.Correlativo)
                                        
                                        .TipoIngreso = "1"
                                        .ITEM = 1
                                        .NumeroComprobante = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2ABREV", "DOCUMENTOS", "F2CODDOC", objAyudaCompra.TipoDocumento, "T")
                                        
                                        .NumeroComprobante = .NumeroComprobante & IIf(objAyudaCompra.SerieDocumento <> vbNullString, objAyudaCompra.SerieDocumento & "/", vbNullString) & objAyudaCompra.NumeroDocumento
                                        
                                        .FechaComprobante = objAyudaCompra.FechaDocumento
                                        .FechaVencimiento = objAyudaCompra.FechaVencimiento
                                        .CodProveedor = objAyudaCompra.CodProveedor
                                        .RucProveedor = objAyudaCompra.RucProveedor
                                        .NomProveedor = objAyudaCompra.NomProveedor
                                        .CodMoneda = objAyudaCompra.CodMoneda
                                        .TotalFacturado = objAyudaCompra.TotalFacturado
                                        .SaldoFacturado = objAyudaCompra.TotalFacturado
                                        .TipoCambio = objAyudaCompra.TipoCambio
                                        .Debe_Haber = "H"
                                        
                                        .Grupo = objAyudaCompra.CodigoGasto
                                        .CtaContable = objAyudaCompra.CuentaContable
                                        .AnnoRegCompra = left(objAyudaCompra.MesMovimiento, 4)
                                        .MovRegCompra = objAyudaCompra.NumeroMovimiento
                                        .TipoDocumento = objAyudaCompra.TipoDocumento
                                        .SerieDocumento = objAyudaCompra.SerieDocumento
                                        .NumeroDocumento = Val(objAyudaCompra.NumeroDocumento)
                                        .Notas = "CUENTA POR PAGAR GENERADA DESDE EL MODULO DE LOGISTICA."
                                        .Concepto = "INGRESO DE COMPRAS A ALMACEN."
                                        .Detalle = "INGRESO DE COMPRAS A ALMACEN."
                                        .referencia = "INGRESO DE COMPRAS A ALMACEN."
                                            
                                        If .obtenerPagDcto Then
                                            If .SaldoFacturado = .TotalFacturado Then
                                                .TotalFacturado = objAyudaCompra.TotalFacturado
                                                .SaldoFacturado = objAyudaCompra.TotalFacturado
                                            Else
                                                .SaldoFacturado = objAyudaCompra.TotalFacturado - (.TotalFacturado - .SaldoFacturado)
                                                .TotalFacturado = objAyudaCompra.TotalFacturado
                                            End If
                                        End If
                                        
                                        If .guardarPagDcto(False) Then
                                        
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                            
                                            'If objAyudaCompra.Correlativo = 0 Then
                                                'ACTUALIZAR POSTERIOR A LA GRABACION
                                                .SQLSelectAlter = vbNullString
                                                .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                                                .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                                                .SQLSelectAlter = .SQLSelectAlter & "SET "
                                                .SQLSelectAlter = .SQLSelectAlter & "F4CORRELA = " & .Correlativo & " "
                                                .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                                                .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & objAyudaCompra.MesMovimiento & "' AND "
                                                .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & objAyudaCompra.NumeroMovimiento & "'"
                                                
                                                cnn_dbbancos.Execute .SQLSelectAlter
                                                
                                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                            'End If
                                        End If
                                        
                                        .inicializarEntidades
                                    End With
                                    
                                    .ValeIngreso = vbNullString
                                    
                                    Do While Not rstVale.EOF
                                        If .ValeIngreso = vbNullString Then
                                            .ValeIngreso = Trim(rstVale!f2codalm & "") & "/" & Trim(rstVale!F4NUMVAL & "")
                                        Else
                                            .ValeIngreso = .ValeIngreso & "," & Trim(rstVale!f2codalm & "") & "/" & Trim(rstVale!F4NUMVAL & "")
                                        End If
                                        
                                        'ACTUALIZAR VALE(S) REFERENCIA DE REGISTRO DE COMPRA
                                        .SQLSelectAlter = vbNullString
                                        .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                                        .SQLSelectAlter = .SQLSelectAlter & "IF4VALES "
                                        .SQLSelectAlter = .SQLSelectAlter & "SET "
                                        .SQLSelectAlter = .SQLSelectAlter & "F4REGCOM = '" & .MesMovimiento & "-" & .NumeroMovimiento & "' "
                                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                                        .SQLSelectAlter = .SQLSelectAlter & "F2CODALM = '" & Trim(rstVale!f2codalm & "") & "' AND "
                                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMVAL = '" & Trim(rstVale!F4NUMVAL & "") & "'"
                                        
                                        cnn_dbbancos.Execute .SQLSelectAlter
                                        
                                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                        
                                        rstVale.MoveNext
                                    Loop
                                        'ACTUALIZAR POSTERIOR A LA GRABACION
                                        .SQLSelectAlter = vbNullString
                                        .SQLSelectAlter = .SQLSelectAlter & "UPDATE "
                                        .SQLSelectAlter = .SQLSelectAlter & "REGISDOC "
                                        .SQLSelectAlter = .SQLSelectAlter & "SET "
                                        .SQLSelectAlter = .SQLSelectAlter & "F4VALESING = '" & Trim(left(.ValeIngreso, 255)) & "' "
                                        .SQLSelectAlter = .SQLSelectAlter & "WHERE "
                                        .SQLSelectAlter = .SQLSelectAlter & "F4MESMOV = '" & .MesMovimiento & "' AND "
                                        .SQLSelectAlter = .SQLSelectAlter & "F4NUMMOV = '" & .NumeroMovimiento & "'"
                                        
                                        cnn_dbbancos.Execute .SQLSelectAlter
                                        
                                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                End If
                                    
                                
                                'MsgBox "Ingreso exportado a Registro de Compras:" & vbNewLine & "Mes de Registro: " & .MesMovimiento & vbNewLine & "Numero de Movimiento: " & .NumeroMovimiento, vbInformation + vbOKOnly, App.ProductName
                                
                                .inicializarEntidades
                                .inicializarEntidadesDetalle
                            End With
                        End If
                    'End If
                    
                    .inicializarEntidades
                End With
            End If
            
            objAyudaComprobante.inicializarEntidades
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Generando Registro de Compras Enero-2015... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstVales.MoveNext
        Loop
            MsgBox "Proceso Finalizado.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errGenerarRegistroCompraDeValeIngPorCompra:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    Resume
    Err.Clear
End Sub

Private Sub verificarConsistenciaStockProducto()
    On Error GoTo errVerificarConsistenciaStockProducto
    
    Dim rstProductoConMov As New ADODB.Recordset
    Dim rstValeDet As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "DET.F5CODPRO, "
    SqlCad = SqlCad & "PROD.F5NOMPRO "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES AS DET "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.F5CODPRO "
    
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F4FECVAL >= CVDATE('01/01/2014')"
    
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "DET.F5CODPRO, "
    SqlCad = SqlCad & "PROD.F5NOMPRO "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "PROD.F5NOMPRO"
    
    If rstProductoConMov.State = 1 Then rstProductoConMov.Close
    
    rstProductoConMov.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProductoConMov.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstProductoConMov)
        pgbProceso.Value = 0
        fraProceso.Caption = "Verificando consistencia de Producto..."
        
        dblCantidadProdObservado = 0
        
        Do While Not rstProductoConMov.EOF
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "DET.F2CODALM, "
            SqlCad = SqlCad & "DET.F4NUMVAL, "
            SqlCad = SqlCad & "CAB.NUMENSAM, "
            SqlCad = SqlCad & "DET.F4FECVAL, "
            SqlCad = SqlCad & "CAB.F4FECGRA, "
            SqlCad = SqlCad & "CAB.F4TIPCAM, "
            SqlCad = SqlCad & "CAB.F1CODORI, "
            SqlCad = SqlCad & "DET.F5CODPRO, "
            SqlCad = SqlCad & "DET.F2CODALM, "
            SqlCad = SqlCad & "TRIM(DET.F4NUMORD & '') AS OC, "
            SqlCad = SqlCad & "DET.COD_SOLICITUD AS PEDIDO, "
            SqlCad = SqlCad & "VAL(FORMAT((DET.F3CANPRO*IIf(TIPO='S',-1,1)),'#.00')) AS CANTIDAD, "
            SqlCad = SqlCad & "DET.F3VALVTA "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3VALES AS DET "
            SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON (CAB.F4NUMVAL=DET.F4NUMVAL) AND (CAB.F2CODALM=DET.F2CODALM) "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "CAB.F4FECVAL >= CVDATE('01/01/2014') AND "
            SqlCad = SqlCad & "DET.F5CODPRO = '" & Trim(rstProductoConMov!f5codpro & "") & "' AND "
            SqlCad = SqlCad & "CAB.F1CODORI NOT IN ('XCS') "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "DET.F4FECVAL, "
            SqlCad = SqlCad & "DET.TIPO, "
            SqlCad = SqlCad & "DET.F4NUMVAL"
            
            If rstValeDet.State = 1 Then rstValeDet.Close
            
            rstValeDet.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstValeDet.EOF Then
                rstValeDet.MoveFirst
                
                dblStock = 0
                
                Do While Not rstValeDet.EOF
                    dblStock = dblStock + Val(rstValeDet!Cantidad & "")
                    
                    If dblStock < 0 Then
                        SqlCad = vbNullString
                        SqlCad = SqlCad & "INSERT INTO TMPVALES_DET("
                        SqlCad = SqlCad & "F2CODALM, F4NUMVAL, NUMENSAM, F4FECVAL, F5CODPRO, F5NOMPRO, F3CANPRO, F3VALVTA, F3IGV"
                        SqlCad = SqlCad & ") "
                        SqlCad = SqlCad & "VALUES("
                        SqlCad = SqlCad & "'" & Trim(rstValeDet!f2codalm & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDet!F4NUMVAL & "") & "', "
                        SqlCad = SqlCad & "'" & Trim(rstValeDet!NUMENSAM & "") & "', "
                        SqlCad = SqlCad & "CVDATE('" & Trim(rstValeDet!F4FECVAL & "") & "'), "
                        SqlCad = SqlCad & "'" & Trim(rstValeDet!f5codpro & "") & "', "
                        SqlCad = SqlCad & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5NOMPRO", "IF5PLA", "F5CODPRO", Trim(rstValeDet!f5codpro & ""), "T") & "', "
                        SqlCad = SqlCad & Val(rstValeDet!Cantidad & "") & ", "
                        SqlCad = SqlCad & Val(rstValeDet!F3VALVTA & "") & ", "
                        SqlCad = SqlCad & Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COUNT(F5CODPRO) AS CANTIDAD", "IF3VALES", "F2CODALM", Trim(rstValeDet!f2codalm & ""), "T", "AND F4NUMVAL = '" & Trim(rstValeDet!F4NUMVAL & "") & "'"))
                        SqlCad = SqlCad & ")"
                        
                        abrirCnTemporal
                        
                        cnDBTemp.Execute SqlCad
                        
                        dblCantidadProdObservado = dblCantidadProdObservado + 1
                        
                        'Exit Do
                    End If
                    
                    rstValeDet.MoveNext
                Loop
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Verificando consistencia de Producto... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstProductoConMov.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se encontraron " & dblCantidadProdObservado & " Producto(s) observado(s).", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errVerificarConsistenciaStockProducto:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    Resume
    Err.Clear
End Sub

Private Sub actualizarEstadoDescargaOPs()
    On Error GoTo errActualizarEstadoDescargaOPs
    
    Dim rstValeSalida As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.F4ORDTRA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF4VALES AS CAB "
    
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.F4TIPOVALE = 'S' AND "
    SqlCad = SqlCad & "TRIM(CAB.F4ORDTRA & '') <> '' "
    
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "CAB.F4ORDTRA "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "CAB.F4ORDTRA"
    
    If rstValeSalida.State = 1 Then rstValeSalida.Close
    
    rstValeSalida.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstValeSalida.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstValeSalida)
        pgbProceso.Value = 0
        fraProceso.Caption = "Actualizando Estado de Descarga de OP..."
        
        Do While Not rstValeSalida.EOF
            
            ModMilano.actualizarEstadoDescargadoOP Trim(rstValeSalida!F4ORDTRA & ""), True
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Actualizando Estado de Descarga de OP: " & Trim(rstValeSalida!F4ORDTRA & "") & " - " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstValeSalida.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se actualizaron " & pgbProceso.Max & " OPs.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errActualizarEstadoDescargaOPs:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub correccionStockComprometidoEnNegativo()
    On Error GoTo errCorreccionStockComprometidoEnNegativo
    
    Dim rstStock As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    Dim dblItem As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F5CODPRO, "
    SqlCad = SqlCad & "COD_SOLICITUD, "
    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F2CODALM = '01' AND "
    SqlCad = SqlCad & "COD_SOLICITUD <> '' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "F5CODPRO, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "COD_SOLICITUD "
    SqlCad = SqlCad & "HAVING "
    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) < 0 "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F5CODPRO, "
    SqlCad = SqlCad & "COD_SOLICITUD"
    
    If rstStock.State = 1 Then rstStock.Close
    
    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStock.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
        pgbProceso.Value = 0
        fraProceso.Caption = "Corrigiendo Stock Comprometido en Negativo..."
        
        dblCantidadProdObservado = 0
        
        Do While Not rstStock.EOF
            
            Set objValeProcesoIni = New ClsVale
            
            With objValeProcesoIni
                .inicializarEntidades
                
                If .verificarStockProductoFisicoCorL(Trim(rstStock!f5codpro & ""), _
                                                Trim(rstStock!f2codalm & ""), _
                                                vbNullString, _
                                                vbNullString, _
                                                Abs(Val(rstStock!Cantidad & "")), _
                                                Trim(Date & "")) Then
                    
                    .inicializarEntidades
                    
                    .CodigoAlmacen = Trim(rstStock!f2codalm & "")
                    .NumeroVale = vbNullString
                    .TipoVale = "I"
                    
                    .Fecha = Format(Date, "dd/mm/yyyy")
                    .CodigoOrigen = "XCS"
                    .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                    
                        If .TipoCambio = 0 Then
                            .TipoCambio = "2.8"
                        End If
                    
                    .CodigoMoneda = "S"
                    
                    .referencia = wnomcia
                    .observaciones = "PROCESO DE CORRECCION DE STOCK COMPROMETIDO EN NEGATIVO."
                    
                    .FecReg = Format(Date, "Short Date")
                    .UsuReg = wusuario
                    .FecMod = Format(Date, "Short Date")
                    .UsuMod = wusuario
                    
                    If .guardarVale Then
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        'Borrar Detalle de Vale
                        SqlCad = vbNullString
                        SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                        
                        cnn_dbbancos.Execute SqlCad
                        Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                        
                        
                        .inicializarEntidadesDetalle
                        
                        .NumeroOrdenCompra = vbNullString
                        .Requerimiento = vbNullString
                        
                        .CodigoProducto = Trim(rstStock!f5codpro & "")
                        .CodigoProductoOriginal = Trim(rstStock!f5codpro & "")
                        .Cantidad = Abs(Val(rstStock!Cantidad & "")) * -1
                        
                        
                                                    
                        dblItem = dblItem + 1
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        .inicializarEntidadesDetalle
                        
                        dblItem = dblItem + 1
                        
                        .Requerimiento = Trim(rstStock!COD_SOLICITUD & "")
                        
                        .CodigoProducto = Trim(rstStock!f5codpro & "")
                        .CodigoProductoOriginal = Trim(rstStock!f5codpro & "")
                        .Cantidad = Abs(Val(rstStock!Cantidad & ""))
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If
                End If
            End With
            
            Set objValeProcesoIni = Nothing
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Corrigiendo Stock Comprometido en Negativo..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstStock.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se encontraron " & dblCantidadProdObservado & " Producto(s) observado(s).", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errCorreccionStockComprometidoEnNegativo:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub correccionStockComprometidoEnNegativoV2()
    On Error GoTo errCorreccionStockComprometidoEnNegativoV2
    
    Dim rstStockCEA As New ADODB.Recordset
    Dim rstStockCEAProd As New ADODB.Recordset
    Dim rstStockCEAProdPed As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    Dim dblItem As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IF3VALES.F2CODALM, "
    SqlCad = SqlCad & "IF3VALES.F5CODPRO, "
    SqlCad = SqlCad & "IF5PLA.F5NOMPRO, "
    SqlCad = SqlCad & "VAL(FORMAT( SUM(IF3VALES.F3CANPRO * IIF(IF3VALES.TIPO = 'S', -1, 1)) , '#0.0000')) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F2CODALM = '01' AND "
    SqlCad = SqlCad & "TRIM(IF3VALES.COD_SOLICITUD & '') <> '' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "IF3VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO "
    SqlCad = SqlCad & "HAVING "
    SqlCad = SqlCad & "VAL(FORMAT( SUM(IF3VALES.F3CANPRO * IIF(IF3VALES.TIPO = 'S', -1, 1)) , '#0.0000')) < 0 "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "IF5PLA.F5NOMPRO"
    
    If rstStockCEA.State = 1 Then rstStockCEA.Close
    
    rstStockCEA.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStockCEA.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStockCEA)
        pgbProceso.Value = 0
        fraProceso.Caption = "Analizando Stock Comprometido en Negativo..."
        
        dblCantidadProdObservado = 0
        
        Do While Not rstStockCEA.EOF
            
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "COD_SOLICITUD, "
            SqlCad = SqlCad & "F5CODPRO, "
            SqlCad = SqlCad & "F2CODALM "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3VALES "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "F2CODALM = '" & Trim(rstStockCEA!f2codalm & "") & "' AND "
            SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstStockCEA!f5codpro & "") & "' AND "
            SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') <> '' "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "COD_SOLICITUD, "
            SqlCad = SqlCad & "F5CODPRO, "
            SqlCad = SqlCad & "F2CODALM "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "F2CODALM, "
            SqlCad = SqlCad & "COD_SOLICITUD"
            
            If rstStockCEAProd.State = 1 Then rstStockCEAProd.Close
            
            rstStockCEAProd.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstStockCEAProd.EOF Then
                Do While Not rstStockCEAProd.EOF
                
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT "
                    SqlCad = SqlCad & "COD_SOLICITUD, "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "F2CODALM, "
                    SqlCad = SqlCad & "VAL(FORMAT( SUM(F3CANPRO * IIF(TIPO = 'S', -1, 1)) , '#0.0000')) AS CANTIDAD "
                    SqlCad = SqlCad & "FROM "
                    SqlCad = SqlCad & "IF3VALES "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F2CODALM = '" & Trim(rstStockCEAProd!f2codalm & "") & "' AND "
                    SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstStockCEAProd!f5codpro & "") & "' AND "
                    SqlCad = SqlCad & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstStockCEAProd!COD_SOLICITUD & "") & "' "
                    SqlCad = SqlCad & "GROUP BY "
                    SqlCad = SqlCad & "COD_SOLICITUD, "
                    SqlCad = SqlCad & "F5CODPRO, "
                    SqlCad = SqlCad & "F2CODALM "
                    SqlCad = SqlCad & "ORDER BY "
                    SqlCad = SqlCad & "COD_SOLICITUD"
                    
                    If rstStockCEAProdPed.State = 1 Then rstStockCEAProdPed.Close
                    
                    rstStockCEAProdPed.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstStockCEAProdPed.EOF Then
                        'Do While Not rstStockCEAProdPed.EOF
                            If Val(rstStockCEAProdPed!Cantidad & "") < 0 Then
                                Set objValeProcesoIni = New ClsVale
                                
                                With objValeProcesoIni
                                    .inicializarEntidades
                                    
                                    If .verificarStockProductoFisicoCorL(Trim(rstStockCEAProdPed!f5codpro & ""), _
                                                                    Trim(rstStockCEAProdPed!f2codalm & ""), _
                                                                    vbNullString, _
                                                                    vbNullString, _
                                                                    Abs(Val(rstStockCEAProdPed!Cantidad & "")), _
                                                                    Trim(Date & "")) Then
                                        
                                        .inicializarEntidades
                                        
                                        .CodigoAlmacen = Trim(rstStockCEAProdPed!f2codalm & "")
                                        .NumeroVale = vbNullString
                                        .TipoVale = "I"
                                        
                                        .Fecha = Format(Date, "dd/mm/yyyy")
                                        .CodigoOrigen = "XCS"
                                        .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                                        
                                            If .TipoCambio = 0 Then
                                                .TipoCambio = "2.8"
                                            End If
                                        
                                        .CodigoMoneda = "S"
                                        
                                        .referencia = wnomcia
                                        .observaciones = "PROCESO DE REGULARIZACION DE STOCK COMPROMETIDO EN NEGATIVO."
                                        
                                        .FecReg = Format(Date, "Short Date")
                                        .UsuReg = wusuario
                                        .FecMod = Format(Date, "Short Date")
                                        .UsuMod = wusuario
                                        
                                        If .guardarVale Then
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                            
                                            'Borrar Detalle de Vale
                                            SqlCad = vbNullString
                                            SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                                            
                                            cnn_dbbancos.Execute SqlCad
                                            Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                                            
                                            
                                            .inicializarEntidadesDetalle
                                            
                                            .NumeroOrdenCompra = vbNullString
                                            .Requerimiento = vbNullString
                                            
                                            .CodigoProducto = Trim(rstStockCEAProdPed!f5codpro & "")
                                            .CodigoProductoOriginal = Trim(rstStockCEAProdPed!f5codpro & "")
                                            .Cantidad = Abs(Val(rstStockCEAProdPed!Cantidad & "")) * -1
                                            
                                            
                                                                        
                                            dblItem = dblItem + 1
                                            
                                            .ITEM = dblItem
                                            
                                            .guardarValeDetalleOneByOne
                                            
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                            
                                            .inicializarEntidadesDetalle
                                            
                                            dblItem = dblItem + 1
                                            
                                            .Requerimiento = Trim(rstStockCEAProdPed!COD_SOLICITUD & "")
                                            
                                            .CodigoProducto = Trim(rstStockCEAProdPed!f5codpro & "")
                                            .CodigoProductoOriginal = Trim(rstStockCEAProdPed!f5codpro & "")
                                            .Cantidad = Abs(Val(rstStockCEAProdPed!Cantidad & ""))
                                            .ITEM = dblItem
                                            
                                            .guardarValeDetalleOneByOne
                                            
                                            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                        End If
                                    Else
                                        Actualiza_Log "No se logro regularizar Stock, ya que no se contaba con Stock Libre Disponible para comprometer. Datos de Regularizacion Trunca: [Producto: " & Trim(rstStockCEAProdPed!f5codpro & "") & " - " & Trim(rstStockCEA!F5NOMPRO & "") & "]     [Cod. Almacen: " & Trim(rstStockCEAProdPed!f2codalm & "") & "]      [No. Pedido: " & Trim(rstStockCEAProdPed!COD_SOLICITUD & "") & "]   [Cantidad a Redistribuir: " & Format(Abs(Val(rstStockCEAProdPed!Cantidad & "")), "#.0000") & "]", StrConexDbBancos
                                    End If
                                End With
                                
                                Set objValeProcesoIni = Nothing
                            End If
                            
                        '    rstStockCEAProdPed.MoveNext
                        'Loop
                    End If
                    
                    rstStockCEAProd.MoveNext
                Loop
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Corrigiendo Stock Comprometido en Negativo..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstStockCEA.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se encontraron " & dblCantidadProdObservado & " Producto(s) observado(s).", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errCorreccionStockComprometidoEnNegativoV2:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub analisisYCorreccionStockComprometidoActual()
    On Error GoTo errAnalisisYCorreccionStockComprometidoActual
    
    Dim rstStock As New ADODB.Recordset
    Dim rstProduccion As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    
    Dim dblLiberarCompromiso As Double
    
    Dim dblItem As Double
    
    'Descargar en el Temporal los compromisos actuales, segun Ordenes de Produccion
    
    Screen.MousePointer = vbHourglass
           
    'objAyudaSolicitud.listarGrillaResumenRequerimiento dbgResumen, Nothing, strNroPedido, strCodProducto
    If Not ModMilano.importarResumenRequerimientoProduccion(fraProceso, pgbProceso) Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    
    
    'Seleccionar Stock Comprometido en Almacen Actual para su verificacion
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F5CODPRO, "
    SqlCad = SqlCad & "COD_SOLICITUD, "
    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F2CODALM = '01' AND "
    SqlCad = SqlCad & "COD_SOLICITUD <> '' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "F5CODPRO, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "COD_SOLICITUD "
    SqlCad = SqlCad & "HAVING "
    SqlCad = SqlCad & "VAL(FORMAT( SUM(VAL(FORMAT(VAL(F3CANPRO & ''), '#0.00')) * IIF(TIPO = 'S', -1, 1)) , '#0.00')) > 0 "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F5CODPRO, "
    SqlCad = SqlCad & "COD_SOLICITUD"
    
    If rstStock.State = 1 Then rstStock.Close
    
    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStock.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
        pgbProceso.Value = 0
        fraProceso.Caption = "Evaluando Stock Comprometido Actual..."
        
        dblCantidadProdObservado = 0
        
        Do While Not rstStock.EOF
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "RES.NROPEDIDO, "
            SqlCad = SqlCad & "RES.CODPRODUCTO, "
            SqlCad = SqlCad & "SUM(RES.SALDO) AS SALDO "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "TMPUTILRESUMENREQUERIMIENTOPRODUCCION AS RES "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "RES.NROPEDIDO = '" & Trim(rstStock!COD_SOLICITUD & "") & "' AND "
            SqlCad = SqlCad & "RES.CODPRODUCTO = '" & Trim(rstStock!f5codpro & "") & "' "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "RES.NROPEDIDO, "
            SqlCad = SqlCad & "RES.CODPRODUCTO "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "RES.NROPEDIDO, "
            SqlCad = SqlCad & "RES.CODPRODUCTO"
            
            If rstProduccion.State = 1 Then rstProduccion.Close
            
            rstProduccion.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
            
            If Not rstProduccion.EOF Then
                Select Case Val(rstProduccion!SALDO & "")
                    Case Is >= Val(rstStock!Cantidad & "")
                        dblLiberarCompromiso = 0
                    Case Else
                        dblLiberarCompromiso = Val(rstStock!Cantidad & "") - Val(rstProduccion!SALDO & "")
                End Select
            Else
                dblLiberarCompromiso = Val(rstStock!Cantidad & "")
            End If
            
            
            If dblLiberarCompromiso > 0 Then
                Set objValeProcesoIni = New ClsVale
                
                With objValeProcesoIni
                
                    .inicializarEntidades
                    
                    .CodigoAlmacen = Trim(rstStock!f2codalm & "")
                    .NumeroVale = vbNullString
                    .TipoVale = "I"
                    
                    .Fecha = Format(Date, "dd/mm/yyyy")
                    .CodigoOrigen = "XCS"
                    .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
                    
                        If .TipoCambio = 0 Then
                            .TipoCambio = "2.8"
                        End If
                    
                    .CodigoMoneda = "S"
                    
                    .referencia = wnomcia
                    .observaciones = "PROCESO DE CORRECCION DE STOCK COMPROMETIDO."
                    
                    .FecReg = Format(Date, "Short Date")
                    .UsuReg = wusuario
                    .FecMod = Format(Date, "Short Date")
                    .UsuMod = wusuario
                    
                    If .guardarVale Then
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        'Borrar Detalle de Vale
                        SqlCad = vbNullString
                        SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                        
                        cnn_dbbancos.Execute SqlCad
                        Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                        
                        
                        .inicializarEntidadesDetalle
                        
                        .NumeroOrdenCompra = vbNullString
                        .Requerimiento = Trim(rstStock!COD_SOLICITUD & "")
                        
                        .CodigoProducto = Trim(rstStock!f5codpro & "")
                        .CodigoProductoOriginal = Trim(rstStock!f5codpro & "")
                        .Cantidad = dblLiberarCompromiso * -1
                        
                                                    
                        dblItem = dblItem + 1
                        
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        .inicializarEntidadesDetalle
                        
                        dblItem = dblItem + 1
                        
                        .Requerimiento = vbNullString
                        
                        .CodigoProducto = Trim(rstStock!f5codpro & "")
                        .CodigoProductoOriginal = Trim(rstStock!f5codpro & "")
                        .Cantidad = dblLiberarCompromiso
                        .ITEM = dblItem
                        
                        .guardarValeDetalleOneByOne
                        
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    End If
                End With
                
                Set objValeProcesoIni = Nothing
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Evaluando Stock Comprometido Actual..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstStock.MoveNext
        Loop
            MsgBox "Proceso Finalizado.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errAnalisisYCorreccionStockComprometidoActual:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub correccionCostoPromedioEnValesAjusteDic14()
    On Error GoTo errCorreccionCostoPromedioEnValesAjusteDic14
    
    Dim rstStock As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    Dim strTipoMovimiento As String
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.F4TIPOVALE, "
    SqlCad = SqlCad & "CAB.F2CODALM, "
    SqlCad = SqlCad & "CAB.F4NUMVAL, "
    SqlCad = SqlCad & "CAB.F4FECVAL, "
    SqlCad = SqlCad & "CAB.F4TIPCAM, "
    SqlCad = SqlCad & "CAB.NUMENSAM, "
    SqlCad = SqlCad & "DET.F5CODPRO, "
    SqlCad = SqlCad & "DET.F3CANPRO, "
    SqlCad = SqlCad & "DET.F3VALVTA, "
    SqlCad = SqlCad & "DET.F3IGV, "
    SqlCad = SqlCad & "DET.F3TOTITE, "
    SqlCad = SqlCad & "DET.F3VALDOL, "
    SqlCad = SqlCad & "DET.F3IGVDOL, "
    SqlCad = SqlCad & "DET.F3TOTDOL "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES AS DET "
    SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.F2CODALM = '01' AND "
    SqlCad = SqlCad & "CAB.F4NUMVAL IN ('I-2014120014', 'S-2014120018') AND "
    SqlCad = SqlCad & "DET.F3VALVTA = 0"
    
    If rstStock.State = 1 Then rstStock.Close
    
    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStock.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
        pgbProceso.Value = 0
        fraProceso.Caption = "Corrigiendo Costo Promedio en Vales de Ajuste Dic-14..."
        
        dblCantidadProdObservado = 0
        
        'abrirCnDBMilano
        
        Do While Not rstStock.EOF
            With objAyudaVale
                .inicializarEntidades
                .inicializarEntidadesDetalle
                
                .TipoCambio = Val(rstStock!F4TIPCAM & "")
                
                .Cantidad = Val(rstStock!F3CANPRO & "")
                
                .ValorVenta = ModMilano.devolverUltimoCostoCompraEnIntegradoEnSoles(Trim(rstStock!f5codpro & ""))
                
                If .ValorVenta > 0 Then
                    .IGV = 0
                    .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                    
                    .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                    .IgvDol = 0
                    .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "IF3VALES "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "F3VALVTA = " & .ValorVenta & ", "
                    SqlCad = SqlCad & "F3IGV = " & .IGV & ", "
                    SqlCad = SqlCad & "F3TOTITE = " & .TOTAL & ", "
                    SqlCad = SqlCad & "F3VALDOL = " & .ValorVentaDol & ", "
                    SqlCad = SqlCad & "F3IGVDOL = " & .IgvDol & ", "
                    SqlCad = SqlCad & "F3TOTDOL = " & .TotalDol & ", "
                    SqlCad = SqlCad & "OBSERVACIONES = 'ACTUALIZACION DE COSTO EN BASE A LA ULTIMA COMPRA REGISTRADA.' "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F2CODALM = '" & Trim(rstStock!f2codalm & "") & "' AND "
                    SqlCad = SqlCad & "F4NUMVAL = '" & Trim(rstStock!F4NUMVAL & "") & "' AND "
                    SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstStock!f5codpro & "") & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
                    
                    Select Case Trim(rstStock!F4TIPOVALE & "")
                        Case "I"
                            strTipoMovimiento = "INGRESO"
                        Case "S"
                            strTipoMovimiento = "SALIDA"
                    End Select
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & strTipoMovimiento & "DETALLE "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "COSTO = " & .ValorVenta & ", "
                    SqlCad = SqlCad & "IMPORTE = " & .TOTAL & " "
                        
                        If Trim(rstStock!F4TIPOVALE & "") = "I" Then
                            SqlCad = SqlCad & ",OBSERVACION = 'ACTUALIZACION DE COSTO EN BASE A LA ULTIMA COMPRA REGISTRADA.' "
                        End If
                        
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "ID" & strTipoMovimiento & " = '" & Trim(rstStock!NUMENSAM & "") & "' AND "
                    SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstStock!f5codpro & "") & "'"
                    
                    cnBdStudioModa.Execute SqlCad
                    
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                End If
            End With
            
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Corrigiendo Costo Promedio en Vales de Ajuste Dic-14..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstStock.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se encontraron " & dblCantidadProdObservado & " Producto(s) observado(s).", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errCorreccionCostoPromedioEnValesAjusteDic14:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

'Private Sub correccionCantidadEnValesAjusteDic14()
'    On Error GoTo errCorreccionCantidadEnValesAjusteDic14
'
'    Dim rstStock As New ADODB.Recordset
'
'    Dim dblStock As Double
'    Dim dblCantidadProdObservado As Double
'    Dim strTipoMovimiento As String
'    Dim strFiltroIdInsumo As String
'
'    With objAyudaVale
'        .inicializarEntidades
'
'        .listarGrillaInventarioProductoV2 Nothing, _
'                                             vbNullString, _
'                                            "31/01/2015", _
'                                            False, _
'                                            "S", _
'                                            False, _
'                                            "01", _
'                                            vbNullString, _
'                                            vbNullString, _
'                                            Nothing, _
'                                            True, _
'                                            vbNullString, _
'                                            False
'
'        SqlCad = vbNullString
'        SqlCad = .SQLSelectAlter
'    End With
'
'    If rstStock.State = 1 Then rstStock.Close
'
'    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstStock.EOF Then
'        fraProceso.Visible = True
'        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
'        pgbProceso.Value = 0
'        fraProceso.Caption = "Armando filtro de Productos..."
'
'        strFiltroIdInsumo = vbNullString
'
'        Do While Not rstStock.EOF
'            If strFiltroIdInsumo = vbNullString Then
'                strFiltroIdInsumo = "'" & Trim(rstStock!f5codpro & "") & "'"
'            Else
'                strFiltroIdInsumo = strFiltroIdInsumo & ", '" & Trim(rstStock!f5codpro & "") & "'"
'            End If
'
'            DoEvents
'
'            pgbProceso.Value = pgbProceso.Value + 1
'            fraProceso.Caption = "Armando filtro de Productos..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
'
'            rstStock.MoveNext
'        Loop
'    End If
'
'    SqlCad = vbNullString
'    SqlCad = SqlCad & "SELECT "
'    SqlCad = SqlCad & "CAB.F4TIPOVALE, "
'    SqlCad = SqlCad & "CAB.F2CODALM, "
'    SqlCad = SqlCad & "CAB.F4NUMVAL, "
'    SqlCad = SqlCad & "CAB.F4FECVAL, "
'    SqlCad = SqlCad & "CAB.F4TIPCAM, "
'    SqlCad = SqlCad & "CAB.NUMENSAM, "
'    SqlCad = SqlCad & "DET.F5CODPRO, "
'    SqlCad = SqlCad & "DET.F3CANPRO, "
'    SqlCad = SqlCad & "DET.F3VALVTA, "
'    SqlCad = SqlCad & "DET.F3IGV, "
'    SqlCad = SqlCad & "DET.F3TOTITE, "
'    SqlCad = SqlCad & "DET.F3VALDOL, "
'    SqlCad = SqlCad & "DET.F3IGVDOL, "
'    SqlCad = SqlCad & "DET.F3TOTDOL "
'    SqlCad = SqlCad & "FROM "
'    SqlCad = SqlCad & "IF3VALES AS DET "
'    SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
'    SqlCad = SqlCad & "WHERE "
'    SqlCad = SqlCad & "CAB.F2CODALM = '01' AND "
'    SqlCad = SqlCad & "CAB.F4TIPOVALE = 'S' AND "
'    SqlCad = SqlCad & "CAB.F4NUMVAL IN ('I-2014120014', 'S-2014120018') AND "
'    SqlCad = SqlCad & "DET.F5CODPRO IN (" & strFiltroIdInsumo & ")"
'
'    If rstStock.State = 1 Then rstStock.Close
'
'    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstStock.EOF Then
'        fraProceso.Visible = True
'        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
'        pgbProceso.Value = 0
'        fraProceso.Caption = "Corrigiendo Cantidad en Vales de Ajuste Dic-14..."
'
'        dblCantidadProdObservado = 0
'
'        'abrirCnDBMilano
'
'        Do While Not rstStock.EOF
'            With objAyudaVale
'                .inicializarEntidades
'                .inicializarEntidadesDetalle
'
'                .Fecha = "31/01/2015"
'                .TipoCambio = Val(rstStock!F4TIPCAM & "")
'
'                .CodigoProducto = Trim(rstStock!f5codpro & "")
'                .Cantidad = Val(rstStock!F3CANPRO & "")
'                .CantidadMaxima = Val(.devuelveStockFisicoDeProducto(vbNullString, False))
'
'                If .Cantidad + .CantidadMaxima = 0 Then
'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "DELETE "
'                    SqlCad = SqlCad & "FROM "
'                    SqlCad = SqlCad & "IF3VALES "
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "F2CODALM = '" & Trim(rstStock!f2codalm & "") & "' AND "
'                    SqlCad = SqlCad & "F4NUMVAL = '" & Trim(rstStock!F4NUMVAL & "") & "' AND "
'                    SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstStock!f5codpro & "") & "'"
'
'                    cnn_dbbancos.Execute SqlCad
'
'                    Actualiza_Log SqlCad, StrConexDbBancos
'
'                    Select Case Trim(rstStock!F4TIPOVALE & "")
'                        Case "I"
'                            strTipoMovimiento = "INGRESO"
'                        Case "S"
'                            strTipoMovimiento = "SALIDA"
'                    End Select
'
'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "DELETE "
'                    SqlCad = SqlCad & "FROM "
'                    SqlCad = SqlCad & strTipoMovimiento & "DETALLE "
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "ID" & strTipoMovimiento & " = '" & Trim(rstStock!NUMENSAM & "") & "' AND "
'                    SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstStock!f5codpro & "") & "'"
'
'                    cnBdStudioModa.Execute SqlCad
'
'                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
'                ElseIf .Cantidad + .CantidadMaxima > 0 Then
'                    .Cantidad = .Cantidad + .CantidadMaxima
'                    .ValorVenta = Val(rstStock!F3VALVTA & "")
'                    .IGV = 0
'                    .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
'
'                    .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
'                    .IgvDol = 0
'                    .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
'
'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "UPDATE "
'                    SqlCad = SqlCad & "IF3VALES "
'                    SqlCad = SqlCad & "SET "
'                    SqlCad = SqlCad & "F3CANPRO = " & .Cantidad & ", "
'                    SqlCad = SqlCad & "F3VALVTA = " & .ValorVenta & ", "
'                    SqlCad = SqlCad & "F3IGV = " & .IGV & ", "
'                    SqlCad = SqlCad & "F3TOTITE = " & .TOTAL & ", "
'                    SqlCad = SqlCad & "F3VALDOL = " & .ValorVentaDol & ", "
'                    SqlCad = SqlCad & "F3IGVDOL = " & .IgvDol & ", "
'                    SqlCad = SqlCad & "F3TOTDOL = " & .TotalDol & ", "
'                    SqlCad = SqlCad & "OBSERVACIONES = 'ACTUALIZACION DE COSTO EN BASE A LA ULTIMA COMPRA REGISTRADA.' "
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "F2CODALM = '" & Trim(rstStock!f2codalm & "") & "' AND "
'                    SqlCad = SqlCad & "F4NUMVAL = '" & Trim(rstStock!F4NUMVAL & "") & "' AND "
'                    SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstStock!f5codpro & "") & "'"
'
'                    cnn_dbbancos.Execute SqlCad
'
'                    Actualiza_Log SqlCad, StrConexDbBancos
'
'                    Select Case Trim(rstStock!F4TIPOVALE & "")
'                        Case "I"
'                            strTipoMovimiento = "INGRESO"
'                        Case "S"
'                            strTipoMovimiento = "SALIDA"
'                    End Select
'
'                    SqlCad = vbNullString
'                    SqlCad = SqlCad & "UPDATE "
'                    SqlCad = SqlCad & strTipoMovimiento & "DETALLE "
'                    SqlCad = SqlCad & "SET "
'                    SqlCad = SqlCad & "CANTIDAD = " & .Cantidad & ", "
'                    SqlCad = SqlCad & "COSTO = " & .ValorVenta & ", "
'                    SqlCad = SqlCad & "IMPORTE = " & .TOTAL & " "
'
'                        If Trim(rstStock!F4TIPOVALE & "") = "I" Then
'                            SqlCad = SqlCad & ",OBSERVACION = 'ACTUALIZACION DE COSTO EN BASE A LA ULTIMA COMPRA REGISTRADA.' "
'                        End If
'
'                    SqlCad = SqlCad & "WHERE "
'                    SqlCad = SqlCad & "ID" & strTipoMovimiento & " = '" & Trim(rstStock!NUMENSAM & "") & "' AND "
'                    SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstStock!f5codpro & "") & "'"
'
'                    cnBdStudioModa.Execute SqlCad
'
'                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
'                End If
'            End With
'
'
'            DoEvents
'
'            pgbProceso.Value = pgbProceso.Value + 1
'            fraProceso.Caption = "Corrigiendo Cantidad en Vales de Ajuste Dic-14..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
'
'            rstStock.MoveNext
'        Loop
'            MsgBox "Proceso Finalizado, se encontraron " & dblCantidadProdObservado & " Producto(s) observado(s).", vbInformation + vbOKOnly, App.ProductName
'    End If
'
'    Exit Sub
'errCorreccionCantidadEnValesAjusteDic14:
'    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
'    'Resume
'    Err.Clear
'End Sub
'
Private Sub correccionCostoPromedioEnValesAjusteDic14paso2()
    On Error GoTo errCorreccionCostoPromedioEnValesAjusteDic14paso2
    
    Dim rstStock As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    Dim strTipoMovimiento As String
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "CAB.F4TIPOVALE, "
    SqlCad = SqlCad & "CAB.F2CODALM, "
    SqlCad = SqlCad & "CAB.F4NUMVAL, "
    SqlCad = SqlCad & "CAB.F4FECVAL, "
    SqlCad = SqlCad & "CAB.F4TIPCAM, "
    SqlCad = SqlCad & "CAB.NUMENSAM, "
    SqlCad = SqlCad & "DET.F5CODPRO, "
    SqlCad = SqlCad & "DET.F3CANPRO, "
    SqlCad = SqlCad & "DET.F3VALVTA, "
    SqlCad = SqlCad & "DET.F3IGV, "
    SqlCad = SqlCad & "DET.F3TOTITE, "
    SqlCad = SqlCad & "DET.F3VALDOL, "
    SqlCad = SqlCad & "DET.F3IGVDOL, "
    SqlCad = SqlCad & "DET.F3TOTDOL "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3VALES AS DET "
    SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "CAB.F2CODALM = '01' AND "
    SqlCad = SqlCad & "CAB.F4NUMVAL IN ('I-2014120014', 'S-2014120018') AND "
    SqlCad = SqlCad & "DET.F3VALVTA = 0"
    
    If rstStock.State = 1 Then rstStock.Close
    
    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStock.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
        pgbProceso.Value = 0
        fraProceso.Caption = "Corrigiendo Costo Promedio en Vales de Ajuste Dic-14 - Paso 2..."
        
        dblCantidadProdObservado = 0
        
        'abrirCnDBMilano
        
        Do While Not rstStock.EOF
            With objAyudaVale
                .inicializarEntidades
                .inicializarEntidadesDetalle
                
                .TipoCambio = Val(rstStock!F4TIPCAM & "")
                
                .Cantidad = Val(rstStock!F3CANPRO & "")
                
                abrirCnTemporal
                
                .ValorVenta = ModUtilitario.ObtenerCampoV2(cnDBTemp, "COSTO", "TMPCOSTOINSUMO", "CODIGO", Trim(rstStock!f5codpro & ""), "T", _
                                                            "AND ALMACEN = '" & Trim(rstStock!f2codalm & "") & "' " & _
                                                            "AND NUMEROVALE = '" & Trim(rstStock!F4NUMVAL & "") & "'") 'ModMilano.devolverUltimoCostoCompraEnIntegradoEnSoles(Trim(rstStock!f5codpro & ""))
                
                If .ValorVenta > 0 Then
                    .IGV = 0
                    .TOTAL = Val(Format(.ValorVenta * .Cantidad, "#0.00"))
                    
                    .ValorVentaDol = Val(Format(.ValorVenta / .TipoCambio, "#0.00"))
                    .IgvDol = 0
                    .TotalDol = Val(Format(.TOTAL / .TipoCambio, "#0.00"))
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "IF3VALES "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "F3VALVTA = " & .ValorVenta & ", "
                    SqlCad = SqlCad & "F3IGV = " & .IGV & ", "
                    SqlCad = SqlCad & "F3TOTITE = " & .TOTAL & ", "
                    SqlCad = SqlCad & "F3VALDOL = " & .ValorVentaDol & ", "
                    SqlCad = SqlCad & "F3IGVDOL = " & .IgvDol & ", "
                    SqlCad = SqlCad & "F3TOTDOL = " & .TotalDol & ", "
                    SqlCad = SqlCad & "OBSERVACIONES = 'ACTUALIZACION DE COSTO EN BASE A EXCEL ENVIADO POR CORREO EL 6/3/2015.' "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F2CODALM = '" & Trim(rstStock!f2codalm & "") & "' AND "
                    SqlCad = SqlCad & "F4NUMVAL = '" & Trim(rstStock!F4NUMVAL & "") & "' AND "
                    SqlCad = SqlCad & "F5CODPRO = '" & Trim(rstStock!f5codpro & "") & "'"
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
                    
                    Select Case Trim(rstStock!F4TIPOVALE & "")
                        Case "I"
                            strTipoMovimiento = "INGRESO"
                        Case "S"
                            strTipoMovimiento = "SALIDA"
                    End Select
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & strTipoMovimiento & "DETALLE "
                    SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "COSTO = " & .ValorVenta & ", "
                    SqlCad = SqlCad & "IMPORTE = " & .TOTAL & " "
                        
                        If Trim(rstStock!F4TIPOVALE & "") = "I" Then
                            SqlCad = SqlCad & ",OBSERVACION = 'ACTUALIZACION DE COSTO EN BASE A EXCEL ENVIADO POR CORREO EL 6/3/2015.' "
                        End If
                        
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "ID" & strTipoMovimiento & " = '" & Trim(rstStock!NUMENSAM & "") & "' AND "
                    SqlCad = SqlCad & "IDINSUMO = '" & Trim(rstStock!f5codpro & "") & "'"
                    
                    'abrirCnDBMilano
                    
                    cnBdStudioModa.Execute SqlCad
                    
                    Actualiza_Log "< DB Externo > " & SqlCad, StrConexDbBancos
                End If
            End With
            
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Corrigiendo Costo Promedio en Vales de Ajuste Dic-14 - Paso 2..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstStock.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se encontraron " & dblCantidadProdObservado & " Producto(s) observado(s).", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errCorreccionCostoPromedioEnValesAjusteDic14paso2:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub descarteAutomaticoDeReposicionCompromiso()
    On Error GoTo errDescarteAutomaticoDeReposicionCompromiso
    
    Dim rstStock As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadDescartados As Double
    Dim strTipoMovimiento As String
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "REP.IDREPOSICION, "
    SqlCad = SqlCad & "REP.NROPEDIDO, "
    SqlCad = SqlCad & "PED.CS_FECHA, "
    SqlCad = SqlCad & "PED.CS_FENTREGA, "
    SqlCad = SqlCad & "PED.CS_NOMREF, "
    SqlCad = SqlCad & "PED.CS_CODSOLICITANTE, "
    SqlCad = SqlCad & "REP.IDINSUMO, "
    SqlCad = SqlCad & "PROD.F5NOMPRO, "
    SqlCad = SqlCad & "REP.CANTIDAD, "
    SqlCad = SqlCad & "REP.DESCARTARREPOSICION, "
    SqlCad = SqlCad & "REP.REPOSICIONENCURSO, "
    SqlCad = SqlCad & "REP.OBSERVACION "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "(SF1REPOSICIONCOMPROMISO AS REP "
    SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS PED ON PED.COD_SOLICITUD = REP.NROPEDIDO) "
    SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = REP.IDINSUMO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "REP.DESCARTARREPOSICION = FALSE AND "
    SqlCad = SqlCad & "REP.REPOSICIONENCURSO = FALSE "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "PED.CS_FENTREGA, "
    SqlCad = SqlCad & "REP.NROPEDIDO"
    
    If rstStock.State = 1 Then rstStock.Close
    
    rstStock.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStock.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStock)
        pgbProceso.Value = 0
        fraProceso.Caption = "Descarte automatico de Reposicion de Compromiso..."
        
        dblCantidadDescartados = 0
        
        'abrirCnDBMilano
        
        Do While Not rstStock.EOF
            
            dblCantidad = ModMilano.devolverSaldoRequerimientoProduccion(Trim(rstStock!NroPedido & ""), Trim(rstStock!IDINSUMO & ""))
            
            If dblCantidad <= 0 Then
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                SqlCad = SqlCad & "SF1REPOSICIONCOMPROMISO "
                SqlCad = SqlCad & "SET "
                SqlCad = SqlCad & "DESCARTARREPOSICION = TRUE, "
                SqlCad = SqlCad & "USUDESCARTE = '" & wusuario & "', "
                SqlCad = SqlCad & "FECDESCARTE = CVDATE('" & Now & "') "
                SqlCad = SqlCad & "WHERE "
                SqlCad = SqlCad & "IDREPOSICION = " & Trim(rstStock!IDREPOSICION & "")
                
                cnn_dbbancos.Execute SqlCad
        
                Actualiza_Log SqlCad, StrConexDbBancos
                
                dblCantidadDescartados = dblCantidadDescartados + 1
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descarte automatico de Reposicion de Compromiso..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Descartados: " & dblCantidadDescartados & " / " & pgbProceso.Max
            
            rstStock.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se descartaron " & dblCantidadDescartados & " Producto(s).", vbInformation + vbOKOnly, App.ProductName
    End If
    
    Exit Sub
errDescarteAutomaticoDeReposicionCompromiso:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub corregirDuplicadosEnDetalleOC()
    On Error GoTo errCorregirDuplicadosEnDetalleOC
    
    Dim rstDuplicadoOC As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblOCsCorregidas As Double
    Dim strTipoMovimiento As String
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F4LOCAL, F4NUMORD, COD_SOLICITUD, F3CODPRO, COUNT(*) AS DUPLICADO, F5NOMPRO, F5NOMPRO_ING, "
    SqlCad = SqlCad & "UNIDAD, SUM(VAL(FORMAT(F3CANPRO, '#0.00'))) AS CANTIDAD, SUM(VAL(FORMAT(F3CANPRO2, '#0.00'))) AS CANTIDAD2, SUM(VAL(FORMAT(F3CANFAL, '#0.00'))) AS CANTIDADFAL, "
    SqlCad = SqlCad & "F3PORCDEMASIA, F3PRECOS, F5AFECTO, F3PORDCT, F3AJUSTE "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF3ORDEN "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F4LOCAL = 'OC' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "F4LOCAL, F4NUMORD, COD_SOLICITUD, F3CODPRO, F5NOMPRO, F5NOMPRO_ING, "
    SqlCad = SqlCad & "UNIDAD, F3PORCDEMASIA, F3PRECOS, F5AFECTO, F3PORDCT, F3AJUSTE "
    SqlCad = SqlCad & "HAVING "
    SqlCad = SqlCad & "COUNT(*) > 1 "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4NUMORD, COD_SOLICITUD, F3CODPRO"
    
    If rstDuplicadoOC.State = 1 Then rstDuplicadoOC.Close
    
    rstDuplicadoOC.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDuplicadoOC.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstDuplicadoOC)
        pgbProceso.Value = 0
        fraProceso.Caption = "Descargando Duplicados en Detalle de OC's..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM TMPUTILSTOCKORDENDET"
        
        abrirCnTemporal
            
        cnDBTemp.Execute SqlCad
        
        Do While Not rstDuplicadoOC.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO TMPUTILSTOCKORDENDET("
            SqlCad = SqlCad & "F4LOCAL, F4NUMORD, COD_SOLICITUD, F3CODPRO, F5NOMPRO, F5NOMPRO_ING, UNIDAD, "
            SqlCad = SqlCad & "F3CANPRO, F3CANPRO2, F3CANFAL, F3PORCDEMASIA, F3PRECOS, F5AFECTO, F3PORDCT, F3AJUSTE) "
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!F4LOCAL & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!F4NUMORD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!COD_SOLICITUD & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!F3CODPRO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!F5NOMPRO & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!F5NOMPRO_ING & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!UNIDAD & "") & "', "
            SqlCad = SqlCad & Val(rstDuplicadoOC!Cantidad & "") & ", "
            SqlCad = SqlCad & Val(rstDuplicadoOC!Cantidad2 & "") & ", "
            SqlCad = SqlCad & Val(rstDuplicadoOC!CANTIDADFAL & "") & ", "
            SqlCad = SqlCad & Val(rstDuplicadoOC!F3PORCDEMASIA & "") & ", "
            SqlCad = SqlCad & Val(rstDuplicadoOC!F3PRECOS & "") & ", "
            SqlCad = SqlCad & "'" & Trim(rstDuplicadoOC!F5AFECTO & "") & "', "
            SqlCad = SqlCad & Val(rstDuplicadoOC!F3PORDCT & "") & ", "
            SqlCad = SqlCad & Val(rstDuplicadoOC!F3AJUSTE & "")
            SqlCad = SqlCad & ")"
            
            abrirCnTemporal
            
            cnDBTemp.Execute SqlCad
            
            dblOCsCorregidas = dblOCsCorregidas + 1
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Descargando Duplicados en Detalle de OC's......" & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Duplicados: " & dblOCsCorregidas & " / " & pgbProceso.Max
            
            rstDuplicadoOC.MoveNext
        Loop
    End If
    
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT * FROM TMPUTILSTOCKORDENDET"
    
    If rstDuplicadoOC.State = 1 Then rstDuplicadoOC.Close
    
    rstDuplicadoOC.Open SqlCad, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDuplicadoOC.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstDuplicadoOC)
        pgbProceso.Value = 0
        fraProceso.Caption = "Corrección de Duplicados en Detalle de OC's..."
        
        dblOCsCorregidas = 0
        
        Do While Not rstDuplicadoOC.EOF
            With objAyudaOrden
                .inicializarEntidadesDetalle
                
                .TipoOrden = Trim(rstDuplicadoOC!F4LOCAL & "")
                .NumeroOrden = Trim(rstDuplicadoOC!F4NUMORD & "")
                .CodigoProducto = Trim(rstDuplicadoOC!F3CODPRO & "")
                .Requerimiento = Trim(rstDuplicadoOC!COD_SOLICITUD & "")
                
                .obtenerConfigOrdenDetalleOnebyOne
                
                If .ITEM <> 0 Then
                    .PorcentajeImpuesto = wwigv / 100
                    .SignoImpuesto = 1
                    
                    .Cantidad = Val(rstDuplicadoOC!F3CANPRO & "")
                    .CantidadMaxima = Val(rstDuplicadoOC!F3CANPRO2 & "")
                    .CantidadFaltante = Val(rstDuplicadoOC!F3CANFAL & "")
                    
                    .PorcentajeDemasia = Val(rstDuplicadoOC!F3PORCDEMASIA & "") / 100
                    .PrecioSinImpuesto = Val(rstDuplicadoOC!F3PRECOS & "")
                    .PrecioConImpuesto = 0
                    .Afecto = IIf(Trim(rstDuplicadoOC!F5AFECTO & "") = "*", True, False)
                    .PorcentajeDscto = Val(rstDuplicadoOC!F3PORDCT & "") / 100
                    .TotalDscto = 0
                    
                    .calculosPorItem
                    
                    .ObservacionPorItem = "****"
                    
                    .PorcentajeDemasia = .PorcentajeDemasia * 100
                    .PorcentajeDscto = .PorcentajeDscto * 100
                    
                    .ItemAjustado = CBool(rstDuplicadoOC!F3AJUSTE)
                    
                    .actualizarOrdenDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "DELETE FROM IF3ORDEN "
                    SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F4LOCAL = '" & Trim(rstDuplicadoOC!F4LOCAL & "") & "' AND "
                    SqlCad = SqlCad & "F4NUMORD = '" & Trim(rstDuplicadoOC!F4NUMORD & "") & "' AND "
                    SqlCad = SqlCad & "COD_SOLICITUD = '" & Trim(rstDuplicadoOC!COD_SOLICITUD & "") & "' AND "
                    SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstDuplicadoOC!F3CODPRO & "") & "' AND "
                    SqlCad = SqlCad & "F3OBSERVA <> '****'"
                    
                    cnn_dbbancos.Execute SqlCad
                    
                    Actualiza_Log SqlCad, StrConexDbBancos
                    
                    dblOCsCorregidas = dblOCsCorregidas + 1
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Corrección de Duplicados en Detalle de OC's..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Correcciones: " & dblOCsCorregidas & " / " & pgbProceso.Max
            
            rstDuplicadoOC.MoveNext
        Loop
            MsgBox "Proceso Finalizado, se Corrigieron " & dblOCsCorregidas & " OCs.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    
    Exit Sub
errCorregirDuplicadosEnDetalleOC:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlRequerimientoCP()
    On Error GoTo errExportarSqlRequerimientoCP
    
    Dim rstExportarCabSql As New ADODB.Recordset
    Dim rstExportarDetSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT CS_DOCUMENTO, COD_SOLICITUD FROM TB_CABSOLICITUD ORDER BY CS_DOCUMENTO, CS_FECHA, COD_SOLICITUD"
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    
    rstExportarCabSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarCabSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarCabSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Requerimientos CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM PROCESOS.TB_DETSOLICITUD"
        
        cnBdCPlus.Execute SqlCad
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM PROCESOS.TB_CABSOLICITUD"
        
        cnBdCPlus.Execute SqlCad
        
        Do While Not rstExportarCabSql.EOF
            With objAyudaSolicitud
                .inicializarEntidades
                
                .TipoDocumento = Trim(rstExportarCabSql!CS_DOCUMENTO & "")
                .Codigo = Trim(rstExportarCabSql!COD_SOLICITUD & "")
                
                .obtenerConfigSolicitud
            End With
            
            With objSqlAyudaSolicitud
                .inicializarEntidades
                
                .TipoDocumento = objAyudaSolicitud.TipoDocumento
                .Codigo = objAyudaSolicitud.Codigo
                .Fecha = Format(objAyudaSolicitud.Fecha, "dd/mm/yyyy")
                .Estado1 = objAyudaSolicitud.Fecha
                .VBJefe = objAyudaSolicitud.VBJefe
                .VBFecha = Format(objAyudaSolicitud.VBFecha, "dd/mm/yyyy")
                .VBUsuario = objAyudaSolicitud.VBUsuario
                .Estado2 = objAyudaSolicitud.Estado2
                .Prioridad = objAyudaSolicitud.Prioridad
                .observaciones = objAyudaSolicitud.observaciones
                .NombreReferencial = objAyudaSolicitud.NombreReferencial
                .Anulado = objAyudaSolicitud.Anulado
                
                If .Anulado Then
                    .Estado1 = objAyudaSolicitud.Estado1
                End If
                
                .CodigoSolicitante = objAyudaSolicitud.CodigoSolicitante
                .CodigoAprobadoPor1 = objAyudaSolicitud.CodigoAprobadoPor1
                .Empresa = objAyudaSolicitud.Empresa
                .FechaEntrega = Format(objAyudaSolicitud.FechaEntrega, "dd/mm/yyyy")
                .Usuario = objAyudaSolicitud.Usuario
                
                .FechaReg = objAyudaSolicitud.FechaReg
                .FechaMod = objAyudaSolicitud.FechaMod
                
                If .guardarSolicitud(False) Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    'Borrar Detalle de Requerimiento
                    SqlCad = vbNullString
                    SqlCad = "DELETE FROM PROCESOS.TB_DETSOLICITUD WHERE CS_DOCUMENTO = '" & .TipoDocumento & "' AND COD_SOLICITUD = '" & .Codigo & "'"
                    
                    cnBdCPlus.Execute SqlCad
                    
                    Actualiza_Log " < Replicación BD > " & SqlCad, cnn_dbbancos.ConnectionString
                    
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "SELECT * FROM TB_DETSOLICITUD WHERE CS_DOCUMENTO = '" & .TipoDocumento & "' AND COD_SOLICITUD = '" & .Codigo & "'"
                    
                    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
                    
                    rstExportarDetSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly '3, 1
                    
                    If Not rstExportarDetSql.EOF Then
                        
                        Do While Not rstExportarDetSql.EOF
                            .inicializarEntidadesDetalle
                            
                            .CodProducto = Trim(rstExportarDetSql!COD_PRODUCTO & "")
                            .Descripcion = Trim(rstExportarDetSql!ds_descripcion & "")
                            .Cantidad = Val(rstExportarDetSql!ds_cantidad & "")
                            .Cantidad2 = Val(rstExportarDetSql!candis & "")
                            .ITEM = Val(rstExportarDetSql!ITEM & "")
                            .CodUniMed = Trim(rstExportarDetSql!ds_unidmed & "")
                            
                            .guardarSolicitudDetalleOneByOne
                            
                            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                            
                            rstExportarDetSql.MoveNext
                        Loop
                    End If
                Else
                    Actualiza_Log " < Replicación BD > Importación de Requerimiento No. " & .Codigo & " fallido.", StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Requerimientos CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarCabSql.MoveNext
        Loop
    End If
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarCabSql = Nothing
    Set rstExportarDetSql = Nothing
    
    Exit Sub
errExportarSqlRequerimientoCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    'Resume
    Err.Clear
End Sub

Private Sub exportarSqlOrdenCompraCP()
    On Error GoTo errExportarSqlOrdenCompraCP
    
    Dim rstExportarCabSql As New ADODB.Recordset
    Dim rstExportarDetSql As New ADODB.Recordset
    
    Dim intAnno As Integer
    Dim intMes As Integer
    
    intAnno = Val(InputBox("Ingrese la Año que desea Exportar:", "Exportacion", Year(Date)))
    
    intMes = Val(InputBox("Ingrese la Mes que desea Exportar:", "Exportacion", Month(Date)))
    
    If intAnno < 2014 Or intAnno > Year(Date) Then
        intAnno = 0
    End If
    
    If intMes < 1 Or intMes > 12 Then
        intMes = 0
    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F4LOCAL, "
    SqlCad = SqlCad & "F4NUMORD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF4ORDEN "
    SqlCad = SqlCad & "WHERE "
    'SqlCad = SqlCad & "F4LOCAL = 'OC' "
    SqlCad = SqlCad & "F4LOCAL <> '' "
        
        If intAnno <> 0 Then
            SqlCad = SqlCad & "AND YEAR(F4FECEMI) = " & intAnno & " "
        End If
        
        If intAnno <> 0 And intMes <> 0 Then
            SqlCad = SqlCad & "AND MONTH(F4FECEMI) = " & intMes & " "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4LOCAL, "
    SqlCad = SqlCad & "F4NUMORD"
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    
    rstExportarCabSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarCabSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarCabSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Ordenes de Compra CP..."
        
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM PROCESOS.IF3ORDEN"
'
'        cnBdCPlus.Execute SqlCad
'
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM PROCESOS.IF4ORDEN"
'
'        cnBdCPlus.Execute SqlCad
        
        Do While Not rstExportarCabSql.EOF
            With objAyudaOrden
                .inicializarEntidades
                
                .TipoOrden = Trim(rstExportarCabSql!F4LOCAL & "")
                .NumeroOrden = Trim(rstExportarCabSql!F4NUMORD & "")
                
                .obtenerConfigOrden
            End With
            
            With objSqlAyudaOrden
                .inicializarEntidades
                
                .TipoOrden = objAyudaOrden.TipoOrden
                .NumeroOrden = objAyudaOrden.NumeroOrden
                
                .FechaEmision = Format(objAyudaOrden.FechaEmision, "Short Date")
                .SinProveedorEspecifico = objAyudaOrden.SinProveedorEspecifico
                .NomProveedor = Replace(objAyudaOrden.NomProveedor, "'", "´", 1)
                .RucProveedor = objAyudaOrden.RucProveedor
                .CodProveedor = objAyudaOrden.CodProveedor
                .ContactoProveedor = Replace(objAyudaOrden.ContactoProveedor, "'", "´", 1)
                
                .CodTipoComprobante = objAyudaOrden.CodTipoComprobante
                .OrdenRegularizada = objAyudaOrden.OrdenRegularizada
                
                .FechaEntrega = Format(objAyudaOrden.FechaEntrega, "Short Date")
                .CodigoSolicitante = objAyudaOrden.CodigoSolicitante
                .CodFormaPago = objAyudaOrden.CodFormaPago
                .CentroCosto = objAyudaOrden.CentroCosto
                .LugarEntrega = objAyudaOrden.LugarEntrega
                .PagoParcial = objAyudaOrden.PagoParcial
        
                .CodMoneda = objAyudaOrden.CodMoneda
                .TipoCambio = Format(objAyudaOrden.TipoCambio, "#.000")
                .NumeroCotizacion = objAyudaOrden.NumeroCotizacion
                
                .Colocada = objAyudaOrden.Colocada
                    .ColocadaUsuario = objAyudaOrden.ColocadaUsuario
                    .ColocadaFecha = IIf(.Colocada, Format(objAyudaOrden.ColocadaFecha, "Short Date"), vbNullString)
                
                .Atendida = objAyudaOrden.Atendida
                    .AtendidaUsuario = objAyudaOrden.AtendidaUsuario
                    .AtendidaFecha = IIf(.Atendida, Format(objAyudaOrden.AtendidaFecha, "Short Date"), vbNullString)
                
                .Empresa = objAyudaOrden.Empresa
                .Observacion = objAyudaOrden.Observacion
        
                .SUBTOTAL = Val(Format(objAyudaOrden.SUBTOTAL, "0.00"))
                .TotalInafecto = Val(Format(objAyudaOrden.TotalInafecto, "0.00"))
                .TotalImpuesto = Val(Format(objAyudaOrden.TotalImpuesto, "0.00"))
                .TotalFacturado = Val(Format(objAyudaOrden.TotalFacturado, "0.00"))
                
                .FechaReg = Format(objAyudaOrden.FechaReg, "Short Date")
                .UsuarioReg = objAyudaOrden.UsuarioReg
                .FechaMod = Format(objAyudaOrden.FechaMod, "Short Date")
                .UsuarioMod = objAyudaOrden.UsuarioMod
                
                .Estado = objAyudaOrden.Estado
                
                If .guardarOrden Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    .SQLSelectAlter = "DELETE FROM PROCESOS.IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'"
        
                    cnBdCPlus.Execute .SQLSelectAlter
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
        
                    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
                    
                    rstExportarDetSql.Open "SELECT * FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly 'AND VAL(F3CANPRO & '') > 0
                    
                    If Not rstExportarDetSql.EOF Then
                        rstExportarDetSql.MoveFirst
                        
                        fraProceso2.Visible = True
                        pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstExportarDetSql)
                        pgbProceso2.Value = 0
                        fraProceso2.Caption = "Exportando a SQL las Ordenes de Compra CP..."
                                        
                        Do While Not rstExportarDetSql.EOF
                            .inicializarEntidadesDetalle
        
                            .ITEM = Val(rstExportarDetSql!ITEM & "")
                            .Requerimiento = Trim(rstExportarDetSql!COD_SOLICITUD & "")
                            .CodigoProducto = Trim(rstExportarDetSql!F3CODPRO & "")
                            .CodigoFabricante = Trim(rstExportarDetSql!F3CODFAB & "")
                            .NombreProducto = Trim(rstExportarDetSql!F5NOMPRO & "")
                            .NombreProductoInterno = Trim(rstExportarDetSql!F5NOMPRO_ING & "")
                            .CodigoUM = Trim(rstExportarDetSql!UNIDAD & "")
                            .Cantidad = Val(rstExportarDetSql!F3CANPRO & "")
                            .CantidadMaxima = Val(rstExportarDetSql!F3CANPRO2 & "")
                            .CantidadFaltante = Val(rstExportarDetSql!F3CANFAL & "")
        
                            .PorcentajeDemasia = Val(rstExportarDetSql!F3PORCDEMASIA & "")
        
                            .PrecioSinImpuesto = Val(rstExportarDetSql!F3PRECOS & "")
                            .PrecioConImpuesto = Val(rstExportarDetSql!F3PREUNI & "")
                            .PrecioNetoSinImpuesto = Val(rstExportarDetSql!F3PRENETO & "")
        
                            .PorcentajeDscto = Val(rstExportarDetSql!F3PORDCT & "")
                            .TotalDscto = Val(rstExportarDetSql!F3TOTDCT & "")
        
                            .Afecto = IIf(Trim(rstExportarDetSql!F5AFECTO) = "*", True, False)
        
                            .BasePorItem = Val(rstExportarDetSql!F5VALVTA & "")
                            .ImpuestoPorItem = Val(rstExportarDetSql!F3IGV & "")
                            .TotalPorItem = Val(rstExportarDetSql!F3TOTAL & "")
                            
                            .CodigoColor = Trim(rstExportarDetSql!CODCOLOR & "")
                            .ObservacionPorItem = Trim(rstExportarDetSql!F3OBSERVA & "")
        
                            .ItemAjustado = CBool(rstExportarDetSql!F3AJUSTE)
        
                            .CodigoGasto = Trim(rstExportarDetSql!F3GASTO & "")
                            .CuentaContable = Trim(rstExportarDetSql!F3CUENTA & "")
        
                            .guardarOrdenDetalleOneByOne
        
                            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                                            
                            DoEvents
                            
                            pgbProceso2.Value = pgbProceso2.Value + 1
                            fraProceso2.Caption = "Exportando a SQL Detalle de Orden... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3) & " - Exportados: " & pgbProceso2.Value & " / " & pgbProceso2.Max
                            
                            rstExportarDetSql.MoveNext
                        Loop
                    End If
                Else
                    Actualiza_Log " < Replicación BD > Importación de Orden de Compra No. " & .NumeroOrden & " fallido.", StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Ordenes de Compra CP [" & intAnno & "-" & intMes & "]... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarCabSql.MoveNext
        Loop
    End If
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarCabSql = Nothing
    Set rstExportarDetSql = Nothing
    
    Exit Sub
errExportarSqlOrdenCompraCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlValeCP()
    On Error GoTo errExportarSqlValeCP
    
    Dim rstExportarCabSql As New ADODB.Recordset
    Dim rstExportarDetSql As New ADODB.Recordset
    
    Dim intAnno As Integer
    Dim intMes As Integer
    
    intAnno = Val(InputBox("Ingrese la Año que desea Exportar:", "Exportacion", Year(Date)))
    
    intMes = Val(InputBox("Ingrese la Mes que desea Exportar:", "Exportacion", Month(Date)))
    
    If intAnno < 2014 Or intAnno > Year(Date) Then
        intAnno = 0
    End If
    
    If intMes < 1 Or intMes > 12 Then
        intMes = 0
    End If
    
    'Vales de Ingreso
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F4TIPOVALE, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F4NUMVAL "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF4VALES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F4NUMVAL <> '' "
            
        If intAnno <> 0 Then
            SqlCad = SqlCad & "AND YEAR(F4FECVAL) = " & intAnno & " "
        End If
        
        If intAnno <> 0 And intMes <> 0 Then
            SqlCad = SqlCad & "AND MONTH(F4FECVAL) = " & intMes & " "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4FECVAL, "
    SqlCad = SqlCad & "F4TIPOVALE, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F4NUMVAL"
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    
    rstExportarCabSql.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rstExportarCabSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarCabSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Vales de CP..."
        
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM PROCESOS.IF3VALES "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "F4NUMVAL <> '' "
'
'            If intAnno <> 0 Then
'                SqlCad = SqlCad & "AND YEAR(F4FECVAL) = " & intAnno & " "
'            End If
'
'            If intAnno <> 0 And intMes <> 0 Then
'                SqlCad = SqlCad & "AND MONTH(F4FECVAL) = " & intMes & " "
'            End If
'
'        cnBdCPlus.Execute SqlCad
'
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM PROCESOS.IF4VALES "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "F4NUMVAL <> '' "
'
'            If intAnno <> 0 Then
'                SqlCad = SqlCad & "AND YEAR(F4FECVAL) = " & intAnno & " "
'            End If
'
'            If intAnno <> 0 And intMes <> 0 Then
'                SqlCad = SqlCad & "AND MONTH(F4FECVAL) = " & intMes & " "
'            End If
'
'        cnBdCPlus.Execute SqlCad
        
        Do While Not rstExportarCabSql.EOF
            With objAyudaVale
                .inicializarEntidades
                
                .TipoVale = Trim(rstExportarCabSql!F4TIPOVALE & "")
                .NumeroVale = Trim(rstExportarCabSql!F4NUMVAL & "")
                .CodigoAlmacen = Trim(rstExportarCabSql!f2codalm & "")
                
                .obtenerConfigVale
            End With
            
            With objSqlAyudaVale
                .inicializarEntidades
                
                .TipoVale = objAyudaVale.TipoVale
                .NumeroVale = objAyudaVale.NumeroVale
                .NumeroValeExterno = objAyudaVale.NumeroValeExterno
                
                .CodigoAlmacen = objAyudaVale.CodigoAlmacen
                .CodigoOrigen = objAyudaVale.CodigoOrigen
                
                .TipoPersona = objAyudaVale.TipoPersona
                    .CodigoProveedor = objAyudaVale.CodigoProveedor
                    
                .CentroCosto = objAyudaVale.CentroCosto
                .SerieGuia = objAyudaVale.SerieGuia
                .NumeroGuia = objAyudaVale.NumeroGuia
                
                .CodTipoComprobante = objAyudaVale.CodTipoComprobante
                .SerieDocumento = objAyudaVale.SerieDocumento
                .NumeroDocumento = objAyudaVale.NumeroDocumento
                
                If .CodTipoComprobante <> vbNullString And .NumeroDocumento <> vbNullString Then
                    .FechaUltima = Format(objAyudaVale.FechaUltima, "Short Date")
                Else
                    .FechaUltima = vbNullString
                End If
                
                .OrdenTrabajo = objAyudaVale.OrdenTrabajo
                
                .Fecha = Format(objAyudaVale.Fecha, "Short Date")
                .CodigoMoneda = objAyudaVale.CodigoMoneda
                .TipoCambio = objAyudaVale.TipoCambio
                
                .NumeroOrdenCompra = objAyudaVale.NumeroOrdenCompra
                .observaciones = objAyudaVale.observaciones
                
                .RegistroCompra = objAyudaVale.RegistroCompra
                
                .ExportarVale = objAyudaVale.ExportarVale
                
                .FecReg = Format(objAyudaVale.FecReg, "Short Date")
                .UsuReg = objAyudaVale.UsuReg
                .FecMod = Format(objAyudaVale.FecMod, "Short Date")
                .UsuMod = objAyudaVale.UsuMod
                
                If .guardarVale Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    .SQLSelectAlter = "DELETE FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                    
                    cnBdCPlus.Execute .SQLSelectAlter
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
                    
                    rstExportarDetSql.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstExportarDetSql.EOF Then
                        rstExportarDetSql.MoveFirst
                        
                        fraProceso2.Visible = True
                        pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstExportarDetSql)
                        pgbProceso2.Value = 0
                        fraProceso2.Caption = "Exportando a SQL los Vales de CP..."
                        
                        Do While Not rstExportarDetSql.EOF
                            .inicializarEntidadesDetalle
                            
                            .CodigoProducto = Trim(rstExportarDetSql!f5codpro & "")
                            .CodigoProductoOriginal = Trim(rstExportarDetSql!F5CODPROORIGINAL & "")
                            .Cantidad = Val(rstExportarDetSql!F3CANPRO & "")
                            .CantidadMaxima = Val(rstExportarDetSql!F3SALPEP & "")
                            
                            .ValorVenta = Val(rstExportarDetSql!F3VALVTA & "")
                            .IGV = Val(rstExportarDetSql!F3IGV & "")
                            .TOTAL = Val(rstExportarDetSql!F3TOTITE & "")
                            .ValorVentaDol = Val(rstExportarDetSql!F3VALDOL & "")
                            .IgvDol = Val(rstExportarDetSql!F3IGVDOL & "")
                            .TotalDol = Val(rstExportarDetSql!F3TOTDOL & "")
                            
                            .Grupo = Trim(rstExportarDetSql!F3GRUPO & "")
                            .ITEM = Val(rstExportarDetSql!F3ITEM & "")
                            .NumeroOrdenCompra = Trim(rstExportarDetSql!F4NUMORD & "")
                            .Requerimiento = Trim(rstExportarDetSql!COD_SOLICITUD & "")
                            .PorcentajeDscto = Val(rstExportarDetSql!F3PORCENTAJEDSCTO & "")
                            .MontoDscto = Val(rstExportarDetSql!F3MONTODSCTO & "")
                            
                            .ObservacionesPorItem = Trim(rstExportarDetSql!observaciones & "")
                            
                            .guardarValeDetalleOneByOne
                            
                            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                            
                            DoEvents
                            
                            pgbProceso2.Value = pgbProceso2.Value + 1
                            fraProceso2.Caption = "Exportando a SQL Detalle de Vale CP... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3) & " - Exportados: " & pgbProceso2.Value & " / " & pgbProceso2.Max
                            
                            rstExportarDetSql.MoveNext
                        Loop
                    End If
                Else
                    Actualiza_Log " < Replicación BD > Importación de Vale No. " & .CodigoAlmacen & " / " & .NumeroVale & " fallido.", StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Vales de Ingreso CP [" & intAnno & "-" & intMes & "]... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarCabSql.MoveNext
        Loop
    End If
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarCabSql = Nothing
    Set rstExportarDetSql = Nothing
    
    Exit Sub
errExportarSqlValeCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlTomaInventarioCP()
    On Error GoTo errExportarSqlTomaInventarioCP
    
    Dim rstExportarCabSql As New ADODB.Recordset
    Dim rstExportarDetSql As New ADODB.Recordset
    
    'Vales de Ingreso
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F4ANNO, "
    SqlCad = SqlCad & "F4MES, "
    SqlCad = SqlCad & "F2CODALM "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "H4TOMAINV "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4ANNO, "
    SqlCad = SqlCad & "F4MES, "
    SqlCad = SqlCad & "F2CODALM"
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    
    rstExportarCabSql.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rstExportarCabSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarCabSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL Toma de Inventario de CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM PROCESOS.H3TOMAINV "

        cnBdCPlus.Execute SqlCad

        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM PROCESOS.H4TOMAINV "

        cnBdCPlus.Execute SqlCad
        
        Do While Not rstExportarCabSql.EOF
            With objAyudaTomaInventario
                .inicializarEntidades
                
                .AnnoTI = Trim(rstExportarCabSql!F4ANNO & "")
                .MesTI = Trim(rstExportarCabSql!F4MES & "")
                .CodigoAlmacen = Trim(rstExportarCabSql!f2codalm & "")
                
                .obtenerConfigTomaInventario
            End With
            
            With objSqlAyudaTomaInventario
                .inicializarEntidades
                
                .AnnoTI = objAyudaTomaInventario.AnnoTI
                .MesTI = objAyudaTomaInventario.MesTI
                .CodigoAlmacen = objAyudaTomaInventario.CodigoAlmacen
                .Fecha = objAyudaTomaInventario.Fecha
                .Observacion = objAyudaTomaInventario.Observacion
                .ValeIngreso = objAyudaTomaInventario.ValeIngreso
                    .ValeIngresoExterno = objAyudaTomaInventario.ValeIngresoExterno
                .ValeSalida = objAyudaTomaInventario.ValeSalida
                    .ValeSalidaExterno = objAyudaTomaInventario.ValeSalidaExterno
                
                .CierreInventario = objAyudaTomaInventario.CierreInventario
                
                .FecReg = objAyudaTomaInventario.FecReg
                .UsuReg = objAyudaTomaInventario.UsuReg
                .FecMod = objAyudaTomaInventario.FecMod
                .UsuMod = objAyudaTomaInventario.UsuMod
                
                If .guardarTomaInventario Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    .SQLSelectAlter = "DELETE FROM PROCESOS.H3TOMAINV WHERE F4ANNO = '" & .AnnoTI & "' AND F4MES = '" & .MesTI & "' AND F2CODALM = '" & .CodigoAlmacen & "'"
                    
                    cnBdCPlus.Execute .SQLSelectAlter
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
                    
                    rstExportarDetSql.Open "SELECT * FROM H3TOMAINV WHERE F4ANNO = '" & .AnnoTI & "' AND F4MES = '" & .MesTI & "' AND F2CODALM = '" & .CodigoAlmacen & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstExportarDetSql.EOF Then
                        rstExportarDetSql.MoveFirst
                        
                        fraProceso2.Visible = True
                        pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstExportarDetSql)
                        pgbProceso2.Value = 0
                        fraProceso2.Caption = "Exportando a SQL Detalle de Toma de Inventario de CP..."
                        
                        Do While Not rstExportarDetSql.EOF
                            .inicializarEntidadesDetalle
                            
                            .CodigoProducto = Trim(rstExportarDetSql!f5codpro & "")
                            .StockSistema = Val(rstExportarDetSql!F3STOCKSISTEMA & "")
                            .StockFisico = Val(rstExportarDetSql!F3STOCKFISICO & "")
                            .Diferencia = Val(rstExportarDetSql!F3DIFERENCIA & "")
                            
                            .guardarTomaInvDetalleOneByOne
                            
                            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                            
                            DoEvents
                            
                            pgbProceso2.Value = pgbProceso2.Value + 1
                            fraProceso2.Caption = "Exportando a SQL Detalle de Toma de Inventario de CP... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3) & " - Exportados: " & pgbProceso2.Value & " / " & pgbProceso2.Max
                            
                            rstExportarDetSql.MoveNext
                        Loop
                    End If
                    
                    If .CierreInventario Then
                        .SQLSelectAlter = "UPDATE PROCESOS.H4TOMAINV SET F4CIERRE = 1 WHERE F4ANNO = '" & .AnnoTI & "' AND F4MES = '" & .MesTI & "' AND F2CODALM = '" & .CodigoAlmacen & "'"
                        
                        cnBdCPlus.Execute .SQLSelectAlter
                        
                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    End If
                Else
                    Actualiza_Log " < Replicación BD > Importación de Toma de Inventario No. " & .AnnoTI & " / " & .MesTI & " / " & .CodigoAlmacen & " fallido.", StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL Toma de Inventario de CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarCabSql.MoveNext
        Loop
    End If
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarCabSql = Nothing
    Set rstExportarDetSql = Nothing
    
    Exit Sub
errExportarSqlTomaInventarioCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlCierreValeCP()
    On Error GoTo errExportarSqlCierreValeCP
    
    Dim rstExportarCabSql As New ADODB.Recordset
    Dim rstExportarDetSql As New ADODB.Recordset
    
    'Vales de Ingreso
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "F4TIPOVALE, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F4NUMVAL, "
    SqlCad = SqlCad & "VB1, "
    SqlCad = SqlCad & "VB1USER, "
    SqlCad = SqlCad & "VB1FECHA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF4VALES "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F1CODORI NOT IN ('XCS') AND "
    SqlCad = SqlCad & "CVDATE(F4FECVAL) < CVDATE('01/08/2015') "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4FECVAL, "
    SqlCad = SqlCad & "F4TIPOVALE, "
    SqlCad = SqlCad & "F2CODALM, "
    SqlCad = SqlCad & "F4NUMVAL"
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    
    rstExportarCabSql.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not rstExportarCabSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarCabSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Vales de CP..."
        
        Do While Not rstExportarCabSql.EOF
            With objSqlAyudaVale
                .inicializarEntidades
                
                .TipoVale = Trim(rstExportarCabSql!F4TIPOVALE & "")
                .CodigoAlmacen = Trim(rstExportarCabSql!f2codalm & "")
                .NumeroVale = Trim(rstExportarCabSql!F4NUMVAL & "")
                
                .VB1 = CBool(rstExportarCabSql!VB1)
                .VB1Usuario = Trim(rstExportarCabSql!VB1USER & "")
                .VB1Fecha = Format(Trim(rstExportarCabSql!VB1Fecha & ""), "Short Date")
                
                If .cerrarVale Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                Else
                    Actualiza_Log " < Replicación BD > Cierre de Vale No. " & .CodigoAlmacen & " / " & .NumeroVale & " fallido.", StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Vales de Ingreso CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarCabSql.MoveNext
        Loop
    End If
    
    If rstExportarCabSql.State = 1 Then rstExportarCabSql.Close
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarCabSql = Nothing
    Set rstExportarDetSql = Nothing
    
    Exit Sub
errExportarSqlCierreValeCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlOrigenesCP()
    On Error GoTo errExportarSqlOrigenesCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F1CODORI FROM SF1ORIGENES"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Origenes CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.SF1ORIGENES"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaOrigen
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F1CODORI & "")
                
                .obtenerConfigOrigen
            End With
            
            With objSqlAyudaOrigen
                .inicializarEntidades
                
                .Codigo = objAyudaOrigen.Codigo
                .CodigoExterno = objAyudaOrigen.CodigoExterno
                .Descripcion = objAyudaOrigen.Descripcion
                .TipoMovimiento = objAyudaOrigen.TipoMovimiento
                .RegistrarCosto = objAyudaOrigen.RegistrarCosto
                .TieneAlmacenDestino = objAyudaOrigen.TieneAlmacenDestino
                .CodigoAyudaProducto = objAyudaOrigen.CodigoAyudaProducto
                
                .Estado = objAyudaOrigen.Estado
                
                .FechaReg = IIf(objAyudaOrigen.FechaReg = vbNullString, Format(Date, "Short Date"), objAyudaOrigen.FechaReg)
                .UsuarioReg = IIf(objAyudaOrigen.UsuarioReg = vbNullString, wusuario, objAyudaOrigen.UsuarioReg)
                .FechaMod = IIf(objAyudaOrigen.FechaMod = vbNullString, Format(Date, "Short Date"), objAyudaOrigen.FechaMod)
                .UsuarioMod = IIf(objAyudaOrigen.UsuarioMod = vbNullString, wusuario, objAyudaOrigen.UsuarioMod)
                
                If .guardarOrigen Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Origenes CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlOrigenesCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlAlmacenesCP()
    On Error GoTo errExportarSqlAlmacenesCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODALM FROM EF2ALMACENES"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Almacenes CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2ALMACENES"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaAlmacen
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!f2codalm & "")
                
                .obtenerConfigAlmacen
            End With
            
            With objSqlAyudaAlmacen
                .inicializarEntidades
                
                .Codigo = objAyudaAlmacen.Codigo
                .CodigoExterno = objAyudaAlmacen.CodigoExterno
                .Descripcion = objAyudaAlmacen.Descripcion
                .Direccion = objAyudaAlmacen.Direccion
                .RucAlmacen = objAyudaAlmacen.RucAlmacen
                .CentroCosto = objAyudaAlmacen.CentroCosto
                
                If .guardarAlmacen Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Almacenes CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlAlmacenesCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlTiposExistenciasCP()
    On Error GoTo errExportarSqlTiposExistenciasCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT CODIGO FROM EF2TIPOEXISTENCIA"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Tipos de Existencia CP..."
        
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2TIPOEXISTENCIA"
'
'        cnBdCPlus.Execute SqlCad
'
'        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaTipoExistencia
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!Codigo & "")
                
                .obtenerConfigTipoExistencia
            End With
            
            With objSqlAyudaTipoExistencia
                .inicializarEntidades
                
                .Codigo = objAyudaTipoExistencia.Codigo
                .CodigoExterno = objAyudaTipoExistencia.CodigoExterno
                .Descripcion = objAyudaTipoExistencia.Descripcion
                .Abreviatura = objAyudaTipoExistencia.Abreviatura
                .Estado = objAyudaTipoExistencia.Estado
                .FechaReg = IIf(objAyudaTipoExistencia.FechaReg = vbNullString, Format(Date, "Short Date"), objAyudaTipoExistencia.FechaReg)
                .UsuarioReg = IIf(objAyudaTipoExistencia.UsuarioReg = vbNullString, wusuario, objAyudaTipoExistencia.UsuarioReg)
                .FechaMod = IIf(objAyudaTipoExistencia.FechaMod = vbNullString, Format(Date, "Short Date"), objAyudaTipoExistencia.FechaMod)
                .UsuarioMod = IIf(objAyudaTipoExistencia.UsuarioMod = vbNullString, wusuario, objAyudaTipoExistencia.UsuarioMod)
                
                .CtaContable = objAyudaTipoExistencia.CtaContable
                .Anexo = objAyudaTipoExistencia.Anexo
                .CtaContableImportacion = objAyudaTipoExistencia.CtaContableImportacion
                .AnexoImportacion = objAyudaTipoExistencia.AnexoImportacion
                
                .CtaContableVta = objAyudaTipoExistencia.CtaContableVta
                .AnexoVta = objAyudaTipoExistencia.AnexoVta
                .CtaContableImportacionVta = objAyudaTipoExistencia.CtaContableImportacionVta
                .AnexoImportacionVta = objAyudaTipoExistencia.AnexoImportacionVta
                
                If .guardarTipoExistencia Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Tipos de Existencia CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlTiposExistenciasCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlCentrosCP()
    On Error GoTo errExportarSqlCentrosCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F3COSTO FROM CENTROS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Centros de Costo CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.CENTROS"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaCentroCosto
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F3COSTO & "")
                
                .obtenerConfigCentroCosto
            End With
            
            With objSqlAyudaCentroCosto
                .inicializarEntidades
                
                .Codigo = objAyudaCentroCosto.Codigo
                .CodigoNivel = objAyudaCentroCosto.CodigoNivel
                .CodigoExterno = objAyudaCentroCosto.CodigoExterno
                .CodigoConcar = objAyudaCentroCosto.CodigoConcar
                .Descripcion = objAyudaCentroCosto.Descripcion
                .Abreviatura = objAyudaCentroCosto.Abreviatura
                .CodigoCliente = objAyudaCentroCosto.CodigoCliente
                .Utilidad = objAyudaCentroCosto.Utilidad
                
                .Estado = objAyudaCentroCosto.Estado
                
                .FechaReg = IIf(objAyudaCentroCosto.FechaReg = vbNullString, Format(Date, "Short Date"), objAyudaCentroCosto.FechaReg)
                .UsuarioReg = IIf(objAyudaCentroCosto.UsuarioReg = vbNullString, wusuario, objAyudaCentroCosto.UsuarioReg)
                .FechaMod = IIf(objAyudaCentroCosto.FechaMod = vbNullString, Format(Date, "Short Date"), objAyudaCentroCosto.FechaMod)
                .UsuarioMod = IIf(objAyudaCentroCosto.UsuarioMod = vbNullString, wusuario, objAyudaCentroCosto.UsuarioMod)
                
                If .guardarCentroCosto Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Centros de Costo CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlCentrosCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlMarcasCP()
    On Error GoTo errExportarSqlMarcasCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODMAR FROM EF2MARCAS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Marcas CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2MARCAS"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaMarca
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!f2codmar & "")
                
                .obtenerConfigMarca
            End With
            
            With objSqlAyudaMarca
                .inicializarEntidades
                
                .Codigo = objAyudaMarca.Codigo
                .Descripcion = objAyudaMarca.Descripcion
                .Origen = objAyudaMarca.Origen
                .MarcaObc = objAyudaMarca.MarcaObc
                
                If .guardarMarca Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Marcas CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlMarcasCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlMedidasCP()
    On Error GoTo errExportarSqlMedidasCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F7CODMED FROM EF7MEDIDAS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Unidades de Medida CP..."
        
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF7MEDIDAS"
'
'        cnBdCPlus.Execute SqlCad
'
'        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaUM
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!f7codmed & "")
                
                .obtenerConfigUM
            End With
            
            With objSqlAyudaUM
                .inicializarEntidades
                
                .Codigo = objAyudaUM.Codigo
                .Descripcion = objAyudaUM.Descripcion
                .Abreviatura = objAyudaUM.Abreviatura
                .UMParaExportacion = objAyudaUM.UMParaExportacion
                .Estado = objAyudaUM.Estado
                
                .FechaReg = IIf(objAyudaUM.FechaReg = vbNullString, Format(Date, "Short Date"), Format(objAyudaUM.FechaReg, "Short Date"))
                .UsuarioReg = IIf(objAyudaUM.UsuarioReg = vbNullString, wusuario, objAyudaUM.UsuarioReg)
                .FechaMod = IIf(objAyudaUM.FechaMod = vbNullString, Format(Date, "Short Date"), Format(objAyudaUM.FechaMod, "Short Date"))
                .UsuarioMod = IIf(objAyudaUM.UsuarioMod = vbNullString, wusuario, objAyudaUM.UsuarioMod)
                
                If .guardarUM Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Unidades de Medida CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlMedidasCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlBienColorCP()
    On Error GoTo errExportarSqlBienColorCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT CODIGO FROM EF2BIENCOLOR"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Colores de Bien CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2BIENCOLOR"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaBienColor
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!Codigo & "")
                
                .obtenerConfigBienColor
            End With
            
            With objSqlAyudaBienColor
                .inicializarEntidades
                
                .Codigo = objAyudaBienColor.Codigo
                .CodigoExterno = objAyudaBienColor.CodigoExterno
                .Descripcion = objAyudaBienColor.Descripcion
                .Estado = objAyudaBienColor.Estado
                
                .FechaReg = IIf(objAyudaBienColor.FechaReg = vbNullString, Format(Date, "Short Date"), Format(objAyudaBienColor.FechaReg, "Short Date"))
                .UsuarioReg = IIf(objAyudaBienColor.UsuarioReg = vbNullString, wusuario, objAyudaBienColor.UsuarioReg)
                .FechaMod = IIf(objAyudaBienColor.FechaMod = vbNullString, Format(Date, "Short Date"), Format(objAyudaBienColor.FechaMod, "Short Date"))
                .UsuarioMod = IIf(objAyudaBienColor.UsuarioMod = vbNullString, wusuario, objAyudaBienColor.UsuarioMod)
                
                If .guardarBienColor Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Colores de Bien CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlBienColorCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlNivel1CP()
    On Error GoTo errExportarSqlNivel1CP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F7CODCON FROM SF7NIVEL01"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL Nivel 1 CP..."
        
        'SqlCad = vbNullString
        'SqlCad = SqlCad & "DELETE FROM MAESTROS.SF7NIVEL01"
        
        'cnBdCPlus.Execute SqlCad
        
        'Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaFamilia
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F7CODCON & "")
                
                .obtenerConfigFamilia
            End With
            
            With objSqlAyudaFamilia
                .inicializarEntidades
                
                .Codigo = objAyudaFamilia.Codigo
                .CodigoExterno = objAyudaFamilia.CodigoExterno
                .Descripcion = objAyudaFamilia.Descripcion
                .DescripcionCorta = objAyudaFamilia.DescripcionCorta
                .Estado = objAyudaFamilia.Estado
                
                .FechaReg = IIf(objAyudaFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(objAyudaFamilia.FechaReg, "Short Date"))
                .UsuarioReg = IIf(objAyudaFamilia.UsuarioReg = vbNullString, wusuario, objAyudaFamilia.UsuarioReg)
                .FechaMod = IIf(objAyudaFamilia.FechaMod = vbNullString, Format(Date, "Short Date"), Format(objAyudaFamilia.FechaMod, "Short Date"))
                .UsuarioMod = IIf(objAyudaFamilia.UsuarioMod = vbNullString, wusuario, objAyudaFamilia.UsuarioMod)
                
                .CodigoExterno2 = objAyudaFamilia.CodigoExterno2
                .Anexo = objAyudaFamilia.Anexo
                .AnexoImportacion = objAyudaFamilia.AnexoImportacion
                
                If .guardarFamilia Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL Nivel 1 CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlNivel1CP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    Resume
    Err.Clear
End Sub

Private Sub exportarSqlNivel2CP()
    On Error GoTo errExportarSqlNivel2CP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F7CODCON FROM SF7NIVEL02"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL Nivel 2 CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.SF7NIVEL02"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaSubFamilia
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F7CODCON & "")
                
                .obtenerConfigSubFamilia
            End With
            
            With objSqlAyudaSubFamilia
                .inicializarEntidades
                
                .Codigo = objAyudaSubFamilia.Codigo
                .CodigoExterno = objAyudaSubFamilia.CodigoExterno
                .CodigoExterno2 = objAyudaSubFamilia.CodigoExterno2
                .CodigoFamilia = objAyudaSubFamilia.CodigoFamilia
                .Descripcion = objAyudaSubFamilia.Descripcion
                .Estado = objAyudaSubFamilia.Estado
                
                .FechaReg = IIf(objAyudaSubFamilia.FechaReg = vbNullString, Format(Date, "Short Date"), Format(objAyudaSubFamilia.FechaReg, "Short Date"))
                .UsuarioReg = IIf(objAyudaSubFamilia.UsuarioReg = vbNullString, wusuario, objAyudaSubFamilia.UsuarioReg)
                .FechaMod = IIf(objAyudaSubFamilia.FechaMod = vbNullString, Format(Date, "Short Date"), Format(objAyudaSubFamilia.FechaMod, "Short Date"))
                .UsuarioMod = IIf(objAyudaSubFamilia.UsuarioMod = vbNullString, wusuario, objAyudaSubFamilia.UsuarioMod)
                
                If .guardarSubFamilia Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL Nivel 2 CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlNivel2CP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlBienCP()
    On Error GoTo errExportarSqlBienCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    Dim rstMV As New ADODB.Recordset
    Dim rstBA As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F5CODPRO FROM IF5PLA"

    If rstExportarSql.State = 1 Then rstExportarSql.Close

    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Bienes de CP..."

        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.IF5PLA"

        cnBdCPlus.Execute SqlCad

        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos

        Do While Not rstExportarSql.EOF
            With objAyudaBien
                .inicializarEntidades

                .Codigo = Trim(rstExportarSql!f5codpro & "")

                .obtenerConfigBien
            End With

            With objSqlAyudaBien
                .inicializarEntidades

                .Codigo = objAyudaBien.Codigo
                .Descripcion = objAyudaBien.Descripcion
                .CodTipoExistencia = objAyudaBien.CodTipoExistencia
                .CodUM = objAyudaBien.CodUM
                .CodMarca = objAyudaBien.CodMarca
                .CodigoSubFamilia = objAyudaBien.CodigoSubFamilia
                .Modelo = objAyudaBien.Modelo

                .Afecto = objAyudaBien.Afecto
                .Descontinuado = objAyudaBien.Descontinuado
                .ParaVenta = objAyudaBien.ParaVenta
                .EsImportado = objAyudaBien.EsImportado
                .EsInsumoParaOP = objAyudaBien.EsInsumoParaOP
                .TieneMovimientoEnAlmacen = objAyudaBien.TieneMovimientoEnAlmacen

                .CodigoFab = objAyudaBien.CodigoFab
                .DescripcionFab = objAyudaBien.DescripcionFab
                .CodAlmacen = objAyudaBien.CodAlmacen
                .StockMin = objAyudaBien.StockMin
                .StockMax = objAyudaBien.StockMax
                .StockReposicion = objAyudaBien.StockReposicion

                .PorcentajeDemasia = objAyudaBien.PorcentajeDemasia

                .CtaContable = objAyudaBien.CtaContable
                .Anexo = objAyudaBien.Anexo
                .CtaContableImportacion = objAyudaBien.CtaContableImportacion
                .AnexoImportacion = objAyudaBien.AnexoImportacion

                .CodGasto = objAyudaBien.CodGasto
                .CodCentroCosto = objAyudaBien.CodCentroCosto

                .CtaContableVenta = objAyudaBien.CtaContableVenta
                .CtaContableInventa = objAyudaBien.CtaContableInventa

                .FechaIngreso = IIf(objAyudaBien.FechaIngreso = vbNullString, Format(Date, "Short Date"), objAyudaBien.FechaIngreso)
                .UsuarioIngreso = IIf(objAyudaBien.UsuarioIngreso = vbNullString, wusuario, objAyudaBien.UsuarioIngreso)
                .FechaModificacion = IIf(objAyudaBien.FechaModificacion = vbNullString, Format(Date, "Short Date"), objAyudaBien.FechaModificacion)
                .UsuarioModificacion = IIf(objAyudaBien.UsuarioModificacion = vbNullString, wusuario, objAyudaBien.UsuarioModificacion)

                If .guardarBien Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos

                    'MEDIDAS ALTERNAS
                    If rstMV.State = 1 Then rstMV.Close

                    rstMV.Open "SELECT * FROM MEDIVENTAS WHERE F5CODPRO = '" & .Codigo & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

                    If Not rstMV.EOF Then
                        cnBdCPlus.Execute "DELETE FROM MAESTROS.MEDIVENTAS WHERE F5CODPRO = '" & .Codigo & "'"

                        Actualiza_Log " < Replicación BD > DELETE FROM MAESTROS.MEDIVENTAS WHERE F5CODPRO = '" & .Codigo & "'", StrConexDbBancos

                        Do While Not rstMV.EOF
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "INSERT INTO MAESTROS.MEDIVENTAS("
                            SqlCad = SqlCad & "F5CODPRO, F7CODMED, F5FACTOR, F5PREVTA"
                            SqlCad = SqlCad & ") "
                            SqlCad = SqlCad & "VALUES("
                            SqlCad = SqlCad & "'" & Trim(rstMV!f5codpro & "") & "', "
                            SqlCad = SqlCad & "'" & Trim(rstMV!f7codmed & "") & "', "
                            SqlCad = SqlCad & Val(rstMV!F5FACTOR & "") & ", "
                            SqlCad = SqlCad & Val(rstMV!F5PREVTA & "")
                            SqlCad = SqlCad & ")"

                            cnBdCPlus.Execute SqlCad

                            Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos

                            rstMV.MoveNext
                        Loop
                    End If
                End If
            End With

            DoEvents

            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Bienes de CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max

            rstExportarSql.MoveNext
        Loop
    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F5CODPRO FROM IF5PLA"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Bienes Alternos de CP..."
        
        Do While Not rstExportarSql.EOF
            'BIENES ALTERNOS
            If rstBA.State = 1 Then rstBA.Close

            rstBA.Open "SELECT * FROM EF2BIENALTERNO WHERE F5CODPRO = '" & Trim(rstExportarSql!f5codpro & "") & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly

            If Not rstBA.EOF Then
                cnBdCPlus.Execute "DELETE FROM MAESTROS.EF2BIENALTERNO WHERE F5CODPRO = '" & Trim(rstExportarSql!f5codpro & "") & "'"

                Actualiza_Log " < Replicación BD > DELETE FROM MAESTROS.EF2BIENALTERNO WHERE F5CODPRO = '" & Trim(rstExportarSql!f5codpro & "") & "'", StrConexDbBancos

                Do While Not rstBA.EOF
                    SqlCad = vbNullString
                    SqlCad = SqlCad & "INSERT INTO MAESTROS.EF2BIENALTERNO("
                    SqlCad = SqlCad & "F5CODPRO, F5CODPROALTERNO, ESTADO, "
                    SqlCad = SqlCad & "FECREG, USUREG, FECMOD, USUMOD"
                    SqlCad = SqlCad & ") "
                    SqlCad = SqlCad & "VALUES("
                    SqlCad = SqlCad & "'" & Trim(rstBA!f5codpro & "") & "', "
                    SqlCad = SqlCad & "'" & Trim(rstBA!F5CODPROALTERNO & "") & "', "
                    SqlCad = SqlCad & IIf(CBool(rstBA!Estado), "1", "0") & ", "
                    SqlCad = SqlCad & "'" & IIf(Trim(rstBA!FecReg & "") = vbNullString, Format(Date, "Short Date"), Trim(rstBA!FecReg & "")) & "', "
                    SqlCad = SqlCad & "'" & IIf(Trim(rstBA!UsuReg & "") = vbNullString, wusuario, Trim(rstBA!UsuReg & "")) & "', "
                    SqlCad = SqlCad & "'" & IIf(Trim(rstBA!FecMod & "") = vbNullString, Format(Date, "Short Date"), Trim(rstBA!FecMod & "")) & "', "
                    SqlCad = SqlCad & "'" & IIf(Trim(rstBA!UsuMod & "") = vbNullString, wusuario, Trim(rstBA!UsuMod & "")) & "'"
                    SqlCad = SqlCad & ")"

                    cnBdCPlus.Execute SqlCad

                    Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos

                    rstBA.MoveNext
                Loop
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Bienes Alternos de CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlBienCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlCategoriaCP()
    On Error GoTo errExportarSqlCategoriaCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT * FROM CATEGORIA"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Categorias de CP..."
        
        cnBdCPlus.Execute "DELETE FROM MAESTROS.CATEGORIA"
        
        Actualiza_Log " < Replicación BD > DELETE FROM MAESTROS.CATEGORIA", StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO MAESTROS.CATEGORIA("
            SqlCad = SqlCad & "IntCodCategoria, StrDesCategoria"
            SqlCad = SqlCad & ") "
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & Val(rstExportarSql!INTCODCATEGORIA & "") & ", "
            SqlCad = SqlCad & "'" & Trim(rstExportarSql!STRDESCATEGORIA & "") & "'"
            SqlCad = SqlCad & ")"
            
            cnBdCPlus.Execute SqlCad

            Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Categorias de CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlCategoriaCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlTipoDocumentoIDCP()
    On Error GoTo errExportarSqlTipoDocumentoIDCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT CODIGO FROM EF2TIPODOCID"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Tipos de Documento de Identidad CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2TIPODOCID"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaTipoDocID
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!Codigo & "")
                
                .obtenerConfigTipoDocumento
            End With
            
            With objSqlAyudaTipoDocID
                .inicializarEntidades
                
                .Codigo = objAyudaTipoDocID.Codigo
                .Descripcion = objAyudaTipoDocID.Descripcion
                .Abreviatura = objAyudaTipoDocID.Abreviatura
                .Longitud = objAyudaTipoDocID.Longitud
                .TipoCadena = objAyudaTipoDocID.TipoCadena
                .TieneLargoFijo = objAyudaTipoDocID.TieneLargoFijo
                .Modulo11 = objAyudaTipoDocID.Modulo11
                .Origen = objAyudaTipoDocID.Origen
                .Persona = objAyudaTipoDocID.Persona
                
                .FechaReg = IIf(objAyudaTipoDocID.FechaReg = vbNullString, Format(Date, "Short Date"), Format(objAyudaTipoDocID.FechaReg, "Short Date"))
                .UsuarioReg = IIf(objAyudaTipoDocID.UsuarioReg = vbNullString, wusuario, objAyudaTipoDocID.UsuarioReg)
                .FechaMod = IIf(objAyudaTipoDocID.FechaMod = vbNullString, Format(Date, "Short Date"), Format(objAyudaTipoDocID.FechaMod, "Short Date"))
                .UsuarioMod = IIf(objAyudaTipoDocID.UsuarioMod = vbNullString, wusuario, objAyudaTipoDocID.UsuarioMod)
                
                If .guardarTipoDocumento Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Tipos de Documento de Identidad CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlTipoDocumentoIDCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlDistritosCP()
    On Error GoTo errExportarSqlDistritosCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODZON FROM EF2ZONAS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Distritos CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2ZONAS"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaDistrito
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2CODZON & "")
                
                .obtenerConfigDistrito
            End With
            
            With objSqlAyudaDistrito
                .inicializarEntidades
                
                .Codigo = objAyudaDistrito.Codigo
                .Descripcion = objAyudaDistrito.Descripcion
                .CodigoExterno1 = objAyudaDistrito.CodigoExterno1
                .CodigoExterno2 = objAyudaDistrito.CodigoExterno2
                
                If .guardarDistrito Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Distritos CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlDistritosCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlTipoClienteCP()
    On Error GoTo errExportarSqlTipoClienteCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT * FROM EF2TIPOS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Tipos de Cliente de CP..."
        
        cnBdCPlus.Execute "DELETE FROM MAESTROS.EF2TIPOS"
        
        Actualiza_Log " < Replicación BD > DELETE FROM MAESTROS.EF2TIPOS", StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO MAESTROS.EF2TIPOS("
            SqlCad = SqlCad & "codtipclie, destipclie"
            SqlCad = SqlCad & ") "
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & Val(rstExportarSql!CODTIPCLIE & "") & ", "
            SqlCad = SqlCad & "'" & Trim(rstExportarSql!DESTIPCLIE & "") & "'"
            SqlCad = SqlCad & ")"
            
            cnBdCPlus.Execute SqlCad

            Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Tipos de Cliente de CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlTipoClienteCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    Resume
    Err.Clear
End Sub

Private Sub exportarSqlTiposComprobantesCP()
    On Error GoTo errExportarSqlTiposComprobantesCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODDOC FROM DOCUMENTOS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Tipos de Comprobantes CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.DOCUMENTOS"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaComprobante
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2CODDOC & "")
                
                .obtenerConfigComprobante
            End With
            
            With objSqlAyudaComprobante
                .inicializarEntidades
                
                .Codigo = objAyudaComprobante.Codigo
                .CodigoSunat = objAyudaComprobante.CodigoSunat
                .CodigoConcar = objAyudaComprobante.CodigoConcar
                .CodigoExterno = objAyudaComprobante.CodigoExterno
                
                .Descripcion = objAyudaComprobante.Descripcion
                .Abreviatura = objAyudaComprobante.Abreviatura
                
                .TipoComprobante = objAyudaComprobante.TipoComprobante
                .DebHab = objAyudaComprobante.DebHab
                .TransFerir = objAyudaComprobante.TransFerir
                .EsOficial = objAyudaComprobante.EsOficial
                
                .ParaRegVentas = objAyudaComprobante.ParaRegVentas
                .RVDehHab = objAyudaComprobante.RVDehHab
                .RVEstados = objAyudaComprobante.RVEstados
                .RVTieneSerie = objAyudaComprobante.RVTieneSerie
                .RVLongitudSerie = objAyudaComprobante.RVLongitudSerie
                .RVTipoCadenaSerie = objAyudaComprobante.RVTipoCadenaSerie
                .RVTieneNumero = objAyudaComprobante.RVTieneNumero
                .RVLongitudNumero = objAyudaComprobante.RVLongitudNumero
                .RVTipoCadenaNumero = objAyudaComprobante.RVTipoCadenaNumero
                
                .ParaRegCompras = objAyudaComprobante.ParaRegCompras
                .RCDebHab = objAyudaComprobante.RCDebHab
                .RCEstados = objAyudaComprobante.RCEstados
                .RCTieneSerie = objAyudaComprobante.RCTieneSerie
                .RCLongitudSerie = objAyudaComprobante.RCLongitudSerie
                .RCTipoCadenaSerie = objAyudaComprobante.RCTipoCadenaSerie
                .RCTieneNumero = objAyudaComprobante.RCTieneNumero
                .RCLongitudNumero = objAyudaComprobante.RCLongitudNumero
                .RCTipoCadenaNumero = objAyudaComprobante.RCTipoCadenaNumero
                
                .CodigoExterno2 = objAyudaComprobante.CodigoExterno2
                .CodCompraRegistro = objAyudaComprobante.CodCompraRegistro
                
                If .guardarComprobante Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Tipos de Comprobantes CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlTiposComprobantesCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlFormaPagoCP()
    On Error GoTo errExportarSqlFormaPagoCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2FORPAG FROM EF2FORPAG"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Formas de Pago CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2FORPAG"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaFormaPago
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2FORPAG & "")
                
                .obtenerConfigFormaPago
            End With
            
            With objSqlAyudaFormaPago
                .inicializarEntidades
                
                .Codigo = objAyudaFormaPago.Codigo
                .CodigoS10 = objAyudaFormaPago.CodigoS10
                .Descripcion = objAyudaFormaPago.Descripcion
                .Dias = objAyudaFormaPago.Dias
                .Tipo = objAyudaFormaPago.Tipo
                .ParaLetra = objAyudaFormaPago.ParaLetra
                
                If .guardarFormaPago Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Formas de Pago CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlFormaPagoCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlProveedoresCP()
    On Error GoTo errExportarSqlProveedoresCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODPROV FROM EF2PROVEEDORES"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Proveedores CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2PROVEEDORES"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaProveedor
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2CODPROV & "")
                
                .obtenerConfigProveedor
            End With
            
            With objSqlAyudaProveedor
                .inicializarEntidades
                
                .OrigenProveedor = objAyudaProveedor.OrigenProveedor
                .ClaseProveedor = objAyudaProveedor.ClaseProveedor
                .CodigoTipoDocumento = objAyudaProveedor.CodigoTipoDocumento
                .NumeroDocumento = objAyudaProveedor.NumeroDocumento
                .NombreProveedor = Replace(objAyudaProveedor.NombreProveedor, "'", "´")
                .DireccionProveedor = Replace(objAyudaProveedor.DireccionProveedor, "'", "´")
                .CodigoDistrito = objAyudaProveedor.CodigoDistrito
                
                .Codigo = objAyudaProveedor.Codigo
                .CodigoExterno = objAyudaProveedor.CodigoExterno
                .NombreAbreviado = Replace(objAyudaProveedor.NombreAbreviado, "'", "´")
                .CodigoCategoria = objAyudaProveedor.CodigoCategoria
                
                .CodigoMoneda = objAyudaProveedor.CodigoMoneda
                .CuentaContable = objAyudaProveedor.CuentaContable
                .CodigoGasto = objAyudaProveedor.CodigoGasto
                .GrupoFlujo = objAyudaProveedor.GrupoFlujo
                .GrupoResultado = objAyudaProveedor.GrupoResultado
                
                .Telefono = objAyudaProveedor.Telefono
                .Fax = objAyudaProveedor.Fax
                .Email = objAyudaProveedor.Email
                .CuentaAbono = objAyudaProveedor.CuentaAbono
                .Contacto = Replace(objAyudaProveedor.Contacto, "'", "´")
                
                .ArticuloProveedor = objAyudaProveedor.ArticuloProveedor
                .CodTipoComprobante = objAyudaProveedor.CodTipoComprobante
                .EsAptoParaOrden = objAyudaProveedor.EsAptoParaOrden
                .CodigoFormaPago = objAyudaProveedor.CodigoFormaPago
                .Observacion = Replace(objAyudaProveedor.Observacion, "'", "´")
                
                .FechaReg = IIf(objAyudaProveedor.FechaReg = vbNullString, Format(Date, "Short Date"), objAyudaProveedor.FechaReg)
                .UsuarioReg = IIf(objAyudaProveedor.UsuarioReg = vbNullString, wusuario, objAyudaProveedor.UsuarioReg)
                .FechaMod = IIf(objAyudaProveedor.FechaMod = vbNullString, Format(Date, "Short Date"), objAyudaProveedor.FechaMod)
                .UsuarioMod = IIf(objAyudaProveedor.UsuarioMod = vbNullString, wusuario, objAyudaProveedor.UsuarioMod)
                
                If .guardarProveedor Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Proveedores CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlProveedoresCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlClientesCP()
    On Error GoTo errExportarSqlClientesCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODCLI FROM EF2CLIENTES"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Clientes CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2CLIENTES"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaCliente
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2CODCLI & "")
                
                .obtenerConfigCliente
            End With
            
            With objSqlAyudaCliente
                .inicializarEntidades
                
                .OrigenCliente = objAyudaCliente.OrigenCliente
                .ClaseCliente = objAyudaCliente.ClaseCliente
                .CodigoTipoDocumento = objAyudaCliente.CodigoTipoDocumento
                .NumeroDocumento = objAyudaCliente.NumeroDocumento
                .NombreCliente = Replace(objAyudaCliente.NombreCliente, "'", "´")
                .DireccionCliente = Replace(objAyudaCliente.DireccionCliente, "'", "´")
                .CodigoDistrito = objAyudaCliente.CodigoDistrito
                
                .Codigo = objAyudaCliente.Codigo
                '.CodigoExterno = objAyudaCliente.CodigoExterno
                .NombreAbreviado = Replace(objAyudaCliente.NombreAbreviado, "'", "´")
                .CodigoSector = objAyudaCliente.CodigoSector
                
                .DireccionRecepcion = Replace(objAyudaCliente.DireccionRecepcion, "'", "´")
                .DireccionCobranza = Replace(objAyudaCliente.DireccionCobranza, "'", "´")
                
                .Telefono = objAyudaCliente.Telefono
                .Fax = objAyudaCliente.Fax
                .Movil = objAyudaCliente.Movil
                .Email = objAyudaCliente.Email
                .EmailCotizacion = objAyudaCliente.EmailCotizacion
                .CuentaAbono = objAyudaCliente.CuentaAbono
                .Contacto = Replace(objAyudaCliente.Contacto, "'", "´")
                .EsAgenteRetencion = objAyudaCliente.EsAgenteRetencion
                
                .CodigoVendedor = objAyudaCliente.CodigoVendedor
                .PorcentajeComision = objAyudaCliente.PorcentajeComision
                .CodigoCobrador = objAyudaCliente.CodigoCobrador
                
                .CodigoFormaPago = objAyudaCliente.CodigoFormaPago
                .Observacion = Replace(objAyudaCliente.Observacion, "'", "´")
                
                .FechaReg = IIf(objAyudaCliente.FechaReg = vbNullString, Format(Date, "Short Date"), objAyudaCliente.FechaReg)
                .UsuarioReg = IIf(objAyudaCliente.UsuarioReg = vbNullString, wusuario, objAyudaCliente.UsuarioReg)
                .FechaMod = IIf(objAyudaCliente.FechaMod = vbNullString, Format(Date, "Short Date"), objAyudaCliente.FechaMod)
                .UsuarioMod = IIf(objAyudaCliente.UsuarioMod = vbNullString, wusuario, objAyudaCliente.UsuarioMod)
                
                If .guardarCliente Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Clientes CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlClientesCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlUsuariosCP()
    On Error GoTo errExportarSqlUsuariosCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODUSER FROM EF2USERS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL los Usuarios CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2USERS"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaUsuario
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2CODUSER & "")
                
                .obtenerConfigUsuario
            End With
            
            With objSqlAyudaUsuario
                .inicializarEntidades
                
                .Codigo = objAyudaUsuario.Codigo
                .CodigoExterno1 = objAyudaUsuario.CodigoExterno1
                .CodigoExterno2 = objAyudaUsuario.CodigoExterno2
                .NombreCompleto = objAyudaUsuario.NombreCompleto
                .Contrasena = objAyudaUsuario.Contrasena
                .Mail = objAyudaUsuario.Mail
                .SerieFacturacion = objAyudaUsuario.SerieFacturacion
                .Telefono = objAyudaUsuario.Telefono
                .Correo = objAyudaUsuario.Correo
                .Cargo = objAyudaUsuario.Cargo
                .UserMail = objAyudaUsuario.UserMail
                .PassMail = objAyudaUsuario.PassMail
                
                If .guardarUsuario Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL los Usuarios CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlUsuariosCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlTareasCP()
    On Error GoTo errExportarSqlTareasCP
    
    Dim rstExportarSql As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT F2CODTAREA FROM EF2TAREAS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Tareas de Usuario CP..."
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM MAESTROS.EF2TAREAS"
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            With objAyudaTarea
                .inicializarEntidades
                
                .Codigo = Trim(rstExportarSql!F2CODTAREA & "")
                
                .obtenerConfigTarea
            End With
            
            With objSqlAyudaTarea
                .inicializarEntidades
                
                .Codigo = objAyudaTarea.Codigo
                .Descripcion = objAyudaTarea.Descripcion
                
                If .guardarTarea Then
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Tareas de Usuario CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT * FROM EF2TAREAUSERS"
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    rstExportarSql.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstExportarSql.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstExportarSql)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Tareas de Usuario CP..."
        
        cnBdCPlus.Execute "DELETE FROM MAESTROS.EF2TIPOS"
        
        Actualiza_Log " < Replicación BD > DELETE FROM MAESTROS.EF2TAREAUSERS", StrConexDbBancos
        
        Do While Not rstExportarSql.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO MAESTROS.EF2TAREAUSERS("
            SqlCad = SqlCad & "F2CODUSER, F2CODTAREA"
            SqlCad = SqlCad & ") "
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstExportarSql!F2CODUSER & "") & "', "
            SqlCad = SqlCad & "'" & Trim(rstExportarSql!F2CODTAREA & "") & "'"
            SqlCad = SqlCad & ")"
            
            cnBdCPlus.Execute SqlCad

            Actualiza_Log " < Replicación BD > " & SqlCad, StrConexDbBancos
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Tareas de Usuario CP... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstExportarSql.MoveNext
        Loop
    End If
    
    If rstExportarSql.State = 1 Then rstExportarSql.Close
    
    Set rstExportarSql = Nothing
    
    MsgBox "Termine de procesar.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errExportarSqlTareasCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub recuperarOrdenCompraBackupCP()
    On Error GoTo errRecuperarOrdenCompraBackupCP
    
    Dim rstRecuperarCab As New ADODB.Recordset
    Dim rstRecuperarDet As New ADODB.Recordset
    
    Dim intAnno As Integer
    Dim intMes As Integer
    
    intAnno = Val(InputBox("Ingrese la Año que desea Exportar:", "Exportacion", Year(Date)))
    
    intMes = Val(InputBox("Ingrese la Mes que desea Exportar:", "Exportacion", Month(Date)))
    
    If intAnno < 2014 Or intAnno > Year(Date) Then
        intAnno = 0
    End If
    
    If intMes < 1 Or intMes > 12 Then
        intMes = 0
    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF4ORDEN "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F4LOCAL = 'OC' AND "
    SqlCad = SqlCad & "F4ESTADO <> 1 "
        
        If intAnno <> 0 Then
            SqlCad = SqlCad & "AND YEAR(F4FECEMI) = " & intAnno & " "
        End If
        
        If intAnno <> 0 And intMes <> 0 Then
            SqlCad = SqlCad & "AND MONTH(F4FECEMI) = " & intMes & " "
        End If
        
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "F4LOCAL, "
    SqlCad = SqlCad & "F4NUMORD"
    
    abrirCnnnDbBancosBackup
    
    If rstRecuperarCab.State = 1 Then rstRecuperarCab.Close
    
    rstRecuperarCab.Open SqlCad, cnn_dbbancos_backup, adOpenForwardOnly, adLockReadOnly
    
    If Not rstRecuperarCab.EOF Then
        fraProceso.Visible = True
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstRecuperarCab)
        pgbProceso.Value = 0
        fraProceso.Caption = "Exportando a SQL las Ordenes de Compra CP..."
        
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM PROCESOS.IF3ORDEN"
'
'        cnBdCPlus.Execute SqlCad
'
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "DELETE FROM PROCESOS.IF4ORDEN"
'
'        cnBdCPlus.Execute SqlCad
        
        Do While Not rstRecuperarCab.EOF
            With objAyudaOrden
                .inicializarEntidades
                
                .TipoOrden = Trim(rstRecuperarCab!F4LOCAL & "")
                .NumeroOrden = Trim(rstRecuperarCab!F4NUMORD & "")
                
                .FechaEmision = Format(Trim(rstRecuperarCab!F4FECEMI & ""), "Short Date")
                .SinProveedorEspecifico = CBool(rstRecuperarCab!F4SINPROVEEDORESP)
                .NomProveedor = Replace(Trim(rstRecuperarCab!F4NOMPROV & ""), "'", "´", 1)
                .RucProveedor = Trim(rstRecuperarCab!F4CODPRV & "")
                .CodProveedor = Trim(rstRecuperarCab!F4CODCLI & "")
                .ContactoProveedor = Replace(Trim(rstRecuperarCab!F4CONTACTO & ""), "'", "´", 1)
                
                .CodTipoComprobante = Trim(rstRecuperarCab!F4TIPDOC & "")
                .OrdenRegularizada = CBool(Val(rstRecuperarCab!F4REGULARIZA & ""))
                
                .FechaEntrega = Format(Trim(rstRecuperarCab!F4FECENT & ""), "Short Date")
                .CodigoSolicitante = Trim(rstRecuperarCab!F4CODSOL & "")
                .CodFormaPago = Trim(rstRecuperarCab!F4FORPAG & "")
                .CentroCosto = Trim(rstRecuperarCab!F4CENTRO & "")
                .LugarEntrega = Trim(rstRecuperarCab!F4LUGAR_ENTREGA & "")
                .PagoParcial = CBool(rstRecuperarCab!F4PAGOPARCIAL & "")
                
                .CodMoneda = Trim(rstRecuperarCab!F4TIPMON & "")
                .TipoCambio = Format(Val(rstRecuperarCab!F4TIPCAM & ""), "#.000")
                .NumeroCotizacion = Trim(rstRecuperarCab!F4NUMCOTIZA & "")
                
                .Colocada = CBool(rstRecuperarCab!F4COLOCADA)
                    .ColocadaUsuario = Trim(rstRecuperarCab!F4COLOCADAUSER & "")
                    .ColocadaFecha = IIf(.Colocada, Format(Trim(rstRecuperarCab!F4COLOCADAFECHA & ""), "Short Date"), vbNullString)
                
                .Atendida = CBool(rstRecuperarCab!F4ATENDIDA)
                    .AtendidaUsuario = Trim(rstRecuperarCab!F4ATENDIDAUSER & "")
                    .AtendidaFecha = IIf(.Atendida, Format(Trim(rstRecuperarCab!F4ATENDIDAFECHA & ""), "Short Date"), vbNullString)
                
                .Empresa = Trim(rstRecuperarCab!F4EMPRESA & "")
                .Observacion = Trim(rstRecuperarCab!F4OBSERVA & "")
        
                .SUBTOTAL = Val(Format(Val(rstRecuperarCab!F4BASIMP & ""), "0.00"))
                .TotalInafecto = Val(Format(Val(rstRecuperarCab!F4MONINA & ""), "0.00"))
                .TotalImpuesto = Val(Format(Val(rstRecuperarCab!F4IGV & ""), "0.00"))
                .TotalFacturado = Val(Format(Val(rstRecuperarCab!F4MONTO & ""), "0.00"))
                
                .FechaReg = Format(Trim(rstRecuperarCab!F4FECGRA & ""), "Short Date")
                .UsuarioReg = Trim(rstRecuperarCab!F4USEGRA & "")
                .FechaMod = Format(IIf(Trim(rstRecuperarCab!F4FECMOD & "") = vbNullString, Trim(rstRecuperarCab!F4FECGRA & ""), Trim(rstRecuperarCab!F4FECMOD & "")), "Short Date")
                .UsuarioMod = IIf(Trim(rstRecuperarCab!F4USEMOD & "") = vbNullString, Trim(rstRecuperarCab!F4USEGRA & ""), Trim(rstRecuperarCab!F4USEMOD & ""))
                
                .Estado = Val(rstRecuperarCab!F4ESTADO & "")
                
                If .guardarOrden Then
                    Actualiza_Log " < Recuperación Backup > " & .SQLSelectAlter, StrConexDbBancos
                    
                    .SQLSelectAlter = "DELETE FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'"
        
                    cnn_dbbancos.Execute .SQLSelectAlter
                    
                    Actualiza_Log " < Recuperación Backup > " & .SQLSelectAlter, StrConexDbBancos
        
                    If rstRecuperarDet.State = 1 Then rstRecuperarDet.Close
                    
                    rstRecuperarDet.Open "SELECT * FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'", cnn_dbbancos_backup, adOpenForwardOnly, adLockReadOnly  'AND VAL(F3CANPRO & '') > 0
                    
                    If Not rstRecuperarDet.EOF Then
                        rstRecuperarDet.MoveFirst
        
                        Do While Not rstRecuperarDet.EOF
                            .inicializarEntidadesDetalle
        
                            .ITEM = Val(rstRecuperarDet!ITEM & "")
                            .Requerimiento = Trim(rstRecuperarDet!COD_SOLICITUD & "")
                            .CodigoProducto = Trim(rstRecuperarDet!F3CODPRO & "")
                            .CodigoFabricante = Trim(rstRecuperarDet!F3CODFAB & "")
                            .NombreProducto = Trim(rstRecuperarDet!F5NOMPRO & "")
                            .NombreProductoInterno = Trim(rstRecuperarDet!F5NOMPRO_ING & "")
                            .CodigoUM = Trim(rstRecuperarDet!UNIDAD & "")
                            .Cantidad = Val(rstRecuperarDet!F3CANPRO & "")
                            .CantidadMaxima = Val(rstRecuperarDet!F3CANPRO2 & "")
                            .CantidadFaltante = Val(rstRecuperarDet!F3CANFAL & "")
        
                            .PorcentajeDemasia = Val(rstRecuperarDet!F3PORCDEMASIA & "")
        
                            .PrecioSinImpuesto = Val(rstRecuperarDet!F3PRECOS & "")
                            .PrecioConImpuesto = Val(rstRecuperarDet!F3PREUNI & "")
                            .PrecioNetoSinImpuesto = Val(rstRecuperarDet!F3PRENETO & "")
        
                            .PorcentajeDscto = Val(rstRecuperarDet!F3PORDCT & "")
                            .TotalDscto = Val(rstRecuperarDet!F3TOTDCT & "")
        
                            .Afecto = IIf(Trim(rstRecuperarDet!F5AFECTO) = "*", True, False)
        
                            .BasePorItem = Val(rstRecuperarDet!F5VALVTA & "")
                            .ImpuestoPorItem = Val(rstRecuperarDet!F3IGV & "")
                            .TotalPorItem = Val(rstRecuperarDet!F3TOTAL & "")
                            
                            .CodigoColor = Trim(rstRecuperarDet!CODCOLOR & "")
                            .ObservacionPorItem = Trim(rstRecuperarDet!F3OBSERVA & "")
        
                            .ItemAjustado = CBool(rstRecuperarDet!F3AJUSTE)
        
                            .CodigoGasto = Trim(rstRecuperarDet!F3GASTO & "")
                            .CuentaContable = Trim(rstRecuperarDet!F3CUENTA & "")
        
                            .guardarOrdenDetalleOneByOne
        
                            Actualiza_Log " < Recuperación Backup > " & .SQLSelectAlter, StrConexDbBancos
        
                            rstRecuperarDet.MoveNext
                        Loop
                    End If
                Else
                    Actualiza_Log " < Recuperación Backup > Recuperación de Orden de Compra No. " & .NumeroOrden & " fallida.", StrConexDbBancos
                End If
            End With
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Exportando a SQL las Ordenes de Compra CP [" & intAnno & "-" & intMes & "]... " & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3) & " - Exportados: " & pgbProceso.Value & " / " & pgbProceso.Max
            
            rstRecuperarCab.MoveNext
        Loop
    End If
    
    If rstRecuperarCab.State = 1 Then rstRecuperarCab.Close
    If rstRecuperarDet.State = 1 Then rstRecuperarDet.Close
    
    Set rstRecuperarCab = Nothing
    Set rstRecuperarDet = Nothing
    
    Exit Sub
errRecuperarOrdenCompraBackupCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub analisisYCorreccionStockLibreActual()
    On Error GoTo errAnalisisYCorreccionStockLibreActual
    
    Dim rstStockLEA As New ADODB.Recordset
    Dim rstStockCEA As New ADODB.Recordset
    
    Dim dblStock As Double
    Dim dblCantidadProdObservado As Double
    
    Dim dblLiberarCompromiso As Double
    Dim strNumValeProceso1 As String
    
    Dim dblItem As Double
    
    Dim cmdOP As ADODB.Command
            
    Set cmdOP = New ADODB.Command
    
    With cmdOP
        .ActiveConnection = cnBdCPlus
        .CommandType = adCmdStoredProc
        '.CommandTimeout = "180"
        .CommandText = "Procesos.usp_ListarResumenStock"
        
        .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, "01")
        .Parameters.Append .CreateParameter("@CODFAMILIA", adVarChar, adParamInput, 10, vbNullString)
        .Parameters.Append .CreateParameter("@CODSUBFAMILIA", adVarChar, adParamInput, 10, vbNullString)
        .Parameters.Append .CreateParameter("@FILTROSENSITIVO", adVarChar, adParamInput, 255, vbNullString)
        .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, "tmpCPResumenStockADMIN")
        
        .Execute
    End With
    
    Set cmdOP = Nothing
    
    fraProceso.Visible = True
    fraProceso2.Visible = True
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT * FROM tmpCPResumenStockADMIN WHERE STOCKLIBRE < 0 AND STOCKCOMPROMETIDO > 0 ORDER BY STOCKLIBRE DESC"
    
    If rstStockLEA.State = 1 Then rstStockLEA.Close
    
    rstStockLEA.Open SqlCad, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstStockLEA.EOF Then
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstStockLEA)
        pgbProceso.Value = 0
        fraProceso.Caption = "Evaluando Stock Libre En Almacen (LEA) Actual..."
        
        Do While Not rstStockLEA.EOF
            dblLiberarCompromiso = Abs(Val(rstStockLEA!STOCKLIBRE & ""))
            
            Set cmdOP = New ADODB.Command
            
            With cmdOP
                .ActiveConnection = cnBdCPlus
                .CommandType = adCmdStoredProc
                .CommandText = "Procesos.usp_ListarDetalleStock"
                
                .Parameters.Append .CreateParameter("@CODALMACEN", adVarChar, adParamInput, 10, "01")
                .Parameters.Append .CreateParameter("@CODPRODUCTO", adVarChar, adParamInput, 50, Trim(rstStockLEA!f5codpro & ""))
                .Parameters.Append .CreateParameter("@TIPOSTOCK", adVarChar, adParamInput, 1, "F")
                .Parameters.Append .CreateParameter("@ESTADOSTOCK", adVarChar, adParamInput, 1, "C")
                .Parameters.Append .CreateParameter("@NROPEDIDO", adVarChar, adParamInput, 10, vbNullString)
                .Parameters.Append .CreateParameter("@NOMBRETABLA", adVarChar, adParamInput, 100, "tmpCPStockDetalleADMIN")
                
                .Execute
            End With
            
            Set cmdOP = Nothing
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT * FROM tmpCPStockDetalleADMIN ORDER BY FECHAENTREGA DESC, NROPEDIDO DESC"
            
            If rstStockCEA.State = 1 Then rstStockCEA.Close
            
            rstStockCEA.Open SqlCad, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
            
            If Not rstStockCEA.EOF Then
                pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstStockCEA)
                pgbProceso2.Value = 0
                fraProceso2.Caption = "Evaluando Stock Comprometido En Almacen (CEA) Actual..."
                
                Do While Not rstStockCEA.EOF
                    'Si la Cantidad a Liberar es mayor a CERO proceder con la Liberacion de Compromiso
                    If dblLiberarCompromiso > 0 Then
                        'El Compromiso sera deshecho a traves de un Vale de Compromiso (Ingreso a Almacen)
                        
                        'Accedemos a las propiedades y procesos de nuestro Objeto de la Clase Vale
                        With objAyudaVale
                            'Limpiamos los Atributos de la Clase - Cabecera
                            .inicializarEntidades
                            .inicializarEntidadesAdicionales
                            .inicializarEntidadesDetalle
                            
                            'Ingresamos los Datos de la Cabecera del Vale de Compromiso
                            .CodigoAlmacen = Trim(rstStockCEA!CodAlmacen & "")
                            .NumeroVale = strNumValeProceso1
                            .TipoVale = "I"
                            
                            .Fecha = Format(Date, "dd/mm/yyyy")
                            .CodigoOrigen = "XCS"
                            .TipoCambio = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", .Fecha, "F"))
        
                                If .TipoCambio = 0 Then
                                    .TipoCambio = "3.1"
                                End If
        
                            .CodigoMoneda = "S"
        
                            .referencia = wnomcia
                            .observaciones = "PROCESO DE CORRECCION DE STOCK LIBRE."
        
                            .FecReg = Format(Date, "Short Date")
                            .UsuReg = wusuario
                            .FecMod = Format(Date, "Short Date")
                            .UsuMod = wusuario
                            'Ejecutamos el Proceso de la Clase 'Guardar'
                            If .guardarVale Then
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                'Este Sub-Proceso solo manejara un Solo Numero de Vale para todos los Compromisos de Liberacion de Stock
                                'Si el Numero de Vale del Proceso esta en blanco
                                If strNumValeProceso1 = vbNullString Then
                                    'Capturamos el Numero de Vale del Primer Vale de Compromiso generado;
                                    'en este Vale se guardaran todos los Compromisos de Liberacion de Stock
                                    strNumValeProceso1 = .NumeroVale
        
                                    'Borrar Detalle de Vale, por unica vez.
                                    SqlCad = vbNullString
                                    SqlCad = "DELETE FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
                                    
                                    cnn_dbbancos.Execute SqlCad
                                    
                                    Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
                                End If
                                'Limpiamos los Atributos de la Clase - Detalle
                                .inicializarEntidadesDetalle
                                
                                .NumeroOrdenCompra = vbNullString
                                'Ingresamos el Nro Pedido que deseamos Liberar Parcial o Totalmente, con esto Restamos la Cantidad del Compromiso
                                .Requerimiento = Trim(rstStockCEA!NroPedido & "")
                                
                                .CodigoProducto = Trim(rstStockCEA!CodProducto & "")
                                .CodigoProductoOriginal = Trim(rstStockCEA!CodProducto & "")
                                'Ingresamos la Cantidad que deseamos Liberar, en negativo
                                .Cantidad = IIf(Val(rstStockCEA!Cantidad & "") >= dblLiberarCompromiso, dblLiberarCompromiso, Val(rstStockCEA!Cantidad & "")) * -1
                                
                                dblItem = dblItem + 1
        
                                .ITEM = dblItem
                                'Guardamos el Primer Registro del Detalle
                                .guardarValeDetalleOneByOne
        
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                'Limpiamos los Atributos de la Clase - Detalle
                                .inicializarEntidadesDetalle
        
                                dblItem = dblItem + 1
                                'Ingresamos como Nro Pedido en Blanco, con esto le Sumaremos al Stock Libre el Compromiso Anterior
                                .Requerimiento = vbNullString
                                
                                .CodigoProducto = Trim(rstStockCEA!CodProducto & "")
                                .CodigoProductoOriginal = Trim(rstStockCEA!CodProducto & "")
                                .Cantidad = IIf(Val(rstStockCEA!Cantidad & "") >= dblLiberarCompromiso, dblLiberarCompromiso, Val(rstStockCEA!Cantidad & ""))
                                .ITEM = dblItem
                                'Guardamos el Segundo Registro del Detalle
                                .guardarValeDetalleOneByOne
        
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                
                                dblLiberarCompromiso = dblLiberarCompromiso - .Cantidad
                            End If
                        End With
                    End If
                    
                    DoEvents
        
                    pgbProceso2.Value = pgbProceso2.Value + 1
                    fraProceso2.Caption = "Evaluando Stock Comprometido En Almacen (CEA) Actual ( " & pgbProceso2.Value & "/" & pgbProceso2.Max & " )..." & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3)
                    
                    If dblLiberarCompromiso = 0 Then Exit Do
                    
                    rstStockCEA.MoveNext
                Loop
            End If
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Evaluando Stock Libre En Almacen (LEA) Actual ( " & pgbProceso.Value & "/" & pgbProceso.Max & " )..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstStockLEA.MoveNext
        Loop
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                'REPLICAR REGISTRO A SQL
                replicarRedistribucionFisica objAyudaVale.TipoVale, _
                                                objAyudaVale.NumeroVale, _
                                                objAyudaVale.CodigoAlmacen
            End If
            
    End If
    
    MsgBox "Proceso Finalizado.", vbInformation + vbOKOnly, App.ProductName
    
    fraProceso.Visible = False
    fraProceso2.Visible = False
    
    Exit Sub
errAnalisisYCorreccionStockLibreActual:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub replicarRedistribucionFisica(ByVal strTipoVale As String, _
                                            ByVal strNumeroVale As String, _
                                            ByVal strCodigoAlmacen As String)
                                            
    On Error GoTo errReplicarRedistribucionFisica
    
    Dim rstExportarRedistFisicaSql As New ADODB.Recordset
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = strTipoVale
        .NumeroVale = strNumeroVale
        .CodigoAlmacen = strCodigoAlmacen
        
        .obtenerConfigVale
    End With
    
    With objSqlAyudaVale
        .inicializarEntidades
        
        .TipoVale = objAyudaVale.TipoVale
        .NumeroVale = objAyudaVale.NumeroVale
        .NumeroValeExterno = objAyudaVale.NumeroValeExterno
        
        .CodigoAlmacen = objAyudaVale.CodigoAlmacen
        .CodigoOrigen = objAyudaVale.CodigoOrigen
        
        .TipoPersona = objAyudaVale.TipoPersona
            .CodigoProveedor = objAyudaVale.CodigoProveedor
            
        .CentroCosto = objAyudaVale.CentroCosto
        .SerieGuia = objAyudaVale.SerieGuia
        .NumeroGuia = objAyudaVale.NumeroGuia
        
        .CodTipoComprobante = objAyudaVale.CodTipoComprobante
        .SerieDocumento = objAyudaVale.SerieDocumento
        .NumeroDocumento = objAyudaVale.NumeroDocumento
        
        If .CodTipoComprobante <> vbNullString And .NumeroDocumento <> vbNullString Then
            .FechaUltima = Format(objAyudaVale.FechaUltima, "Short Date")
        Else
            .FechaUltima = vbNullString
        End If
        
        .OrdenTrabajo = objAyudaVale.OrdenTrabajo
        
        .Fecha = Format(objAyudaVale.Fecha, "Short Date")
        .CodigoMoneda = objAyudaVale.CodigoMoneda
        .TipoCambio = objAyudaVale.TipoCambio
        
        .NumeroOrdenCompra = objAyudaVale.NumeroOrdenCompra
        .observaciones = objAyudaVale.observaciones
        
        .RegistroCompra = objAyudaVale.RegistroCompra
        
        .ExportarVale = objAyudaVale.ExportarVale
        
        .FecReg = Format(objAyudaVale.FecReg, "Short Date")
        .UsuReg = objAyudaVale.UsuReg
        .FecMod = Format(objAyudaVale.FecMod, "Short Date")
        .UsuMod = objAyudaVale.UsuMod
        
        If .guardarVale Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnBdCPlus.Execute .SQLSelectAlter
            
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            If rstExportarRedistFisicaSql.State = 1 Then rstExportarRedistFisicaSql.Close
            
            rstExportarRedistFisicaSql.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstExportarRedistFisicaSql.EOF Then
                rstExportarRedistFisicaSql.MoveFirst
                
                pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstExportarRedistFisicaSql)
                pgbProceso2.Value = 0
                fraProceso2.Caption = "Replicando Detalle de Vale..."
                
                Do While Not rstExportarRedistFisicaSql.EOF
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstExportarRedistFisicaSql!f5codpro & "")
                    .CodigoProductoOriginal = Trim(rstExportarRedistFisicaSql!F5CODPROORIGINAL & "")
                    .Cantidad = Val(rstExportarRedistFisicaSql!F3CANPRO & "")
                    .CantidadMaxima = Val(rstExportarRedistFisicaSql!F3SALPEP & "")
                    
                    .ValorVenta = Val(rstExportarRedistFisicaSql!F3VALVTA & "")
                    .IGV = Val(rstExportarRedistFisicaSql!F3IGV & "")
                    .TOTAL = Val(rstExportarRedistFisicaSql!F3TOTITE & "")
                    .ValorVentaDol = Val(rstExportarRedistFisicaSql!F3VALDOL & "")
                    .IgvDol = Val(rstExportarRedistFisicaSql!F3IGVDOL & "")
                    .TotalDol = Val(rstExportarRedistFisicaSql!F3TOTDOL & "")
                    
                    .Grupo = Trim(rstExportarRedistFisicaSql!F3GRUPO & "")
                    .ITEM = Val(rstExportarRedistFisicaSql!F3ITEM & "")
                    .NumeroOrdenCompra = Trim(rstExportarRedistFisicaSql!F4NUMORD & "")
                    .Requerimiento = Trim(rstExportarRedistFisicaSql!COD_SOLICITUD & "")
                    .PorcentajeDscto = Val(rstExportarRedistFisicaSql!F3PORCENTAJEDSCTO & "")
                    .MontoDscto = Val(rstExportarRedistFisicaSql!F3MONTODSCTO & "")
                    
                    .ObservacionesPorItem = Trim(rstExportarRedistFisicaSql!observaciones & "")
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    DoEvents
                    
                    pgbProceso2.Value = pgbProceso2.Value + 1
                    fraProceso2.Caption = "Replicando Detalle de Vale ( " & pgbProceso2.Value & "/" & pgbProceso2.Max & " )..." & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3)
                    
                    rstExportarRedistFisicaSql.MoveNext
                Loop
            End If
        Else
            Actualiza_Log " < Replicación BD > Importación de Vale No. " & .CodigoAlmacen & " / " & .NumeroVale & " fallido.", StrConexDbBancos
        End If
    End With
    
    If rstExportarRedistFisicaSql.State = 1 Then rstExportarRedistFisicaSql.Close
    
    Set rstExportarRedistFisicaSql = Nothing
    
    Exit Sub
errReplicarRedistribucionFisica:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - Replicacion Redistribución Fisica"
    
    Err.Clear
End Sub

Private Sub cantidadRegistrosSQLvsAccess()
    On Error GoTo errCantidadRegistrosSQLvsAccess
    
    Dim rstSQL As New ADODB.Recordset
    Dim rstAccess As New ADODB.Recordset
    
    fraProceso.Visible = True
    fraProceso2.Visible = True
    lstSQL.Clear
    lstAccess.Clear
    
    SqlCad = vbNullString
    
    'SQL
    Select Case cmbTipoComparacion.ListIndex
        Case 0 'Ordenes C/S - Cabecera
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(F4FECEMI) AS ANNO, "
            SqlCad = SqlCad & "MONTH(F4FECEMI) AS MES, "
            SqlCad = SqlCad & "COUNT(*) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "PROCESOS.IF4ORDEN "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(F4FECEMI), "
            SqlCad = SqlCad & "MONTH (F4FECEMI) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(F4FECEMI), "
            SqlCad = SqlCad & "MONTH (F4FECEMI)"
        Case 1 'Ordenes C/S - Detalle
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(CAB.F4FECEMI) AS ANNO, "
            SqlCad = SqlCad & "MONTH(CAB.F4FECEMI) AS MES, "
            SqlCad = SqlCad & "COUNT(DET.F4NUMORD) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "PROCESOS.IF3ORDEN AS DET "
            SqlCad = SqlCad & "LEFT JOIN PROCESOS.IF4ORDEN  AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECEMI), "
            SqlCad = SqlCad & "MONTH(CAB.F4FECEMI) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECEMI), "
            SqlCad = SqlCad & "MONTH (Cab.F4FECEMI)"
        Case 2 'Vales I/S - Cabecera
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(F4FECVAL) AS ANNO, "
            SqlCad = SqlCad & "MONTH(F4FECVAL) AS MES, "
            SqlCad = SqlCad & "COUNT(*) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "PROCESOS.IF4VALES "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(F4FECVAL), "
            SqlCad = SqlCad & "MONTH (F4FECVAL) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(F4FECVAL), "
            SqlCad = SqlCad & "MONTH (F4FECVAL)"
        Case 3 'Vales I/S - Detalle
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(CAB.F4FECVAL) AS ANNO, "
            SqlCad = SqlCad & "MONTH(CAB.F4FECVAL) AS MES, "
            SqlCad = SqlCad & "COUNT(DET.F4NUMVAL) As CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "PROCESOS.IF3VALES AS DET "
            SqlCad = SqlCad & "LEFT JOIN PROCESOS.IF4VALES  AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECVAL), "
            SqlCad = SqlCad & "MONTH (CAB.F4FECVAL) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECVAL), "
            SqlCad = SqlCad & "MONTH (CAB.F4FECVAL)"
    End Select
    
    If rstSQL.State = 1 Then rstSQL.Close
    
    rstSQL.Open SqlCad, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSQL.EOF Then
        lstSQL.Clear
        lstSQL.AddItem "PERIODO SQL" & Space(5) & "CANTIDAD"
        lstSQL.AddItem "================================="
        
        pgbProceso.Max = ModUtilitario.devuelveCantRegistros(rstSQL)
        pgbProceso.Value = 0
        fraProceso.Caption = "Resumen Comparativo SQL..."
        
        Do While Not rstSQL.EOF
            lstSQL.AddItem Trim(rstSQL!Anno & "") & "-" & Trim(rstSQL!mes & "") & Space(20) & Val(rstSQL!Cantidad & "")
            
            DoEvents
            
            pgbProceso.Value = pgbProceso.Value + 1
            fraProceso.Caption = "Resumen Comparativo SQL ( " & pgbProceso.Value & "/" & pgbProceso.Max & " )..." & FormatPercent(pgbProceso.Value / pgbProceso.Max, 3)
            
            rstSQL.MoveNext
        Loop
    End If
    
    If rstSQL.State = 1 Then rstSQL.Close
    
    Set rstSQL = Nothing
    
    SqlCad = vbNullString
    
    'Access
    Select Case cmbTipoComparacion.ListIndex
        Case 0 'Ordenes C/S - Cabecera
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(F4FECEMI) AS ANNO, "
            SqlCad = SqlCad & "MONTH(F4FECEMI) AS MES, "
            SqlCad = SqlCad & "COUNT(*) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF4ORDEN "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(F4FECEMI), "
            SqlCad = SqlCad & "MONTH (F4FECEMI) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(F4FECEMI), "
            SqlCad = SqlCad & "MONTH (F4FECEMI)"
        Case 1 'Ordenes C/S - Detalle
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(CAB.F4FECEMI) AS ANNO, "
            SqlCad = SqlCad & "MONTH(CAB.F4FECEMI) AS MES, "
            SqlCad = SqlCad & "COUNT(DET.F4NUMORD) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3ORDEN AS DET "
            SqlCad = SqlCad & "LEFT JOIN IF4ORDEN  AS CAB ON CAB.F4LOCAL = DET.F4LOCAL AND CAB.F4NUMORD = DET.F4NUMORD "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECEMI), "
            SqlCad = SqlCad & "MONTH(CAB.F4FECEMI) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECEMI), "
            SqlCad = SqlCad & "MONTH (Cab.F4FECEMI)"
        Case 2 'Vales I/S - Cabecera
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(F4FECVAL) AS ANNO, "
            SqlCad = SqlCad & "MONTH(F4FECVAL) AS MES, "
            SqlCad = SqlCad & "COUNT(*) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF4VALES "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(F4FECVAL), "
            SqlCad = SqlCad & "MONTH (F4FECVAL) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(F4FECVAL), "
            SqlCad = SqlCad & "MONTH (F4FECVAL)"
        Case 3 'Vales I/S - Detalle
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "YEAR(CAB.F4FECVAL) AS ANNO, "
            SqlCad = SqlCad & "MONTH(CAB.F4FECVAL) AS MES, "
            SqlCad = SqlCad & "COUNT(DET.F4NUMVAL) As CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3VALES AS DET "
            SqlCad = SqlCad & "LEFT JOIN IF4VALES  AS CAB ON CAB.F2CODALM = DET.F2CODALM AND CAB.F4NUMVAL = DET.F4NUMVAL "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECVAL), "
            SqlCad = SqlCad & "MONTH (CAB.F4FECVAL) "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "YEAR(CAB.F4FECVAL), "
            SqlCad = SqlCad & "MONTH (CAB.F4FECVAL)"
    End Select
    
    If rstAccess.State = 1 Then rstAccess.Close
    
    rstAccess.Open SqlCad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAccess.EOF Then
        lstAccess.Clear
        lstAccess.AddItem "PERIODO ACC." & Space(5) & "CANTIDAD"
        lstAccess.AddItem "================================="
        
        pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstAccess)
        pgbProceso2.Value = 0
        fraProceso2.Caption = "Resumen Comparativo Access..."
        
        Do While Not rstAccess.EOF
            lstAccess.AddItem Trim(rstAccess!Anno & "") & "-" & Trim(rstAccess!mes & "") & Space(20) & Val(rstAccess!Cantidad & "")
            
            DoEvents
            
            pgbProceso2.Value = pgbProceso2.Value + 1
            fraProceso2.Caption = "Resumen Comparativo Access ( " & pgbProceso2.Value & "/" & pgbProceso2.Max & " )..." & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3)
            
            
            rstAccess.MoveNext
        Loop
    End If
    
    If rstAccess.State = 1 Then rstAccess.Close
    
    Set rstAccess = Nothing
    
    MsgBox "Proceso Finalizado.", vbInformation + vbOKOnly, App.ProductName
    
    fraProceso.Visible = False
    fraProceso2.Visible = False
    
    Exit Sub
errCantidadRegistrosSQLvsAccess:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlOrdenCompraEspecificaCP()
    On Error GoTo errExportarSqlOrdenCompraEspecificaCP
    
    Dim rstExportarDetSql As New ADODB.Recordset
    
    With objAyudaOrden
        .inicializarEntidades
        .inicializarEntidadesDetalle
        
        .TipoOrden = Trim(txtTipoOrden.Text)
        .NumeroOrden = Trim(txtNumeroOrden.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Orden no existe, verifique los datos ingresados.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        .obtenerConfigOrden
    End With
    
    With objSqlAyudaOrden
        .inicializarEntidades
        
        .TipoOrden = objAyudaOrden.TipoOrden
        .NumeroOrden = objAyudaOrden.NumeroOrden
        
        .FechaEmision = Format(objAyudaOrden.FechaEmision, "Short Date")
        .SinProveedorEspecifico = objAyudaOrden.SinProveedorEspecifico
        .NomProveedor = Replace(objAyudaOrden.NomProveedor, "'", "´", 1)
        .RucProveedor = objAyudaOrden.RucProveedor
        .CodProveedor = objAyudaOrden.CodProveedor
        .ContactoProveedor = Replace(objAyudaOrden.ContactoProveedor, "'", "´", 1)
        
        .CodTipoComprobante = objAyudaOrden.CodTipoComprobante
        .OrdenRegularizada = objAyudaOrden.OrdenRegularizada
        
        .FechaEntrega = Format(objAyudaOrden.FechaEntrega, "Short Date")
        .CodigoSolicitante = objAyudaOrden.CodigoSolicitante
        .CodFormaPago = objAyudaOrden.CodFormaPago
        .CentroCosto = objAyudaOrden.CentroCosto
        .LugarEntrega = objAyudaOrden.LugarEntrega
        .PagoParcial = objAyudaOrden.PagoParcial

        .CodMoneda = objAyudaOrden.CodMoneda
        .TipoCambio = Format(objAyudaOrden.TipoCambio, "#.000")
        .NumeroCotizacion = objAyudaOrden.NumeroCotizacion
        
        .Colocada = objAyudaOrden.Colocada
            .ColocadaUsuario = objAyudaOrden.ColocadaUsuario
            .ColocadaFecha = IIf(.Colocada, Format(objAyudaOrden.ColocadaFecha, "Short Date"), vbNullString)
        
        .Atendida = objAyudaOrden.Atendida
            .AtendidaUsuario = objAyudaOrden.AtendidaUsuario
            .AtendidaFecha = IIf(.Atendida, Format(objAyudaOrden.AtendidaFecha, "Short Date"), vbNullString)
        
        .Empresa = objAyudaOrden.Empresa
        .Observacion = objAyudaOrden.Observacion

        .SUBTOTAL = Val(Format(objAyudaOrden.SUBTOTAL, "0.00"))
        .TotalInafecto = Val(Format(objAyudaOrden.TotalInafecto, "0.00"))
        .TotalImpuesto = Val(Format(objAyudaOrden.TotalImpuesto, "0.00"))
        .TotalFacturado = Val(Format(objAyudaOrden.TotalFacturado, "0.00"))
        
        .FechaReg = Format(objAyudaOrden.FechaReg, "Short Date")
        .UsuarioReg = objAyudaOrden.UsuarioReg
        .FechaMod = Format(objAyudaOrden.FechaMod, "Short Date")
        .UsuarioMod = objAyudaOrden.UsuarioMod
        
        .Estado = objAyudaOrden.Estado
        
        If .guardarOrden Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM PROCESOS.IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'"

            cnBdCPlus.Execute .SQLSelectAlter
            
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos

            If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
            
            rstExportarDetSql.Open "SELECT * FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly 'AND VAL(F3CANPRO & '') > 0
            
            If Not rstExportarDetSql.EOF Then
                rstExportarDetSql.MoveFirst
                
                fraProceso2.Visible = True
                pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstExportarDetSql)
                pgbProceso2.Value = 0
                fraProceso2.Caption = "Exportando a SQL las Ordenes de Compra CP..."
                                
                Do While Not rstExportarDetSql.EOF
                    .inicializarEntidadesDetalle

                    .ITEM = Val(rstExportarDetSql!ITEM & "")
                    .Requerimiento = Trim(rstExportarDetSql!COD_SOLICITUD & "")
                    .CodigoProducto = Trim(rstExportarDetSql!F3CODPRO & "")
                    .CodigoFabricante = Trim(rstExportarDetSql!F3CODFAB & "")
                    .NombreProducto = Trim(rstExportarDetSql!F5NOMPRO & "")
                    .NombreProductoInterno = Trim(rstExportarDetSql!F5NOMPRO_ING & "")
                    .CodigoUM = Trim(rstExportarDetSql!UNIDAD & "")
                    .Cantidad = Val(rstExportarDetSql!F3CANPRO & "")
                    .CantidadMaxima = Val(rstExportarDetSql!F3CANPRO2 & "")
                    .CantidadFaltante = Val(rstExportarDetSql!F3CANFAL & "")

                    .PorcentajeDemasia = Val(rstExportarDetSql!F3PORCDEMASIA & "")

                    .PrecioSinImpuesto = Val(rstExportarDetSql!F3PRECOS & "")
                    .PrecioConImpuesto = Val(rstExportarDetSql!F3PREUNI & "")
                    .PrecioNetoSinImpuesto = Val(rstExportarDetSql!F3PRENETO & "")

                    .PorcentajeDscto = Val(rstExportarDetSql!F3PORDCT & "")
                    .TotalDscto = Val(rstExportarDetSql!F3TOTDCT & "")

                    .Afecto = IIf(Trim(rstExportarDetSql!F5AFECTO) = "*", True, False)

                    .BasePorItem = Val(rstExportarDetSql!F5VALVTA & "")
                    .ImpuestoPorItem = Val(rstExportarDetSql!F3IGV & "")
                    .TotalPorItem = Val(rstExportarDetSql!F3TOTAL & "")
                    
                    .CodigoColor = Trim(rstExportarDetSql!CODCOLOR & "")
                    .ObservacionPorItem = Trim(rstExportarDetSql!F3OBSERVA & "")

                    .ItemAjustado = CBool(rstExportarDetSql!F3AJUSTE)

                    .CodigoGasto = Trim(rstExportarDetSql!F3GASTO & "")
                    .CuentaContable = Trim(rstExportarDetSql!F3CUENTA & "")

                    .guardarOrdenDetalleOneByOne

                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                                    
                    DoEvents
                    
                    pgbProceso2.Value = pgbProceso2.Value + 1
                    fraProceso2.Caption = "Exportando a SQL Detalle de Orden: " & .TipoOrden & " / " & .NumeroOrden & "... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3) & " - Exportados: " & pgbProceso2.Value & " / " & pgbProceso2.Max
                    
                    rstExportarDetSql.MoveNext
                Loop
            End If
        Else
            Actualiza_Log " < Replicación BD > Importación de Orden de Compra No. " & .NumeroOrden & " fallido.", StrConexDbBancos
        End If
        
        .inicializarEntidades
        .inicializarEntidadesDetalle
    End With
    
    With objAyudaOrden
        .inicializarEntidades
        .inicializarEntidadesDetalle
    End With
    
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarDetSql = Nothing
    
    txtTipoOrden.Text = vbNullString
    txtNumeroOrden.Text = vbNullString
    
    Exit Sub
errExportarSqlOrdenCompraEspecificaCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub exportarSqlValeEspecificoCP(Optional ByVal bolObviarMensajes As Boolean)
    On Error GoTo errExportarSqlValeEspecificoCP
    
    Dim rstExportarDetSql As New ADODB.Recordset
    
    With objAyudaVale
        .inicializarEntidades
        .inicializarEntidadesAdicionales
        .inicializarEntidadesDetalle
        
        .TipoVale = Trim(txtTipoVale.Text)
        .CodigoAlmacen = Trim(txtCodAlmacen.Text)
        .NumeroVale = Trim(txtNumeroVale.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Vale no existe, verifique los datos ingresados.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        .obtenerConfigVale
    End With
    
    With objSqlAyudaVale
        .inicializarEntidades
        
        .TipoVale = objAyudaVale.TipoVale
        .NumeroVale = objAyudaVale.NumeroVale
        .NumeroValeExterno = objAyudaVale.NumeroValeExterno
        
        .CodigoAlmacen = objAyudaVale.CodigoAlmacen
        .CodigoOrigen = objAyudaVale.CodigoOrigen
        
        .TipoPersona = objAyudaVale.TipoPersona
            .CodigoProveedor = objAyudaVale.CodigoProveedor
            
        .CentroCosto = objAyudaVale.CentroCosto
        .SerieGuia = objAyudaVale.SerieGuia
        .NumeroGuia = objAyudaVale.NumeroGuia
        
        .CodTipoComprobante = objAyudaVale.CodTipoComprobante
        .SerieDocumento = objAyudaVale.SerieDocumento
        .NumeroDocumento = objAyudaVale.NumeroDocumento
        
        If .CodTipoComprobante <> vbNullString And .NumeroDocumento <> vbNullString Then
            .FechaUltima = Format(objAyudaVale.FechaUltima, "Short Date")
        Else
            .FechaUltima = vbNullString
        End If
        
        .OrdenTrabajo = objAyudaVale.OrdenTrabajo
        
        .Fecha = Format(objAyudaVale.Fecha, "Short Date")
        .CodigoMoneda = objAyudaVale.CodigoMoneda
        .TipoCambio = objAyudaVale.TipoCambio
        
        .NumeroOrdenCompra = objAyudaVale.NumeroOrdenCompra
        .observaciones = objAyudaVale.observaciones
        
        .RegistroCompra = objAyudaVale.RegistroCompra
        
        .ExportarVale = objAyudaVale.ExportarVale
        
        .FecReg = Format(objAyudaVale.FecReg, "Short Date")
        .UsuReg = objAyudaVale.UsuReg
        .FecMod = Format(objAyudaVale.FecMod, "Short Date")
        .UsuMod = objAyudaVale.UsuMod
        
        If .guardarVale Then
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "DELETE FROM PROCESOS.IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'"
            
            cnBdCPlus.Execute .SQLSelectAlter
            
            Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
            
            If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
            
            rstExportarDetSql.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & .CodigoAlmacen & "' AND F4NUMVAL = '" & .NumeroVale & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
            
            If Not rstExportarDetSql.EOF Then
                rstExportarDetSql.MoveFirst
                
                fraProceso2.Visible = True
                pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstExportarDetSql)
                pgbProceso2.Value = 0
                fraProceso2.Caption = "Exportando a SQL los Vales de CP..."
                
                Do While Not rstExportarDetSql.EOF
                    .inicializarEntidadesDetalle
                    
                    .CodigoProducto = Trim(rstExportarDetSql!f5codpro & "")
                    .CodigoProductoOriginal = Trim(rstExportarDetSql!F5CODPROORIGINAL & "")
                    .Cantidad = Val(rstExportarDetSql!F3CANPRO & "")
                    .CantidadMaxima = Val(rstExportarDetSql!F3SALPEP & "")
                    
                    .ValorVenta = Val(rstExportarDetSql!F3VALVTA & "")
                    .IGV = Val(rstExportarDetSql!F3IGV & "")
                    .TOTAL = Val(rstExportarDetSql!F3TOTITE & "")
                    .ValorVentaDol = Val(rstExportarDetSql!F3VALDOL & "")
                    .IgvDol = Val(rstExportarDetSql!F3IGVDOL & "")
                    .TotalDol = Val(rstExportarDetSql!F3TOTDOL & "")
                    
                    .Grupo = Trim(rstExportarDetSql!F3GRUPO & "")
                    .ITEM = Val(rstExportarDetSql!F3ITEM & "")
                    .NumeroOrdenCompra = Trim(rstExportarDetSql!F4NUMORD & "")
                    .Requerimiento = Trim(rstExportarDetSql!COD_SOLICITUD & "")
                    .PorcentajeDscto = Val(rstExportarDetSql!F3PORCENTAJEDSCTO & "")
                    .MontoDscto = Val(rstExportarDetSql!F3MONTODSCTO & "")
                    
                    .ObservacionesPorItem = Trim(rstExportarDetSql!observaciones & "")
                    
                    .guardarValeDetalleOneByOne
                    
                    Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                    
                    DoEvents
                    
                    pgbProceso2.Value = pgbProceso2.Value + 1
                    fraProceso2.Caption = "Exportando a SQL Detalle de Vale CP No. " & .CodigoAlmacen & " / " & .NumeroVale & "... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3) & " - Exportados: " & pgbProceso2.Value & " / " & pgbProceso2.Max
                    
                    rstExportarDetSql.MoveNext
                Loop
            End If
        Else
            Actualiza_Log " < Replicación BD > Importación de Vale No. " & .CodigoAlmacen & " / " & .NumeroVale & " fallido.", StrConexDbBancos
        End If
        
        .inicializarEntidades
        .inicializarEntidadesAdicionales
        .inicializarEntidadesDetalle
    End With
    
    With objAyudaVale
        .inicializarEntidades
        .inicializarEntidadesAdicionales
        .inicializarEntidadesDetalle
    End With
    
    If rstExportarDetSql.State = 1 Then rstExportarDetSql.Close
    
    Set rstExportarDetSql = Nothing
    
    txtTipoVale.Text = vbNullString
    txtCodAlmacen.Text = vbNullString
    txtNumeroVale.Text = vbNullString
    
    Exit Sub
errExportarSqlValeEspecificoCP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub cambioDeAlmacenVale()
    On Error GoTo errCambioDeAlmacenVale
    
    Dim strNuevoNumeroVale As String
    
    If Trim(txtCambioTipoVale.Text) <> vbNullString And Trim(txtCambioCodAlmacen.Text) <> vbNullString And Trim(txtCambioNumeroVale.Text) <> vbNullString Then
        txtCambioTipo.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4TIPOVALE", "IF4VALES", "F4TIPOVALE", Trim(txtCambioTipoVale.Text), "T", "AND F2CODALM = '" & Trim(txtCambioCodAlmacen.Text) & "' AND F4NUMVAL = '" & Trim(txtCambioNumeroVale.Text) & "'")
        txtCambioIdNumero.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "NUMENSAM", "IF4VALES", "F4TIPOVALE", Trim(txtCambioTipoVale.Text), "T", "AND F2CODALM = '" & Trim(txtCambioCodAlmacen.Text) & "' AND F4NUMVAL = '" & Trim(txtCambioNumeroVale.Text) & "'")
        
        If Trim(txtCambioTipo.Text) = vbNullString Or Trim(txtCambioIdNumero.Text) = vbNullString Then
            MsgBox "Verifique los Datos del Vale.", vbInformation + vbOKOnly, App.ProductName
            
            txtCambioTipoVale.SetFocus
            
            Exit Sub
        End If
    ElseIf Trim(txtCambioTipo.Text) <> vbNullString And Trim(txtCambioIdNumero.Text) <> vbNullString Then
        txtCambioTipoVale.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4TIPOVALE", "IF4VALES", "F4TIPOVALE", Trim(txtCambioTipo.Text), "T", "AND NUMENSAM = '" & Trim(txtCambioIdNumero.Text) & "'")
        txtCambioCodAlmacen.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALM", "IF4VALES", "F4TIPOVALE", Trim(txtCambioTipo.Text), "T", "AND NUMENSAM = '" & Trim(txtCambioIdNumero.Text) & "'")
        txtCambioNumeroVale.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F4NUMVAL", "IF4VALES", "F4TIPOVALE", Trim(txtCambioTipoVale.Text), "T", "AND NUMENSAM = '" & Trim(txtCambioIdNumero.Text) & "'")
        
        If Trim(txtCambioTipoVale.Text) = vbNullString Or Trim(txtCambioCodAlmacen.Text) = vbNullString Or Trim(txtCambioNumeroVale.Text) = vbNullString Then
            MsgBox "Verifique los Datos del ID Externo.", vbInformation + vbOKOnly, App.ProductName
            
            txtCambioTipo.SetFocus
            
            Exit Sub
        End If
    Else
        MsgBox "Datos incompletos para realizar la acción de Cambio de Almacen, ingrese los Datos del Vale o del Id Externo.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    If txtCambioCodAlmacenNuevo.Text = vbNullString Or lblCambioAlmacen.Caption = vbNullString Then
        MsgBox "Ingrese el Codigo de Almacen Nuevo para el Movimiento de Almacén.", vbInformation + vbOKOnly, App.ProductName
        
        txtCambioCodAlmacenNuevo.SetFocus
        
        Exit Sub
    End If
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = Trim(txtCambioTipoVale.Text)
        .CodigoAlmacen = Trim(txtCambioCodAlmacen.Text)
        .NumeroVale = Trim(txtCambioNumeroVale.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Vale no existe, verifique los datos ingresados.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If .CodigoAlmacen = Trim(txtCambioCodAlmacenNuevo.Text) Then
            MsgBox "El Nuevo Almacén debe ser diferente al Vale de Origen, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        .obtenerConfigVale
        
        .TipoVale = Trim(txtCambioTipoVale.Text)
        .CodigoAlmacen = Trim(txtCambioCodAlmacenNuevo.Text)
        .NumeroVale = Trim(txtCambioNumeroVale.Text)
        
        strNuevoNumeroVale = vbNullString
        
        
        'Actualización de 'db_bancos.mdb'
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "IF4VALES "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "F2CODALM = '" & Trim(txtCambioCodAlmacenNuevo.Text) & "' "
            
            If .verificarExistencia Then
                strNuevoNumeroVale = .generarNumeroVale
                
                SqlCad = SqlCad & ", F4NUMVAL = '" & strNuevoNumeroVale & "' "
            End If
            
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '" & Trim(txtCambioCodAlmacen.Text) & "' AND "
        SqlCad = SqlCad & "F4NUMVAL = '" & Trim(txtCambioNumeroVale.Text) & "'"
        
        abrirCnnDbBancos
        
        cnn_dbbancos.Execute SqlCad
        
        Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
        
        
        'Actualización de 'BdCPlus'
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM "
        SqlCad = SqlCad & "PROCESOS.IF3VALES "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '" & Trim(txtCambioCodAlmacen.Text) & "' AND "
        SqlCad = SqlCad & "F4NUMVAL = '" & Trim(txtCambioNumeroVale.Text) & "'"
        
        abrirCnBdCPlus
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < REPLICACION > " & SqlCad, cnn_dbbancos.ConnectionString
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM "
        SqlCad = SqlCad & "PROCESOS.IF4VALES "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '" & Trim(txtCambioCodAlmacen.Text) & "' AND "
        SqlCad = SqlCad & "F4NUMVAL = '" & Trim(txtCambioNumeroVale.Text) & "'"
        
        abrirCnBdCPlus
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < REPLICACION > " & SqlCad, cnn_dbbancos.ConnectionString
        
        txtTipoVale.Text = Trim(txtCambioTipoVale.Text)
        txtCodAlmacen.Text = Trim(txtCambioCodAlmacenNuevo.Text)
            
            If strNuevoNumeroVale <> vbNullString Then
                txtNumeroVale.Text = strNuevoNumeroVale
            Else
                txtNumeroVale.Text = Trim(txtCambioNumeroVale.Text)
            End If
        
        exportarSqlValeEspecificoCP
        
        With objSqlAyudaVale
            .inicializarEntidades
            
            .TipoVale = Trim(txtCambioTipoVale.Text)
            .CodigoAlmacen = Trim(txtCambioCodAlmacenNuevo.Text)
                
                If strNuevoNumeroVale <> vbNullString Then
                    .NumeroVale = strNuevoNumeroVale
                Else
                    .NumeroVale = Trim(txtCambioNumeroVale.Text)
                End If
            
            If Not .verificarExistencia Then
                MsgBox "Lo sentimos, la replicación del Vale a la Base de Datos SQL a fallado, por favor solicitar al Usuario volver a Guardar "
                
                MsgBox "Lo sentimos, la replicación del Vale a la Base de Datos SQL a fallado, por favor solicitar al Usuario volver a Guardar el siguiente Vale:" & vbNewLine & _
                        "Almacén: " & Trim(txtCambioCodAlmacenNuevo.Text) & " - " & "Número: " & IIf(strNuevoNumeroVale = vbNullString, Trim(txtCambioNumeroVale.Text), strNuevoNumeroVale), vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            .inicializarEntidades
        End With
        
        'Actualización de 'BdStudioModa'
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & IIf(Trim(txtCambioTipo.Text) = "I", "INGRESO ", "SALIDA ")
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "IDALMACEN = '" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODALMEXTERNO", "EF2ALMACENES", "F2CODALM", Trim(txtCambioCodAlmacenNuevo.Text), "T") & "' "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & IIf(Trim(txtCambioTipo.Text) = "I", "IDINGRESO", "IDSALIDA") & " = '" & Trim(txtCambioIdNumero.Text) & "'"
        
        abrirCnDBMilano
        
        cnBdStudioModa.Execute SqlCad
        
        Actualiza_Log " < EXPORTACION > " & SqlCad, cnn_dbbancos.ConnectionString
    End With
    
    MsgBox "Cambio de Almacen completado, por favor solicitar al Usuario que solicito el Cambio, verificar el Cambio de Almacen del Vale y Stock de los Items en el Vale." & vbNewLine & _
            "Vale de Origen = Almacén: " & Trim(txtCambioCodAlmacen.Text) & " - " & "Número: " & Trim(txtCambioNumeroVale.Text) & vbNewLine & _
            "Nuevo Vale = Almacén: " & Trim(txtCambioCodAlmacenNuevo.Text) & " - " & "Número: " & IIf(strNuevoNumeroVale = vbNullString, Trim(txtCambioNumeroVale.Text), strNuevoNumeroVale), vbInformation + vbOKOnly, App.ProductName
    
    txtCambioTipoVale.Text = vbNullString
    txtCambioCodAlmacen.Text = vbNullString
    txtCambioNumeroVale.Text = vbNullString
    
    txtCambioTipo.Text = vbNullString
    txtCambioIdNumero.Text = vbNullString
    
    txtCambioCodAlmacenNuevo.Text = vbNullString
    lblCambioAlmacen.Caption = vbNullString
    
    Exit Sub
errCambioDeAlmacenVale:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub reinicioRegistroCompraDeValeIngreso()
    On Error GoTo errReinicioRegistroCompraDeValeIngreso
    
    If Trim(txtReinicioTipoVale.Text) = vbNullString Or Trim(txtReinicioCodAlmacen.Text) = vbNullString Or Trim(txtReinicioNumeroVale.Text) = vbNullString Then
        MsgBox "Datos incompletos para realizar la acción de Reinicio de Registro de Compra para Vale de Ingreso, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        Exit Sub
    End If
    
    With objAyudaVale
        .inicializarEntidades
        
        .TipoVale = Trim(txtReinicioTipoVale.Text)
        .CodigoAlmacen = Trim(txtReinicioCodAlmacen.Text)
        .NumeroVale = Trim(txtReinicioNumeroVale.Text)
        
        If Not .verificarExistencia Then
            MsgBox "Vale no existe, verifique los datos ingresados.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        .obtenerConfigVale
        
        'Actualización de 'db_bancos.mdb'
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "IF4VALES "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "F4REGCOM = NULL "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '" & Trim(txtReinicioCodAlmacen.Text) & "' AND "
        SqlCad = SqlCad & "F4NUMVAL = '" & Trim(txtReinicioNumeroVale.Text) & "'"
        
        abrirCnnDbBancos
        
        cnn_dbbancos.Execute SqlCad
        
        Actualiza_Log SqlCad, cnn_dbbancos.ConnectionString
        
        
        'Actualización de 'BdCPlus'
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "PROCESOS.IF4VALES "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "F4REGCOM = '' "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F2CODALM = '" & Trim(txtReinicioCodAlmacen.Text) & "' AND "
        SqlCad = SqlCad & "F4NUMVAL = '" & Trim(txtReinicioNumeroVale.Text) & "'"
        
        abrirCnBdCPlus
        
        cnBdCPlus.Execute SqlCad
        
        Actualiza_Log " < REPLICACION > " & SqlCad, cnn_dbbancos.ConnectionString
    End With
    
    MsgBox "Reinicio de Registro de Compra de Vale de Ingreso por Compra completado, solicite la verificación del Vale por el Usuario." & vbNewLine & _
            "Vale = Almacén: " & Trim(txtReinicioCodAlmacen.Text) & " - " & "Número: " & Trim(txtReinicioNumeroVale.Text), vbInformation + vbOKOnly, App.ProductName
    
    txtReinicioTipoVale.Text = vbNullString
    txtReinicioCodAlmacen.Text = vbNullString
    txtReinicioNumeroVale.Text = vbNullString
    
    Exit Sub
errReinicioRegistroCompraDeValeIngreso:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub actualizarCatYopHistorialCambiosOP()
    On Error GoTo errActualizarCatYopHistorialCambiosOP
    
    Dim rstHistorialCambioOP As New ADODB.Recordset
    
    Dim strCategoria As String
    Dim strNumeroOP As String
    
    If rstHistorialCambioOP.State = 1 Then rstHistorialCambioOP.Close
    
    rstHistorialCambioOP.Open "SELECT IDORDENPRODUCCION FROM SF1ORDENPRODUCCION_LOG GROUP BY IDORDENPRODUCCION", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstHistorialCambioOP.EOF Then
        rstHistorialCambioOP.MoveFirst
        
        fraProceso2.Visible = True
        pgbProceso2.Max = ModUtilitario.devuelveCantRegistros(rstHistorialCambioOP)
        pgbProceso2.Value = 0
        fraProceso2.Caption = "Actualizando Datos en Historial de Cambios..."
        
        Do While Not rstHistorialCambioOP.EOF
            strCategoria = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "IDCATEGORIATIPO", "ORDENPRODUCCION", "IDORDENPRODUCCION", Trim(rstHistorialCambioOP!IdOrdenProduccion & ""), "N")
            strCategoria = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "NOMBRE", "CATEGORIATIPO", "IDCATEGORIATIPO", strCategoria, "T")
            strNumeroOP = ModUtilitario.ObtenerCampoV2(cnBdStudioModa, "OP", "ORDENPRODUCCION", "IDORDENPRODUCCION", Trim(rstHistorialCambioOP!IdOrdenProduccion & ""), "N")
                    
            SqlCad = vbNullString
            SqlCad = SqlCad & "UPDATE "
            SqlCad = SqlCad & "SF1ORDENPRODUCCION_LOG "
            SqlCad = SqlCad & "SET "
            SqlCad = SqlCad & "CATEGORIA = '" & strCategoria & "', "
            SqlCad = SqlCad & "NUMEROOP = '" & strNumeroOP & "' "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "IDORDENPRODUCCION = '" & Trim(rstHistorialCambioOP!IdOrdenProduccion & "") & "'"
            
            cnn_dbbancos.Execute SqlCad
            
            DoEvents
            
            pgbProceso2.Value = pgbProceso2.Value + 1
            fraProceso2.Caption = "Actualizando Datos en Historial de Cambios... " & FormatPercent(pgbProceso2.Value / pgbProceso2.Max, 3) & " - Procesados: " & pgbProceso2.Value & " / " & pgbProceso2.Max
            
            rstHistorialCambioOP.MoveNext
        Loop
    End If
    
    MsgBox "Proceso de Actualizacion Culminado.", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errActualizarCatYopHistorialCambiosOP:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub cmdReinicioReiniciar_Click()
    reinicioRegistroCompraDeValeIngreso
End Sub

Private Sub txtCambioCodAlmacen_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub


Private Sub txtCambioCodAlmacenNuevo_DblClick()
    txtCambioCodAlmacenNuevo_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCambioCodAlmacenNuevo_GotFocus()
    txtCambioCodAlmacenNuevo.SelStart = 0: txtCambioCodAlmacenNuevo.SelLength = Len(txtCambioCodAlmacenNuevo.Text)
End Sub

Private Sub txtCambioCodAlmacenNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            wcod_alm = vbNullString
            
            ayuda_almacen.Show 1
            
            Unload ayuda_almacen
            
            If Len(Trim(wcod_alm)) > 0 Then
                txtCambioCodAlmacenNuevo.Text = wcod_alm
                lblCambioAlmacen.Caption = wnomalmacen
                
                txtCambioCodAlmacenNuevo_KeyPress 13
            End If
    End Select
End Sub

Private Sub txtCambioCodAlmacenNuevo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCambioCodAlmacenNuevo_LostFocus()
    txtCambioCodAlmacenNuevo.Text = Format(txtCambioCodAlmacenNuevo.Text, "00")
    
    lblCambioAlmacen.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", txtCambioCodAlmacenNuevo.Text, "T")
End Sub

Private Sub txtCambioIdNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCambioNumeroVale_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub


Private Sub txtCambioTipo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
End Sub


Private Sub txtCambioTipoVale_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
End Sub


Private Sub txtCodAlmacen_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtNumeroOrden_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub


Private Sub txtNumeroVale_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub


Private Sub txtTipoOrden_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
End Sub


Private Sub txtTipoVale_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaTexto(KeyAscii)
End Sub


