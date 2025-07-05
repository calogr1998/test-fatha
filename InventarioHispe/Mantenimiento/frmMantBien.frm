VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMantBien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Producto y/o Servicio"
   ClientHeight    =   9030
   ClientLeft      =   210
   ClientTop       =   1695
   ClientWidth     =   9600
   Icon            =   "frmMantBien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   9600
   Begin ActiveToolBars.SSActiveToolBars stlbBien 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   4
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmMantBien.frx":058A
      ToolBars        =   "frmMantBien.frx":38A0
   End
   Begin TabDlg.SSTab tabBien 
      Height          =   8415
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   14843
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Principal"
      TabPicture(0)   =   "frmMantBien.frx":3982
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dbgMedidaAlterna"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPrincipal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Adicional"
      TabPicture(1)   =   "frmMantBien.frx":399E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAdicionales"
      Tab(1).Control(1)=   "dbgBienAlterno"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Config. Contable"
      TabPicture(2)   =   "frmMantBien.frx":39BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCodGasto"
      Tab(2).Control(1)=   "txtCtaContableInventa"
      Tab(2).Control(2)=   "txtCtaContableVenta"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(6)=   "lblCtaContableInventa"
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(8)=   "lblCtaContableVenta"
      Tab(2).Control(9)=   "Label20"
      Tab(2).ControlCount=   10
      Begin VB.TextBox txtCodGasto 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72840
         TabIndex        =   72
         Text            =   "Text3"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtCtaContableInventa 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72840
         TabIndex        =   71
         Text            =   "Text3"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtCtaContableVenta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -72840
         TabIndex        =   70
         Text            =   "Text3"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   " Ventas "
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
         Left            =   -74880
         TabIndex        =   59
         Top             =   2160
         Width           =   9135
         Begin VB.TextBox txtAnexoImpVta 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   63
            Text            =   "Text3"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtAnexoVta 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   62
            Text            =   "Text3"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtCtaContableImpVta 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   61
            Text            =   "Text3"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtCtaContableVta 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   60
            Text            =   "Text3"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Anexo Importación"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Anexo"
            Height          =   255
            Left            =   720
            TabIndex        =   68
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Cta. Cont. Importación"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblCtaContableImpVta 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3600
            TabIndex        =   66
            Top             =   960
            Width           =   5295
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Cta. Contable"
            Height          =   255
            Left            =   720
            TabIndex        =   65
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblCtaContableVta 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3600
            TabIndex        =   64
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Compras "
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
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtCtaContable 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   52
            Text            =   "Text3"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtCtaContableImp 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   51
            Text            =   "Text3"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtAnexo 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   50
            Text            =   "Text3"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtAnexoImp 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   49
            Text            =   "Text3"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblCtaContable 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3600
            TabIndex        =   58
            Top             =   240
            Width           =   5295
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Cta. Contable"
            Height          =   255
            Left            =   720
            TabIndex        =   57
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblCtaContableImp 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3600
            TabIndex        =   56
            Top             =   960
            Width           =   5295
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Cta. Cont. Importación"
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Anexo"
            Height          =   255
            Left            =   720
            TabIndex        =   54
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Anexo Importación"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.Frame fraAdicionales 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtMarcaModelo 
            Height          =   285
            Left            =   1680
            TabIndex        =   82
            Text            =   "Text3"
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox txtTalla 
            Height          =   285
            Left            =   1680
            TabIndex        =   80
            Text            =   "Text3"
            Top             =   3000
            Width           =   1815
         End
         Begin VB.TextBox txtColor 
            Height          =   285
            Left            =   1680
            TabIndex        =   78
            Text            =   "Text3"
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox txtStockReposicion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7320
            TabIndex        =   45
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtPorcentajeDemasia 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   39
            Text            =   "Text3"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtCodFab 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtDescripcionFab 
            Height          =   525
            Left            =   1680
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Text            =   "frmMantBien.frx":39D6
            Top             =   600
            Width           =   7095
         End
         Begin VB.TextBox txtAlmacen 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Text            =   "Text3"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtStockMinimo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Text            =   "Text3"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtStockMaximo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   16
            Text            =   "Text3"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtCodCentro 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   17
            Text            =   "Text3"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Marca Modelo"
            Height          =   255
            Left            =   360
            TabIndex        =   83
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Talla"
            Height          =   255
            Left            =   360
            TabIndex        =   81
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Color"
            Height          =   255
            Left            =   360
            TabIndex        =   79
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Stock Reposicion"
            Height          =   255
            Left            =   5880
            TabIndex        =   46
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Porcentaje Demasia (%)"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Codigo Fabricante"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Descripción de Fabricante"
            Height          =   375
            Left            =   360
            TabIndex        =   36
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Almacen"
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblAlmacen 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3240
            TabIndex        =   34
            Top             =   1200
            Width           =   4215
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Stock Minimo"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Stock Maximo"
            Height          =   255
            Left            =   3120
            TabIndex        =   32
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Centro de Costo"
            Height          =   255
            Left            =   600
            TabIndex        =   31
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblCentroCosto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3600
            TabIndex        =   30
            Top             =   2280
            Width           =   5295
         End
      End
      Begin VB.Frame fraPrincipal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   9135
         Begin VB.CheckBox chkTieneMovEnAlm 
            Caption         =   "Bien genera Movimiento en Almacen."
            Height          =   255
            Left            =   1680
            TabIndex        =   47
            Top             =   4440
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CheckBox chkEsInsumoParaOP 
            Caption         =   "Producto para Orden de Producción."
            Height          =   255
            Left            =   5640
            TabIndex        =   44
            Top             =   3960
            Width           =   3135
         End
         Begin VB.TextBox txtSubFamilia 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Text            =   "Text3"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtUnidadMedida 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Text            =   "Text3"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtModelo 
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Text            =   "Text3"
            Top             =   3000
            Width           =   3975
         End
         Begin VB.TextBox txtMarca 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Text            =   "Text3"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CheckBox chkBienImportado 
            Caption         =   "Bien Importado."
            Height          =   255
            Left            =   5640
            TabIndex        =   11
            Top             =   3720
            Width           =   1815
         End
         Begin VB.ComboBox cmbTipoExistencia 
            Height          =   315
            ItemData        =   "frmMantBien.frx":39DC
            Left            =   1680
            List            =   "frmMantBien.frx":39E6
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CheckBox chkParaVenta 
            Caption         =   "Bien disponible para Venta."
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            Top             =   4200
            Width           =   3615
         End
         Begin VB.CheckBox chkDescontinuado 
            Caption         =   "Bien descontinuado (Inhabilitar)."
            Height          =   255
            Left            =   1680
            TabIndex        =   9
            Top             =   3960
            Width           =   2775
         End
         Begin VB.CheckBox chkAfectoIGV 
            Caption         =   "Bien Afecto a I.G.V."
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   3720
            Width           =   1815
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   765
            Left            =   1680
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Text            =   "frmMantBien.frx":39FE
            Top             =   720
            Width           =   7095
         End
         Begin MSComCtl2.DTPicker dtpFecIngreso 
            Height          =   300
            Left            =   7440
            TabIndex        =   1
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   117702657
            CurrentDate     =   41780
         End
         Begin VB.Label lblSubFamilia 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3240
            TabIndex        =   43
            Top             =   2640
            Width           =   4215
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Sub-Familia"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Unidad de Medida"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblUnidadMedida 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3240
            TabIndex        =   27
            Top             =   1920
            Width           =   4215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Modelo"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Marca"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblMarca 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   285
            Left            =   3240
            TabIndex        =   24
            Top             =   2280
            Width           =   4215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo de Existencia"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha de Ingreso"
            Height          =   255
            Left            =   5880
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Descripción Interna"
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Codigo"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dbgMedidaAlterna 
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "frmMantBien.frx":3A04
         TabIndex        =   38
         Top             =   5400
         Width           =   9135
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dbgBienAlterno 
         Height          =   3735
         Left            =   -74880
         OleObjectBlob   =   "frmMantBien.frx":62C6
         TabIndex        =   41
         Top             =   4320
         Visible         =   0   'False
         Width           =   9135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. Gasto"
         Height          =   255
         Left            =   -74040
         TabIndex        =   77
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblCtaContableInventa 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   -71280
         TabIndex        =   76
         Top             =   4800
         Width           =   5295
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Contable Inventario"
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblCtaContableVenta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label5"
         Height          =   285
         Left            =   -71280
         TabIndex        =   74
         Top             =   4440
         Width           =   5295
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Contable Venta"
         Height          =   255
         Left            =   -74760
         TabIndex        =   73
         Top             =   4440
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMantBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda            As Boolean
Private bolNuevoBien        As Boolean
Private strCodBien          As String

Private objBien             As ClsBien


Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Public Property Let NuevoBien(ByVal Value As Boolean)
    bolNuevoBien = Value
End Property

Public Property Get NuevoBien() As Boolean
    NuevoBien = bolNuevoBien
End Property

Public Property Let Codigo(ByVal Value As String)
    strCodBien = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodBien
End Property

Private Sub listarTipoExistencia()
    objAyudaTipoExistencia.listarTipoExistencia cmbTipoExistencia, False, False
End Sub


Private Sub configurarGrillaMedidaAlterna()
   With dbgMedidaAlterna.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
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
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
    With dbgMedidaAlterna.Columns.ColumnByFieldName("F7CODMED").LookupColumn
        .LookupDataset.ADODataset.ConnectionString = cnn_dbbancos 'strCadenaConexioBdCPlus 'cnn_dbbancos
        .LookupDataset.ADODataset.CommandText = "SELECT F7CODMED,F7NOMMED FROM EF7MEDIDAS ORDER BY F7CODMED" '"SELECT F7CODMED,F7NOMMED FROM MAESTROS.EF7MEDIDAS ORDER BY F7CODMED"
        .LookupKeyField = "F7CODMED"
        .LookupResultField = "F7NOMMED"
        .LookupDataset.Active = True
        .ListFieldIndex = 0
        .DisplaySize = 15
        .LookupCache = True
        .ListFieldName = "F7NOMMED"
        .ListWidth = 50
    End With
    
    abrirCnTemporal
    
    With dbgMedidaAlterna.Dataset
        .ADODataset.ConnectionString = cnDBTemp
        .Active = False
        .Active = True
        .Close
        .Open
    End With
End Sub

Private Sub configurarGrillaBienAlterno()
   With dbgBienAlterno.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
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
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
    End With
    
    With dbgBienAlterno.Columns.ColumnByFieldName("CODIGO").LookupColumn
        .LookupDataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus 'cnn_dbbancos
        .LookupDataset.ADODataset.CommandText = "SELECT F5CODPRO, F5NOMPRO FROM MAESTROS.IF5PLA " & _
                                                "WHERE F5CODPRO NOT IN ('" & strCodBien & "') AND F5INSUMO = 'S' AND TIENEMOVENALM = 1 " & _
                                                "ORDER BY F5NOMPRO"
                                                
        .LookupKeyField = "F5CODPRO"
        .LookupResultField = "F5NOMPRO"
        .LookupDataset.Active = True
        .ListFieldIndex = 0
        .DisplaySize = 255
        .LookupCache = True
        .ListFieldName = "F5NOMPRO"
        .ListWidth = 200
    End With
    
    abrirCnTemporal
    
    With dbgBienAlterno.Dataset
        .ADODataset.ConnectionString = cnDBTemp
        .Active = False
        .Active = True
        .Close
        .Open
    End With
End Sub

Private Sub listarGrillaMedidaAlterna()
    With dbgMedidaAlterna.Dataset
        abrirCnTemporal
        
        .Active = False
        .Refresh
        
        abrirCnTemporal
        
        .ADODataset.ConnectionString = cnDBTemp
        .Active = False
        .Active = True
        .Close
        .Open
    
        dbgMedidaAlterna.Columns.ColumnByFieldName("F7CODMED").SummaryFooterType = cstCount
        dbgMedidaAlterna.Columns.ColumnByFieldName("F7CODMED").SummaryFooterFormat = "U.M. Alterno(s) = " & .RecordCount
    End With
End Sub

Private Sub adicionarItemMedidaAlterna()
    With dbgMedidaAlterna.Dataset
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE * FROM TMPMEDVENTA"
        
        .Active = False
        
        .Refresh

        abrirCnTemporal
        
        .ADODataset.ConnectionString = cnDBTemp
        
        .Active = True
        .Close
        .Open
        
        If .State = dsEdit Or .State = dsInsert Then
            .Append
        Else
            .Edit
        End If
        
        .FieldValues("ITEM") = .RecordCount + 1
        .FieldValues("F7CODMED") = vbNullString
        .FieldValues("DESCRIPCION") = vbNullString
        .FieldValues("FACTOR") = 0
        .FieldValues("F5PREVTA") = 0
        
        .Post
        
        .Close
        .Open
    
        dbgMedidaAlterna.Columns.ColumnByFieldName("F7CODMED").SummaryFooterType = cstCount
        dbgMedidaAlterna.Columns.ColumnByFieldName("F7CODMED").SummaryFooterFormat = "U.M. Alterno(s) = " & .RecordCount
    End With
End Sub

Private Sub listarGrillaBienAlterno()
    With dbgBienAlterno.Dataset
        abrirCnTemporal
        
        .Active = False
        .Refresh
        
        abrirCnTemporal
        
        .ADODataset.ConnectionString = cnDBTemp
        .Active = False
        .Active = True
        .Close
        .Open
        
        dbgBienAlterno.Columns.ColumnByFieldName("CODIGO").SummaryFooterType = cstCount
        dbgBienAlterno.Columns.ColumnByFieldName("CODIGO").SummaryFooterFormat = "Producto(s) Alterno(s) = " & .RecordCount
    End With
End Sub

Private Sub adicionarItemBienAlterno()
    With dbgBienAlterno.Dataset
        .Close
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE * FROM TMPBIENALTERNO"
        
        .Active = False
        
        .Refresh

        abrirCnTemporal
        
        .ADODataset.ConnectionString = cnDBTemp
        
        .Active = True
        .Close
        .Open
        
        If .State = dsEdit Or .State = dsInsert Then
            .Append
        Else
            .Edit
        End If
        
        .FieldValues("CODIGO") = vbNullString
        .FieldValues("ESTADO") = True
        
        .Post
        
        .Close
        .Open
    
        dbgBienAlterno.Columns.ColumnByFieldName("CODIGO").SummaryFooterType = cstCount
        dbgBienAlterno.Columns.ColumnByFieldName("CODIGO").SummaryFooterFormat = "Producto(s) Alterno(s) = " & .RecordCount
    End With
End Sub

Private Sub limpiarCajas()
    txtCodigo.Text = vbNullString
    txtDescripcion.Text = vbNullString
    dtpFecIngreso.Value = Date
    
    cmbTipoExistencia.ListIndex = -1
    txtUnidadMedida.Text = vbNullString
        lblUnidadMedida.Caption = vbNullString: lblUnidadMedida.BackColor = DF
    txtSubFamilia.Text = vbNullString
        lblSubFamilia.Caption = vbNullString: lblSubFamilia.BackColor = DF
    txtMarca.Text = vbNullString
        lblMarca.Caption = vbNullString: lblMarca.BackColor = DF
    txtModelo.Text = vbNullString
    
    chkAfectoIGV.Value = vbChecked
    chkDescontinuado.Value = vbUnchecked
    chkParaVenta.Value = vbUnchecked
    chkBienImportado.Value = vbUnchecked
    chkEsInsumoParaOP.Value = vbUnchecked
    chkTieneMovEnAlm.Value = vbChecked
    
    txtCodFab.Text = vbNullString
    txtDescripcionFab.Text = vbNullString
    txtAlmacen.Text = vbNullString
        lblAlmacen.Caption = vbNullString: lblAlmacen.BackColor = DF
    txtStockMinimo.Text = "0.00"
    txtStockMaximo.Text = "0.00"
    
    txtPorcentajeDemasia.Text = "0.00"
    
    
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
    
    
    txtCodGasto.Text = vbNullString
    txtCodCentro.Text = vbNullString
        lblCentroCosto.Caption = vbNullString: lblCentroCosto.BackColor = DF
    txtCtaContableVenta.Text = vbNullString
        lblCtaContableVenta.Caption = vbNullString: lblCtaContableVenta.BackColor = DF
    txtCtaContableInventa.Text = vbNullString
        lblCtaContableInventa.Caption = vbNullString: lblCtaContableInventa.BackColor = DF
    
    txtColor.Text = vbNullString
    txtTalla.Text = vbNullString
    txtMarcaModelo.Text = vbNullString
    
    txtCodigo.Locked = False
    txtCodigo.BackColor = HA
    
    tabBien.Tab = 0
    dbgMedidaAlterna.Enabled = False: dbgMedidaAlterna.BackColor = DH
    dbgBienAlterno.Enabled = False: dbgBienAlterno.BackColor = DH
End Sub

Private Sub consultarBien()
    Set objBien = New ClsBien
    
    limpiarCajas
    
    dbgMedidaAlterna.Dataset.Close
'    dbgBienAlterno.Dataset.Close
    
    With objBien
        .Codigo = strCodBien
        
        If .obtenerBien Then
            txtCodigo.Text = .Codigo
            txtDescripcion.Text = .Descripcion
                dbgBienAlterno.Bands.BandByName("bndPrincipal").Caption = .Descripcion
            
            dtpFecIngreso.Value = IIf(.FechaIngreso = vbNullString, Format(Date, "Short Date"), .FechaIngreso)
            cmbTipoExistencia.ListIndex = ModUtilitario.seleccionarItem(cmbTipoExistencia, .CodTipoExistencia, "DER", 2)
            
            txtUnidadMedida.Text = .CodUM
                lblUnidadMedida.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7NOMMED", "EF7MEDIDAS", "F7CODMED", .CodUM, "T")
            
            txtMarca.Text = .CodMarca
                lblMarca.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESMAR", "EF2MARCAS", "F2CODMAR", .CodMarca, "T")
            
            txtSubFamilia.Text = .CodigoSubFamilia
                lblSubFamilia.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7DESCON", "SF7NIVEL02", "F7CODCON", .CodigoSubFamilia, "T")
                
            txtModelo.Text = .Modelo
            
            chkAfectoIGV.Value = IIf(.Afecto, vbChecked, vbUnchecked)
            chkDescontinuado.Value = IIf(.Descontinuado, vbChecked, vbUnchecked)
            chkParaVenta.Value = IIf(.ParaVenta, vbChecked, vbUnchecked)
            chkBienImportado.Value = IIf(.EsImportado, vbChecked, vbUnchecked)
            chkEsInsumoParaOP.Value = IIf(.EsInsumoParaOP, vbChecked, vbUnchecked)
            chkTieneMovEnAlm.Value = IIf(.TieneMovimientoEnAlmacen, vbChecked, vbUnchecked)
            
            txtCodFab.Text = .CodigoFab
            txtDescripcionFab.Text = .DescripcionFab
            txtAlmacen.Text = .CodAlmacen
                lblAlmacen.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", .CodAlmacen, "T")
            
            txtStockMinimo.Text = Format(.StockMin, "#,0.00")
            txtStockMaximo.Text = Format(.StockMax, "#,0.00")
            txtStockReposicion.Text = Format(.StockReposicion, "#,0.00")
            
            txtPorcentajeDemasia.Text = Format(.PorcentajeDemasia, "#,0.00")
            
            
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
            
            
            txtCodGasto.Text = .CodGasto
            txtCodCentro.Text = .CodCentroCosto
                lblCentroCosto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", .CodCentroCosto, "T")
                
            txtCtaContableVenta.Text = .CtaContableVenta
                lblCtaContableVenta.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableVenta, "T")
            txtCtaContableInventa.Text = .CtaContableInventa
                lblCtaContableInventa.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", .CtaContableInventa, "T")
                
            txtColor.Text = .Color
            txtTalla.Text = .Talla
            txtMarcaModelo.Text = .MarcaModelo
            
            listarGrillaMedidaAlterna
            
'            listarGrillaBienAlterno
            
            txtCodigo.Locked = True
            txtCodigo.BackColor = DH
            
            dbgMedidaAlterna.Enabled = True: dbgMedidaAlterna.BackColor = HA
'            dbgBienAlterno.Enabled = True: dbgBienAlterno.BackColor = HA
        End If
    End With
    
    If dbgMedidaAlterna.Dataset.RecordCount = 0 Then
        dbgMedidaAlterna.Dataset.Close
        
        adicionarItemMedidaAlterna
    End If
    
'    If dbgBienAlterno.Dataset.RecordCount = 0 Then
'        dbgBienAlterno.Dataset.Close
'
'        adicionarItemBienAlterno
'    End If
    
    Set objBien = Nothing
End Sub

Private Sub validarCajas()
    If dbgMedidaAlterna.Dataset.State = dsEdit Or dbgMedidaAlterna.Dataset.State = dsInsert Then
        dbgMedidaAlterna.Dataset.Post
    End If
    
    If dbgBienAlterno.Dataset.State = dsEdit Or dbgBienAlterno.Dataset.State = dsInsert Then
        dbgBienAlterno.Dataset.Post
    End If
    
    If Trim(txtDescripcion.Text) = vbNullString Then
        MsgBox "El Campo Descripción es obligatorio.", vbCritical, App.ProductName
        
        txtDescripcion.SetFocus
        
        Exit Sub
    End If
    
    If cmbTipoExistencia.ListIndex = -1 Then
        MsgBox "Seleccione el Tipo de Existencia.", vbCritical, App.ProductName
        
        cmbTipoExistencia.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtUnidadMedida.Text) = vbNullString Then
        MsgBox "El Campo Unidad Medida es obligatorio.", vbCritical, App.ProductName
        
        txtUnidadMedida.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarBien
    End If
End Sub

Private Sub guardarBien()
    Set objBien = New ClsBien
    
    With objBien
        .Codigo = Trim(txtCodigo.Text)
        .Descripcion = Trim(txtDescripcion.Text)
        
        .CodTipoExistencia = right(cmbTipoExistencia.Text, 2)
        .AbrevTipoExistencia = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "ABREVIATURA", "EF2TIPOEXISTENCIA", "CODIGO", .CodTipoExistencia, "T")
        
        .CodUM = Trim(txtUnidadMedida.Text)
        .CodMarca = Trim(txtMarca.Text)
        .CodigoSubFamilia = Trim(txtSubFamilia.Text)
        
        .Modelo = Trim(txtModelo.Text)
        
        .Afecto = CBool(chkAfectoIGV.Value)
        .Descontinuado = CBool(chkDescontinuado.Value)
        .ParaVenta = CBool(chkParaVenta.Value)
        .EsImportado = CBool(chkBienImportado.Value)
        .EsInsumoParaOP = CBool(chkEsInsumoParaOP.Value)
        .TieneMovimientoEnAlmacen = CBool(chkTieneMovEnAlm.Value)
        
        .CodigoFab = Trim(txtCodFab.Text)
        .DescripcionFab = Trim(txtDescripcionFab.Text)
        .CodAlmacen = Trim(txtAlmacen.Text)
        
        .StockMin = Val(Format(txtStockMinimo.Text, "#0.00"))
        .StockMax = Val(Format(txtStockMaximo.Text, "#0.00"))
        .StockReposicion = Val(Format(txtStockReposicion.Text, "#0.00"))
        
        .PorcentajeDemasia = Val(Format(txtPorcentajeDemasia.Text, "#0.00"))
        
        .CtaContable = Trim(txtCtaContable.Text)
        .Anexo = Trim(txtAnexo.Text)
        .CtaContableImportacion = Trim(txtCtaContableImp.Text)
        .AnexoImportacion = Trim(txtAnexoImp.Text)
        
        .CtaContableVta = Trim(txtCtaContableVta.Text)
        .AnexoVta = Trim(txtAnexoVta.Text)
        .CtaContableImportacionVta = Trim(txtCtaContableImpVta.Text)
        .AnexoImportacionVta = Trim(txtAnexoImpVta.Text)
        
        .CodGasto = Trim(txtCodGasto.Text)
        .CodCentroCosto = Trim(txtCodCentro.Text)
            
        .CtaContableVenta = Trim(txtCtaContableVenta.Text)
        .CtaContableInventa = Trim(txtCtaContableInventa.Text)
        
        .Color = Trim(txtColor.Text)
        .Talla = Trim(txtTalla.Text)
        .MarcaModelo = Trim(txtMarcaModelo.Text)
        
        .FechaIngreso = Format(dtpFecIngreso.Value, "Short Date")
        .UsuarioIngreso = wusuario
        .FechaModificacion = Format(Date, "Short Date")
        .UsuarioModificacion = wusuario
        
        dbgMedidaAlterna.Dataset.Close
'        dbgBienAlterno.Dataset.Close
        
        If .guardarBien Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos
            
            .SQLSelectAlter = "SELECT F7CODMED, FACTOR, F5PREVTA FROM TMPMEDVENTA WHERE TRIM(F7CODMED & '') <> ''"
            
            .guardarMedidaAlterna
            
'            .SQLSelectAlter = "SELECT CODIGO, ESTADO FROM TMPBIENALTERNO WHERE TRIM(CODIGO & '') <> ''"
'
'            .guardarBienAlterno
            
            
            strCodBien = .Codigo
            
            consultarBien
            
            MsgBox "Registro Actualizado.", vbInformation + vbOKOnly, App.ProductName
        Else
            listarGrillaMedidaAlterna
            
'            listarGrillaBienAlterno
        End If
    End With
    
    Set objBien = Nothing
End Sub

Private Sub eliminarBien()
    Set objBien = New ClsBien
    
    With objBien
        .Codigo = Trim(txtCodigo.Text)
        
        If Not objBien.verificarExistencia Then
            MsgBox "Registro no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        If MsgBox("¿Desea Eliminar la Bien con Codigo No. " & .Codigo & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarBien Then
                strCodBien = .Codigo
                
                consultarBien
                
                MsgBox "Producto Eliminado.", vbInformation + vbOKOnly, App.ProductName
             End If
        End If
    End With
    
    Set objBien = Nothing
End Sub



Private Sub cmbTipoExistencia_LostFocus()
    If InStr(cmbTipoExistencia.Text, "SERVICIOS") = 1 Then
        txtMarca.Enabled = False
        txtMarca.BackColor = vbGrayText
        txtSubFamilia.Enabled = False
        txtSubFamilia.BackColor = vbGrayText
        chkTieneMovEnAlm.Value = 0
        chkTieneMovEnAlm.Enabled = False
    Else
        txtMarca.Enabled = True
        txtMarca.BackColor = vbRed
        txtSubFamilia.Enabled = True
    End If
End Sub

Private Sub dbgMedidaAlterna_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If Action = daInsert Then
        dbgMedidaAlterna.Dataset.Edit
        
        With dbgMedidaAlterna.Columns
            .ColumnByFieldName("F7CODMED").Value = vbNullString
            .ColumnByFieldName("FACTOR").Value = 0
            .ColumnByFieldName("F5PREVTA").Value = 0
            
            .FocusedIndex = 0
        End With
    End If
End Sub

Private Sub dbgMedidaAlterna_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    With dbgMedidaAlterna.Dataset
        If Action = daInsert Then
            If .RecordCount > 0 Then
                If Len(Trim(dbgMedidaAlterna.Columns.ColumnByFieldName("F7CODMED").Value & "")) = 0 And _
                    Len(Trim(dbgMedidaAlterna.Columns.ColumnByFieldName("FACTOR").Value & "")) = 0 Then
                    Allow = False
                Else
                    .Last
                End If
            End If
        End If
        
        If Action = daDelete Then
            .Refresh
            
            dbgMedidaAlterna.Columns.ColumnByFieldName("NOMCOMPLETO").SummaryFooterFormat = "Cant. de Inscritos = " & .RecordCount
        End If
    End With
End Sub

Private Sub dbgMedidaAlterna_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyF4
            If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                dbgMedidaAlterna.Dataset.Delete
                
                If dbgMedidaAlterna.Dataset.RecordCount > 0 Then
                    dbgMedidaAlterna.Dataset.Last
                    dbgMedidaAlterna.Columns.FocusedIndex = 0
                Else
                    dbgMedidaAlterna.Dataset.Close
                    
                    adicionarItemMedidaAlterna
                    
                    dbgMedidaAlterna.Dataset.Open
                End If
            End If
            
            dbgMedidaAlterna.SetFocus
        Case vbKeyReturn
            If dbgMedidaAlterna.Dataset.State = dsEdit Or dbgMedidaAlterna.Dataset.State = dsInsert Then
                dbgMedidaAlterna.Dataset.Post
                
                dbgMedidaAlterna.Dataset.Edit: dbgMedidaAlterna.Dataset.Post
            End If
        Case vbKeyEscape
            KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()
    abrirCnContaTabla
    
'    dbgMedidaAlterna.Dataset.Close
'    dbgBienAlterno.Dataset.Close
    Me.top = 1000
        Me.left = 1250
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPMEDVENTA"
    
'    abrirCnTemporal
'
'    cnDBTemp.Execute "DELETE FROM TMPBIENALTERNO"
    
'    dbgMedidaAlterna.Dataset.Open
'    dbgBienAlterno.Dataset.Open
    
    configurarGrillaMedidaAlterna
    
'    configurarGrillaBienAlterno
    
    listarTipoExistencia
    
    consultarBien
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    With frmListaBien
'        .listarBien
'
'        .Show
'    End With
End Sub

Private Sub stlbBien_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
            strCodBien = vbNullString
            
            consultarBien
        Case "Guardar"
            validarCajas
        Case "Eliminar"
            eliminarBien
        Case "Salir"
            'Programar
            'Me.Hide
            Unload Me
    End Select
End Sub

Private Sub txtAlmacen_DblClick()
    txtAlmacen_KeyDown vbKeyF2, 0
End Sub

Private Sub txtAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_almacen
                wcod_alm = vbNullString
                
                .Show 1
                
                If wcod_alm <> vbNullString Then
                    txtAlmacen.Text = wcod_alm
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtAlmacen_LostFocus()
    If Trim(txtAlmacen.Text) <> vbNullString Then
        lblAlmacen.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMALM", "EF2ALMACENES", "F2CODALM", Trim(txtAlmacen.Text), "T")
    Else
        txtAlmacen.Text = vbNullString
        lblAlmacen.Caption = vbNullString
    End If
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCodCentro_DblClick()
    txtCodCentro_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCodCentro_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_Centros
                wcodcosto = vbNullString
                
                .Show 1
                
                If wcodcosto <> vbNullString Then
                    txtCodCentro.Text = wcodcosto
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCodCentro_LostFocus()
    If Trim(txtCodCentro.Text) <> vbNullString Then
        lblCentroCosto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", Trim(txtCodCentro.Text), "T")
    Else
        txtCodCentro.Text = vbNullString
        lblCentroCosto.Caption = vbNullString
    End If
End Sub

Private Sub txtCodCentro_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCodFab_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtCodGasto_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtColor_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
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

Private Sub txtCtaContableImp_LostFocus()
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

Private Sub txtCtaContableImpVta_LostFocus()
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



Private Sub txtCtaContableInventa_DblClick()
    txtCtaContableInventa_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCtaContableInventa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_PlanCta
                wctacont = vbNullString
                
                .Show 1
                
                If wctacont <> vbNullString Then
                    txtCtaContableInventa.Text = wctacont
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCtaContableInventa_LostFocus()
    If Trim(txtCtaContableInventa.Text) <> vbNullString Then
        lblCtaContableInventa.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(txtCtaContableInventa.Text), "T")
    Else
        txtCtaContableInventa.Text = vbNullString
        lblCtaContableInventa.Caption = vbNullString
    End If
End Sub

Private Sub txtCtaContableInventa_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtCtaContableVenta_DblClick()
    txtCtaContableVenta_KeyDown vbKeyF2, 0
End Sub

Private Sub txtCtaContableVenta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With Ayuda_PlanCta
                wctacont = vbNullString
                
                .Show 1
                
                If wctacont <> vbNullString Then
                    txtCtaContableVenta.Text = wctacont
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtCtaContableVenta_LostFocus()
    If Trim(txtCtaContableVenta.Text) <> vbNullString Then
        lblCtaContableVenta.Caption = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", Trim(txtCtaContableVenta.Text), "T")
    Else
        txtCtaContableVenta.Text = vbNullString
        lblCtaContableVenta.Caption = vbNullString
    End If
End Sub

Private Sub txtCtaContableVenta_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtDescripcionFab_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtmarca_DblClick()
    txtmarca_KeyDown vbKeyF2, 0
End Sub

Private Sub txtMarcaModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub


Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtStockMaximo_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtStockMaximo_LostFocus()
    txtStockMaximo.Text = Format(txtStockMaximo.Text, "#,0.00")
End Sub

Private Sub txtStockMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtStockMinimo_LostFocus()
    txtStockMinimo.Text = Format(txtStockMinimo.Text, "#,0.00")
End Sub

Private Sub txtPorcentajeDemasia_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtPorcentajeDemasia_LostFocus()
    txtPorcentajeDemasia.Text = Format(txtPorcentajeDemasia.Text, "#,0.00")
End Sub

Private Sub txtStockReposicion_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtStockReposicion_LostFocus()
    txtStockReposicion.Text = Format(txtStockReposicion.Text, "#,0.00")
End Sub

Private Sub txtSubFamilia_DblClick()
    txtSubFamilia_KeyDown vbKeyF2, 0
End Sub

Private Sub txtSubFamilia_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With frmListaSubFamilia
                .Ayuda = True
                
                objAyudaSubFamilia.inicializarEntidades
                
                .Show 1
                
                If objAyudaSubFamilia.Codigo <> vbNullString Then
                    txtSubFamilia.Text = objAyudaSubFamilia.Codigo
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtSubFamilia_LostFocus()
    If Trim(txtSubFamilia.Text) <> vbNullString Then
        lblSubFamilia.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7DESCON", "SF7NIVEL02", "F7CODCON", Trim(txtSubFamilia.Text), "T")
    Else
        txtSubFamilia.Text = vbNullString
        lblSubFamilia.Caption = vbNullString
    End If
End Sub

Private Sub txtSubFamilia_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtTalla_KeyPress(KeyAscii As Integer)
    KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
End Sub


Private Sub txtUnidadMedida_DblClick()
    txtUnidadMedida_KeyDown vbKeyF2, 0
End Sub

Private Sub txtUnidadMedida_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_unidades
                wcodmed = vbNullString
                
                .Show 1
                
                If wcodmed <> vbNullString Then
                    txtUnidadMedida.Text = wcodmed
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtUnidadMedida_LostFocus()
    If Trim(txtUnidadMedida.Text) <> vbNullString Then
        lblUnidadMedida.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7NOMMED", "EF7MEDIDAS", "F7CODMED", Trim(txtUnidadMedida.Text), "T")
    Else
        txtUnidadMedida.Text = vbNullString
        lblUnidadMedida.Caption = vbNullString
    End If
End Sub

Private Sub txtUnidadMedida_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaAlfaNumerica(KeyAscii)
End Sub

Private Sub txtmarca_KeyPress(KeyAscii As Integer)
    KeyAscii = validarCajaNumerica(KeyAscii)
End Sub

Private Sub txtmarca_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_marcas
                wcodmar = vbNullString
                
                .Show 1
                
                If wcodmar <> vbNullString Then
                    txtMarca.Text = wcodmar
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtmarca_LostFocus()
    If Trim(txtMarca.Text) <> vbNullString Then
        lblMarca.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESMAR", "EF2MARCAS", "F2CODMAR", Trim(txtMarca.Text), "T")
    Else
        txtMarca.Text = vbNullString
        lblMarca.Caption = vbNullString
    End If
End Sub



