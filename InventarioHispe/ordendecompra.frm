VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{791923BA-56CB-4A36-9EA3-1B4ED74622AA}#1.0#0"; "csimxctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ordendecompra 
   Caption         =   "Orden de Compra"
   ClientHeight    =   9030
   ClientLeft      =   210
   ClientTop       =   1725
   ClientWidth     =   18735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ordendecompra.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   18735
   WindowState     =   2  'Maximized
   Begin InternetMailCtl.InternetMail InternetMail 
      Left            =   -120
      Top             =   600
      _cx             =   741
      _cy             =   741
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ImageList imgLstEstado 
      Left            =   0
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":058A
            Key             =   "Estado 1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":0B24
            Key             =   "Estado 2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":10BE
            Key             =   "Estado 3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":1658
            Key             =   "Estado 4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":1BF2
            Key             =   "Estado 5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":218C
            Key             =   "Estado 6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":2726
            Key             =   "Estado 7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ordendecompra.frx":2CC0
            Key             =   "Estado 8"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSeguimiento 
      Caption         =   " Seguimiento "
      ForeColor       =   &H00FF0000&
      Height          =   2145
      Left            =   11760
      TabIndex        =   69
      Top             =   1560
      Width           =   3195
      Begin VB.TextBox txtRecepcionadoPor 
         Height          =   315
         Left            =   720
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox chkOrdenRecepcionada 
         Caption         =   "Orden Recepcionada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1200
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpFechaEnvio 
         Height          =   300
         Left            =   1800
         TabIndex        =   74
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   138018817
         CurrentDate     =   41863
      End
      Begin VB.TextBox txtEnviadoPor 
         Height          =   315
         Left            =   720
         TabIndex        =   73
         Text            =   "Text2"
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkOrdenEnviada 
         Caption         =   "Orden Enviada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpFechaRecepcion 
         Height          =   300
         Left            =   1800
         TabIndex        =   79
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   138018817
         CurrentDate     =   41863
      End
      Begin VB.Label Label15 
         Caption         =   "Fec. de Recepción"
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Por"
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha de Envio"
         Height          =   255
         Left            =   360
         TabIndex        =   72
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Por"
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraNroOrden 
      Height          =   1425
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   2775
      Begin MSComctlLib.ImageCombo imgCmbEstado 
         Height          =   345
         Left            =   120
         TabIndex        =   93
         Top             =   960
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "imgCmbEstado"
         ImageList       =   "imgLstEstado"
      End
      Begin VB.TextBox Txt_NumOC 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   2085
      End
      Begin VB.TextBox Txt_TOC 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lblAnulada 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ANULADA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   885
         TabIndex        =   92
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nº Orden Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   2265
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1530
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   11160
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12000
         MaxLength       =   100
         TabIndex        =   63
         Text            =   "0.00"
         Top             =   240
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox txtempresa 
         Height          =   315
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   14
         Top             =   300
         Width           =   6270
      End
      Begin VB.TextBox txtobserva 
         Height          =   675
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9600
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   360
         TabIndex        =   58
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   50
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame FrameOC 
      Caption         =   " Datos del Proveedor "
      ForeColor       =   &H00FF0000&
      Height          =   1425
      Left            =   3000
      TabIndex        =   31
      Top             =   120
      Width           =   11955
      Begin VB.CheckBox chkSinProveedorEsp 
         Alignment       =   1  'Right Justify
         Caption         =   "Orden sin Proveedor Especifico."
         Height          =   210
         Left            =   5200
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtcontacto 
         Height          =   315
         Left            =   9360
         TabIndex        =   4
         Top             =   600
         Width           =   2460
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   315
         Left            =   9240
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.TextBox Txt_Prove 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   1185
      End
      Begin VB.ComboBox CmbTipDoc 
         Height          =   330
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   2475
      End
      Begin VB.TextBox txtusuario 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6120
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txt_NumSolComp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9240
         MaxLength       =   12
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin Threed.SSPanel pnldireprv 
         Height          =   270
         Left            =   1080
         TabIndex        =   83
         Top             =   960
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   476
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlnomprv 
         Height          =   315
         Left            =   2280
         TabIndex        =   84
         Top             =   600
         Width           =   5610
         _Version        =   65536
         _ExtentX        =   9895
         _ExtentY        =   556
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin MSComCtl2.DTPicker txt_fecha 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   138018817
         CurrentDate     =   40611
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   90
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   88
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   8040
         TabIndex        =   87
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   8040
         TabIndex        =   86
         Top             =   960
         Width           =   1155
      End
      Begin CONTROLSLibCtl.dxCheckBox ChK_regularizacion 
         Height          =   270
         Left            =   9000
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
         _Version        =   65536
         _cx             =   2461
         _cy             =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Regularización"
         Enabled         =   -1  'True
         AutoSize        =   -1  'True
         BackStyle       =   1
         BackColor       =   -2147483633
         ForeColor       =   0
         ViewStyle       =   1
         Checked         =   0   'False
         GroupIndex      =   -1
         TextLayout      =   0
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   5160
         TabIndex        =   33
         Top             =   1680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Requerimiento"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   7800
         TabIndex        =   32
         Top             =   1560
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdFirmaAprob 
      Caption         =   "Firmas de Aprobación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6735
      TabIndex        =   29
      Top             =   7920
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Frame Frame2 
      Caption         =   " Especificaciones de Orden "
      ForeColor       =   &H00FF0000&
      Height          =   2145
      Left            =   120
      TabIndex        =   34
      Top             =   1560
      Width           =   11595
      Begin VB.TextBox txtAutorizado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9960
         TabIndex        =   98
         Top             =   960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtFechaPago 
         Height          =   315
         Left            =   9960
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtlugar_entrega 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1750
         Width           =   7155
      End
      Begin VB.TextBox TxtCodCosto 
         Height          =   312
         Left            =   1440
         TabIndex        =   10
         Top             =   1380
         Width           =   1020
      End
      Begin VB.TextBox txtcodsoli 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txtcodforma 
         Height          =   312
         Left            =   1440
         TabIndex        =   9
         Top             =   1020
         Width           =   1020
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   315
         Left            =   9960
         TabIndex        =   19
         Top             =   1680
         Width           =   1515
      End
      Begin VB.TextBox txt_tc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9960
         TabIndex        =   17
         Text            =   "2.7"
         Top             =   600
         Width           =   1515
      End
      Begin VB.ComboBox Cmbmone 
         Height          =   330
         ItemData        =   "ordendecompra.frx":325A
         Left            =   9960
         List            =   "ordendecompra.frx":3267
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1515
      End
      Begin Threed.SSPanel pnlnomsoli 
         Height          =   300
         Left            =   2520
         TabIndex        =   35
         Top             =   660
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlnomforma 
         Height          =   300
         Left            =   2520
         TabIndex        =   36
         Top             =   1020
         Width           =   3372
         _Version        =   65536
         _ExtentX        =   5948
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel PnlNomCosto 
         Height          =   300
         Left            =   2520
         TabIndex        =   55
         Top             =   1380
         Width           =   6015
         _Version        =   65536
         _ExtentX        =   10610
         _ExtentY        =   529
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin MSComCtl2.DTPicker abofechaentrega 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   138018817
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker aBoHoraEntrega 
         Height          =   315
         Left            =   3120
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   138018818
         CurrentDate     =   40611
      End
      Begin VB.Label lblAutorizado 
         Caption         =   "Autorizado por"
         Height          =   255
         Left            =   8760
         TabIndex        =   99
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin CONTROLSLibCtl.dxCheckBox Chk_pagoparcial 
         Height          =   270
         Left            =   4680
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   2550
         _Version        =   65536
         _cx             =   4498
         _cy             =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Control de Pagos Parciales"
         Enabled         =   -1  'True
         AutoSize        =   -1  'True
         BackStyle       =   1
         BackColor       =   -2147483633
         ForeColor       =   16711680
         ViewStyle       =   1
         Checked         =   0   'False
         GroupIndex      =   -1
         TextLayout      =   0
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Pago"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   8760
         TabIndex        =   60
         Top             =   1320
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Entrega"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   1248
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Entrega"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Cotización"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   8760
         TabIndex        =   41
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   132
         TabIndex        =   40
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda "
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   19
         Left            =   8760
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cambio"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   17
         Left            =   8760
         TabIndex        =   37
         Top             =   600
         Width           =   1200
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   1365
      Left            =   18600
      OleObjectBlob   =   "ordendecompra.frx":3282
      TabIndex        =   57
      Top             =   3480
      Visible         =   0   'False
      Width           =   11715
   End
   Begin VB.Frame Frame3 
      Height          =   1530
      Left            =   11280
      TabIndex        =   42
      Top             =   7080
      Width           =   7335
      Begin VB.CheckBox chkDetraccionAplicar 
         Caption         =   "Evaluar aplicación % de Detracción."
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   97
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtDetraccionPorc 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   5520
         TabIndex        =   96
         Text            =   "0.00"
         Top             =   1100
         Width           =   1290
      End
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtigv 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   720
         Width           =   1290
      End
      Begin VB.TextBox txtmonto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox txtbase 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtRnd 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   720
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblmoneda 
         Alignment       =   2  'Center
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6360
         TabIndex        =   54
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblmoneda 
         Alignment       =   2  'Center
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5040
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblmoneda 
         Alignment       =   2  'Center
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   43
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label5 
         Caption         =   "Redondeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4200
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "B. Imponible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label10 
         Caption         =   "Monto Inaf."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1500
         TabIndex        =   48
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblImpuesto 
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         TabIndex        =   47
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5520
         TabIndex        =   46
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblmoneda 
         Alignment       =   2  'Center
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   45
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblmoneda 
         Alignment       =   2  'Center
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   44
         Top             =   480
         Width           =   360
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid Grid 
      Height          =   3285
      Left            =   120
      OleObjectBlob   =   "ordendecompra.frx":A4A0
      TabIndex        =   12
      Top             =   3720
      Width           =   18495
   End
   Begin ActiveToolBars.SSActiveToolBars atbmenu 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   11
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "ordendecompra.frx":17D39
      ToolBars        =   "ordendecompra.frx":2087B
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox9 
      Height          =   270
      Left            =   15000
      TabIndex        =   95
      Top             =   3000
      Width           =   2460
      _Version        =   65536
      _cx             =   4339
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cierre de Orden por Item."
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   12582912
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox8 
      Height          =   270
      Left            =   15000
      TabIndex        =   94
      Top             =   2640
      Width           =   2910
      _Version        =   65536
      _cx             =   5133
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visualizar Cliente de Requerimiento."
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox7 
      Height          =   270
      Left            =   15000
      TabIndex        =   91
      Top             =   2280
      Width           =   2535
      _Version        =   65536
      _cx             =   4471
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visualizar Descripción Interna."
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox6 
      Height          =   270
      Left            =   15000
      TabIndex        =   81
      Top             =   1920
      Width           =   2865
      _Version        =   65536
      _cx             =   5054
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visualizar Observaciones por Item."
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox5 
      Height          =   270
      Left            =   15000
      TabIndex        =   80
      Top             =   1560
      Width           =   2625
      _Version        =   65536
      _cx             =   4630
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visualizar Descuentos por Item."
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox4 
      Height          =   270
      Left            =   15000
      TabIndex        =   68
      Top             =   1200
      Width           =   3405
      _Version        =   65536
      _cx             =   6006
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visualizar Porcentaje de Demasia por Item."
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox3 
      Height          =   270
      Left            =   15000
      TabIndex        =   67
      Top             =   840
      Width           =   2565
      _Version        =   65536
      _cx             =   4524
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Visualizar B. Imponible por Item"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox2 
      Height          =   270
      Left            =   15000
      TabIndex        =   21
      Top             =   480
      Width           =   3045
      _Version        =   65536
      _cx             =   5371
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Activar 6 decimales en precio unitario"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox1 
      Height          =   270
      Left            =   15000
      TabIndex        =   20
      Top             =   120
      Width           =   3390
      _Version        =   65536
      _cx             =   5980
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Activar columnas de Gasto y Cta Contable"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin VB.Label lbldescripcion 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      Height          =   510
      Left            =   9600
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.Menu mnuitems 
      Caption         =   "&Item"
      Visible         =   0   'False
      Begin VB.Menu opcdel 
         Caption         =   "&Eliminar item"
      End
      Begin VB.Menu opcinsert 
         Caption         =   "&Insertar"
      End
   End
End
Attribute VB_Name = "ordendecompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Af As New ADOFunctions
Dim StrCn As String
Dim CnTmp As New ADODB.Connection
Dim editCc As String
Dim editBcc As String
Dim editFrom As String
Dim editTo As String
Dim editSubject As String
Dim editMessageText As String
'************************
Dim rsOrdenCab              As ADODB.Recordset
Dim rsOrdenDet              As ADODB.Recordset
Dim rssolcab                As ADODB.Recordset
Dim rsSolDet                As ADODB.Recordset
Dim rst                     As ADODB.Recordset
Dim rstaux                  As ADODB.Recordset
Dim rsproductos             As ADODB.Recordset
Dim SWcondipago             As Integer
Dim Wnuevo                  As Boolean
Dim flawigv                 As Boolean
Dim seleccion               As Boolean
Dim CadSql                  As String
Dim cnn_form                As New ADODB.Connection
Dim cconex_form             As String
Dim sw_nuevo_item           As Boolean
Dim ExisteOrdenCompra       As Boolean
Dim wIgv                    As Single
Dim swGrabacion             As Boolean
Dim inicio                  As Boolean
Dim sw_activate             As Boolean
Dim sw_cabecera             As Boolean
Dim sw_detalle              As Boolean
Dim sw_ayuda                As Boolean
Dim wgraba                  As Integer
Dim FlagGeneraOC            As Boolean
Dim jc                      As Integer
Dim flagwin                 As Boolean
Dim FlagAcceso              As Boolean
Dim whelp_solicitud         As Boolean
Dim xnombre                 As String
Dim flag                    As Boolean
Dim wgrabar                 As Boolean
Dim Fecha                   As String
Dim existe As Boolean
Dim SwRenovar               As Boolean
Dim wNumOc                  As String
Dim CtaPuntos               As Integer, i As Integer
'Private cImgInfo            As cImageInfo

Rem SK ADD:
Private bolAyuda            As Boolean
Private strTipoOrden        As String
Private strNumeroOrden      As String

Private strFichero          As String

Private objOrden            As ClsOrden



Rem Variables para Controlar la Devolucion de Foco del Registro en Grilla señalado antes de alguna Modificacion o Uso
Dim d As Double
Dim nSaveRecNo As Double


Public Property Let Ayuda(ByVal value As Boolean)
    bolAyuda = value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property
'Propiedad Tipo de Orden
Public Property Let TipoOrden(ByVal value As String)
    strTipoOrden = value
End Property

Public Property Get TipoOrden() As String
    TipoOrden = strTipoOrden
End Property

'Propiedad Numero de Orden
Public Property Let NumeroOrden(ByVal value As String)
    strNumeroOrden = value
End Property

Public Property Get NumeroOrden() As String
    NumeroOrden = strNumeroOrden
End Property

Private Sub Imprime_Orden(Opcion)
    On Error GoTo CapturaError
    
    Dim sql As String
    Dim RSCONSULTA As New ADODB.Recordset
    Set RSCONSULTA = New ADODB.Recordset
    Dim RsPago As New ADODB.Recordset
    Set RsPago = New ADODB.Recordset
    Dim RsCTR_COM As New ADODB.Recordset
    Set RsCTR_COM = New ADODB.Recordset
    Dim nAnchoHoja As Double
    
    If Opcion = 1 Then
        Dim rpt As New Acr_OrdenCompra
        
        With rpt 'Acr_OrdenCompra
            Set cImgInfo = New cImageInfo
       ' MsgBox "Acr_OrdenC_Otros"
            If Cmbmone.ListIndex = 0 Then
                .LblTotF.Caption = "Total " & "S/" '"S/"
            Else
                .LblTotF.Caption = "Total " & "US$"
            End If
            
            
            
            
            .flddirec1.Text = wf1direc1
            .FldTelf.Text = "Teléfono: " & wtelefono & " // Fax: " & wfax
            .LblCentroCosto.Text = PnlNomCosto.Caption
            '.flddirec2.Text = wf1direc2
            nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
            .fldruc.Text = "R.U.C. " & wrucempresa
            If FileExist(App.Path & "\Logo" & left(wempresa, 5) & ".bmp") = True Then 'wrucempresa & ".jpg") = True Then
'                .fldempresa.Visible = False
            With cImgInfo
                    .ReadImageInfo App.Path & "\Logo" & left(wempresa, 5) & ".bmp"
                End With
            Else
'                .fldempresa.Visible = True
'                .ImageLogo.Visible = False
'                .fldempresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
            End If
            
            '.IGV.Caption = wigv
            strNumeroOrden = Txt_NumOC.Text
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            sql = "SELECT A.F4NUMORD,A.F4NUMCOTIZA,A.F4ESTNUL, A.F4CODSOLICITUD,A.F4TIPDOC, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO, A.F4REGULARIZA,A.F4DIAPAGO,  B.F2TELPROV,  B.F2FAXPROV, " & _
                  "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4RND,A.F4OBSERVA, " & _
                  "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4ESTADO,A.F4FECENT,A.F4OBSERVA,A.F4CODPRV,A.F4TIPMON,A.F4REFERE,A.F4TIPCAM,A.F4FECGRA,A.F4USEGRA,A.F4FECMOD,A.F4USEMOD " & _
                  " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & strNumeroOrden & _
                  "' AND A.F4LOCAL='" & strTipoOrden & "' ORDER BY A.F4NUMORD DESC"
        
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
                
                '.F4NUMORD.Text = ("" & rsconsulta.Fields("F4NUMORD"))
                If left(strNumeroOrden, 2) = "OC" Then
                    .LblTitle.Caption = "ORDEN DE COMPRA"
                Else
                    .LblTitle.Caption = "ORDEN DE SERVICIO"
'                    .Field334.Text = "Formato LOG-F-08"
                End If
                .LblNroOC.Caption = "N° " & RSCONSULTA.Fields("F4NUMORD")
                '.fldsolicitud.Text = "" & RSCONSULTA.Fields("F4CODSOLICITUD")
                 
                'If Trim(RSCONSULTA.Fields("F4CODSOLICITUD") & "") <> vbNullString Then
                    objAyudaOrden.TipoOrden = strTipoOrden
                    objAyudaOrden.NumeroOrden = strNumeroOrden
                    
                    .fldsolicitud.Text = objAyudaOrden.generarCadenaSolicitud()
                If Val(RSCONSULTA!F4ESTADO & "") = 1 Then
                    '.LblEstado.Visible = True
                Else
                    .LblEstado.Visible = False
                End If
                
    '            If Val(RSCONSULTA!F4REGULARIZA & "") = 1 Then
    '                .LBLREGULARIZA.Visible = True
    '            Else
    '                .LBLREGULARIZA.Visible = False
    '            End If
                
                .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
                .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
                .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
                .F2CONTACTO.Visible = True
                .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
                .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
                .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
                .FldFchEntrega = "" & Format(RSCONSULTA.Fields("F4FECent"), "dd/mm/yyyy")
                .FldTipCam.Text = Format(Val("" & RSCONSULTA.Fields("F4tipcam")), "0.000")
                .FldTipDoc.Text = UCase("" & ObtenerCampo("DOCUMENTOS", "F2DESDOC", "F2CODDOC", RSCONSULTA!F4TIPDOC & "", "T", cnn_dbbancos))
                If RSCONSULTA!F4TIPDOC & "" = "02" Then
                   .LblImp.Caption = "Reten."
                Else
                    .LblImp.Caption = "I.G.V."
                End If
    '            .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "###,###,###,##0.00")
    '            '.Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                .FldObservaAll.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
                .FldSon.Text = CADENANUM(Val("" & RSCONSULTA.Fields("F4MONTO")), "" & RSCONSULTA.Fields("F4TIPMON"), "")
                '.referencia.Text = "" & Txt_Referencia.Text
                '.solicitado.Text = "" & pnlnomsoli.Caption
                If RsPago.State = adStateOpen Then RsPago.Close
                RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsPago.EOF Then
                    .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
                End If
                RsPago.Close
                .F4COTIZACION.Text = "" & RSCONSULTA.Fields("F4NUMCOTIZA")
                .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
                If Rs.State = 1 Then Rs.Close
                
                Rs.Open "Select F2NOMUSER From EF2USERS where F2CODUSER = '" & wusuario & "'", cnn_dbbancos, _
                adOpenKeyset, adLockOptimistic
                
            End If
            .DataControl1.ConnectionString = cnn_dbbancos
            
            CadSql = vbNullString
            CadSql = CadSql & "SELECT "
            CadSql = CadSql & "DET.F3CODFAB AS F3CODPRO, "
            CadSql = CadSql & "DET.F5NOMPRO, "
            CadSql = CadSql & "COLOR.DESCRIPCION AS DESCOLOR, "
            CadSql = CadSql & "MED.F7SIGMED AS F3MEDIDA, "
            CadSql = CadSql & "SUM(DET.F3CANPRO * (1 + (F3PORCDEMASIA/100))) AS F3CANPRO, "
            CadSql = CadSql & "DET.F3PRENETO AS F3PREUNI, "
            CadSql = CadSql & "SUM(DET.F5VALVTA) AS F5VALVTA, "
            CadSql = CadSql & "SUM(DET.F3IGV) AS F3IGV, "
            CadSql = CadSql & "SUM(DET.F3PORDCT) AS F3PORDCT, "
            CadSql = CadSql & "SUM(DET.F3TOTAL) AS F3TOTAL "
            CadSql = CadSql & "FROM "
            CadSql = CadSql & "((IF3ORDEN AS DET "
            CadSql = CadSql & "LEFT JOIN CENTROS AS CC ON CC.F3COSTO = DET.F3CENCOS) "
            CadSql = CadSql & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
            CadSql = CadSql & "LEFT JOIN EF2BIENCOLOR AS COLOR ON COLOR.CODIGO = DET.CODCOLOR "
            CadSql = CadSql & "WHERE "
            CadSql = CadSql & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
            CadSql = CadSql & "DET.F4LOCAL = '" & strTipoOrden & "' "
            CadSql = CadSql & "GROUP BY "
            CadSql = CadSql & "DET.F3CODFAB,  "
            CadSql = CadSql & "DET.F5NOMPRO, "
            CadSql = CadSql & "COLOR.DESCRIPCION, "
            CadSql = CadSql & "MED.F7SIGMED, "
            CadSql = CadSql & "DET.F3PRENETO"
            .DataControl1.Source = CadSql
            .Caption = "ORDEN DE COMPRA NACIONAL"
            RSCONSULTA.Close
            
            .DescargarReporte = False
            
            .Show '1
        End With
    ElseIf Opcion = 2 Then
        With Acr_OrdenCompra
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            
            sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
                    "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA, " & _
                    "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT,A.CODPRV,A.F4TIPMON " & _
                    " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & strNumeroOrden & _
                    " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"
        
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            
            If Not RSCONSULTA.EOF Then
                '.F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
                .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
                .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").value
                .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
                .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
                .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
                .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
                .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
                '.F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
                .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
                If RSCONSULTA.Fields("F4TIPMON") = "S" Then
                  '.TIPMON.Text = "SOLES"
                  Else
                  If RSCONSULTA.Fields("F4TIPMON") = "D" Then
                  '.TIPMON.Text = "DOLARES"
                  End If
                End If
        '
                If RsPago.State = adStateOpen Then RsPago.Close
                RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsPago.EOF Then
                    .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
                End If
                RsPago.Close
        '
                If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
                RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
                If Not RsCTR_COM.EOF Then
                    '.F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
                    '.F4NOTA.Text = "" & RsCTR_COM.Fields("F1NOTA_OC")
                    '.F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                    '.F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
                End If
                RsCTR_COM.Close
                .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
        
            End If
            
            RSCONSULTA.Close
            
            .DataControl1.ConnectionString = cnn_form
            .DataControl1.Source = "select * from TMPORDENDECOMPRA"
            
            .Show 'vbModal
        End With
    End If
    
    Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Logistica"
    
    Resume Next
End Sub

Private Sub imprimeOrdenV2(ByVal strTipoOrden As String, _
                            ByVal strNumeroOrden As String, _
                            Optional ByVal imprimirPDFparaEnvioMail As Boolean)
    
    On Error GoTo errImprimeOrdenV2
    
    Dim nAnchoHoja As Double, strTextoObs As String, strTextoObs2 As String
    
    With objAyudaOrden
        .TipoOrden = strTipoOrden
        .NumeroOrden = strNumeroOrden
        
        .obtenerConfigOrden
    End With
    
    If strTipoOrden <> "OF" Then
    
        Dim rpt As New Acr_OrdenCompra
        
        With rpt
            .DescargarReporte = imprimirPDFparaEnvioMail
            
            If imprimirPDFparaEnvioMail Then
                If Dir(wrutatemp & "\ParaAtencionDeOrden.pdf", vbArchive) <> vbNullString Then
                    Kill wrutatemp & "\ParaAtencionDeOrden.pdf"
                End If
                
                .TipoOrden = strTipoOrden
                .NumeroOrden = strNumeroOrden
            End If
            
'            Set cImgInfo = New cImageInfo
            
            Select Case objAyudaOrden.CodMoneda
                Case "S"
                    .LblTotF.Caption = "Total " & "S/"
                Case "D"
                    .LblTotF.Caption = "Total " & "US$"
                Case "E"
                    .LblTotF.Caption = "Total " & "€"
            End Select
            
            .flddirec1.Text = wf1direc1
            .FldTelf.Text = "Teléfono: " & wtelefono & " // Fax: " & wfax
            .LblCentroCosto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", objAyudaOrden.CentroCosto, "T")    'PnlNomCosto.Caption
            
            nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
            .fldruc.Text = "R.U.C. " & wrucempresa
            
    
            Select Case objAyudaOrden.TipoOrden
                Case "OC"
                    .LblTitle.Caption = "ORDEN DE COMPRA"
                Case "OS"
                    .LblTitle.Caption = "ORDEN DE SERVICIO"
'                .Field334.Text = "Formato LOG-F-08"
            End Select
            
            strTextoObs = "INSTRUCCIONES:  "
            strTextoObs = strTextoObs & Chr(13) & "1.- ENTREGA DE MATERIAL Y/O BIEN"
            strTextoObs = strTextoObs & Chr(13) & "1.1 Presentar orden de compra en estado “APROBADO” debidamente firmada."
            strTextoObs = strTextoObs & Chr(13) & "1.2 Embalaje adecuado y respectivo certificado de calidad si lo amerita."
            strTextoObs = strTextoObs & Chr(13) & "1.3 Adjuntar:"
            strTextoObs = strTextoObs & Chr(13) & "    * Cotización del proveedor."
            strTextoObs = strTextoObs & Chr(13) & "    * Guía de remisión (original) y dos fotocopias."
            strTextoObs = strTextoObs & Chr(13) & "    * Copia de la factura."
            strTextoObs = strTextoObs & Chr(13) & "2.- PARA LA PRESENTACIÓN DE COMPROBANTE DE PAGO"
            strTextoObs = strTextoObs & Chr(13) & "2.1 Adjuntar:"
            strTextoObs = strTextoObs & Chr(13) & "    * Factura comercial original"
            strTextoObs = strTextoObs & Chr(13) & "    * Guía de remisión sellada por almacén."
            strTextoObs = strTextoObs & Chr(13) & "    * Orden de compra debidamente aprobada."
            strTextoObs = strTextoObs & Chr(13) & "    * Cotización del proveedor."
            strTextoObs = strTextoObs & Chr(13) & "    * Acta de conformidad (si la Orden de Compra corresponde a un servicio)."
            strTextoObs = strTextoObs & Chr(13) & "    * En caso de Comprobante Electrónico, enviar dos fotocopias, archivos en formato PDF y xml al correo contabilidad@ hispe.com.pe"
            strTextoObs = strTextoObs & Chr(13) & "2.2 Indicar el número de orden de compra."
            strTextoObs = strTextoObs & Chr(13) & "2.3 Descripción detallada del bien vendido, o tipo de servicio prestado indicando periodo, cantidad yunidad de medida."
            strTextoObs = strTextoObs & Chr(13) & "2.4 Indicar en el Comprobante de Pago si es agente de retención o buen contribuyente."
            strTextoObs = strTextoObs & Chr(13) & "2.5 De ser un servicio colocar tasa, código detracción y cuenta de detracciones."
            strTextoObs = strTextoObs & Chr(13) & "2.6 Horario y lugar de recepción de Comprobantes de Pago: Martes: 8:30am – 1pm / 2pm a 5:00pm Lt. 12A Mz E Asociacion Sumac Pacha - Lurin"
            strTextoObs2 = "3.- PARA EL PAGO"
            strTextoObs2 = strTextoObs2 & Chr(13) & "3.1 La condición rige desde la fecha de recepción del Comprobante de Pago."
            strTextoObs2 = strTextoObs2 & Chr(13) & "3.2 Atención de Tesorería: Jueves: 3:00pm - 5:30pm contabilidad@ hispe.com.pe"
            
            strTextoObs2 = strTextoObs2 & Chr(13) & "4.- CONDICIONES Y PENALIDADES"
            strTextoObs2 = strTextoObs2 & Chr(13) & "4.1 Nos reservamos el derecho de devolver la mercadería que no cumpla los estándares de calidad y/o no haya sido entregada en la fecha pactada."
            strTextoObs2 = strTextoObs2 & Chr(13) & "4.2 La factura debe estar conforme con lo descrito en la orden de compra; en caso contrario se procederá a la devolución de la misma."
    
'            .Label130.Caption = strTextoObs
'            .Label131.Caption = strTextoObs2
    
            .LblNroOC.Caption = "N° " & objAyudaOrden.NumeroOrden
            
            .fldsolicitud.Text = objAyudaOrden.generarCadenaSolicitud
            
            With objAyudaProveedor
                .Codigo = objAyudaOrden.CodProveedor
                
                .obtenerProveedor
            End With
            
            .F2NOMPROV.Text = objAyudaProveedor.NombreProveedor
            .F2DIRPROV.Text = objAyudaProveedor.DireccionProveedor
            .F2CONTACTO.Text = objAyudaProveedor.Contacto
            .F2CONTACTO.Visible = True
            .F2TELPROV.Text = objAyudaProveedor.Telefono
            .F2FAXPROV.Text = objAyudaProveedor.Fax
            
            .F4FECEMI.Text = objAyudaOrden.FechaEmision
            .FldFchEntrega = objAyudaOrden.FechaEntrega
            .FldTipCam.Text = Format(objAyudaOrden.TipoCambio, "0.000")
'            .fldcuentaabono.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NUMCTA", "EF2PROVEEDORES", "F2CODPROV", objAyudaOrden.CodProveedor, "T")  'ObtenerCampo("ef2users", "f2nomuser", "f2coduser", RSCONSULTA.Fields("F4CODSOL"), "T", cnn_dbbancos)
            
            .FldTipDoc.Text = UCase(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESDOC", "DOCUMENTOS", "F2CODDOC", objAyudaOrden.CodTipoComprobante, "T")) 'UCase("" & ObtenerCampo("DOCUMENTOS", "F2DESDOC", "F2CODDOC", RSCONSULTA!F4TIPDOC & "", "T", cnn_dbbancos))
            
            Select Case objAyudaOrden.CodTipoComprobante
                Case "02"
                    .LblImp.Caption = "Reten."
                Case Else
                    .LblImp.Caption = "I.G.V."
            End Select
            
            .FldObservaAll.Text = objAyudaOrden.Observacion
            .F4CODPRV.Text = objAyudaOrden.RucProveedor
            .FldSon.Text = CADENANUM(objAyudaOrden.TotalFacturado, objAyudaOrden.CodMoneda, vbNullString)
            
            .F2DESPAG.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", objAyudaOrden.CodFormaPago, "T")
                            
            .F4COTIZACION.Text = objAyudaOrden.NumeroCotizacion
            .REMITIR.Text = objAyudaOrden.LugarEntrega
            
            .DataControl1.ConnectionString = cnn_dbbancos
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "DET.ITEM, "
            SqlCad = SqlCad & "DET.F3CODFAB AS F3CODPRO, "
            SqlCad = SqlCad & "DET.F5NOMPRO, "
            SqlCad = SqlCad & "COLOR.DESCRIPCION AS DESCOLOR, "
            SqlCad = SqlCad & "MED.F7SIGMED AS F3MEDIDA, "
            SqlCad = SqlCad & "SUM(DET.F3CANPRO * (1 + (F3PORCDEMASIA/100))) AS F3CANPRO, "
            SqlCad = SqlCad & "VAL(FORMAT(DET.F3PRENETO, '#0.00')) AS F3PREUNI, "
            SqlCad = SqlCad & "SUM(DET.F5VALVTA) AS F5VALVTA, "
            SqlCad = SqlCad & "SUM(DET.F3IGV) AS F3IGV, "
            SqlCad = SqlCad & "SUM(DET.F3PORDCT) AS F3PORDCT, "
            SqlCad = SqlCad & "SUM(DET.F3TOTAL) AS F3TOTAL "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "((IF3ORDEN AS DET "
            SqlCad = SqlCad & "LEFT JOIN CENTROS AS CC ON CC.F3COSTO = DET.F3CENCOS) "
            SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
            SqlCad = SqlCad & "LEFT JOIN EF2BIENCOLOR AS COLOR ON COLOR.CODIGO = DET.CODCOLOR "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
            SqlCad = SqlCad & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
            SqlCad = SqlCad & "DET.F3CANPRO <> 0 "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "DET.ITEM,  "
            SqlCad = SqlCad & "DET.F3CODFAB,  "
            SqlCad = SqlCad & "DET.F5NOMPRO, "
            SqlCad = SqlCad & "COLOR.DESCRIPCION, "
            SqlCad = SqlCad & "MED.F7SIGMED, "
            SqlCad = SqlCad & "VAL(FORMAT(DET.F3PRENETO, '#0.00'))"
            
            .DataControl1.Source = SqlCad
            
            .Caption = "ORDEN DE COMPRA NACIONAL"
            
            .Show
        
        End With
    
    Else
        
        With Acr_Oferta
            .DescargarReporte = imprimirPDFparaEnvioMail
            
            If imprimirPDFparaEnvioMail Then
                If Dir(wrutatemp & "\ParaAtencionDeOrden.pdf", vbArchive) <> vbNullString Then
                    Kill wrutatemp & "\ParaAtencionDeOrden.pdf"
                End If
                
                .TipoOrden = strTipoOrden
                .NumeroOrden = strNumeroOrden
            End If
            
'            Set cImgInfo = New cImageInfo
            
            Select Case objAyudaOrden.CodMoneda
                Case "S"
                    .LblTotF.Caption = "Total " & "S/"
                Case "D"
                    .LblTotF.Caption = "Total " & "US$"
                Case "E"
                    .LblTotF.Caption = "Total " & "€"
            End Select
            
            
            nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
            .LblTitle.Caption = "OFERTA"
            
            strTextoObs = "INSTRUCCIONES:  "
            strTextoObs = strTextoObs & Chr(13) & "1.- ENTREGA DE MATERIAL Y/O BIEN"
            strTextoObs = strTextoObs & Chr(13) & "1.1 Presentar orden de compra en estado “APROBADO” debidamente firmada."
            strTextoObs = strTextoObs & Chr(13) & "1.2 Embalaje adecuado y respectivo certificado de calidad si lo amerita."
            strTextoObs = strTextoObs & Chr(13) & "1.3 Adjuntar:"
            strTextoObs = strTextoObs & Chr(13) & "    * Cotización del proveedor."
            strTextoObs = strTextoObs & Chr(13) & "    * Guía de remisión (original) y dos fotocopias."
            strTextoObs = strTextoObs & Chr(13) & "    * Copia de la factura."
            strTextoObs = strTextoObs & Chr(13) & "2.- PARA LA PRESENTACIÓN DE COMPROBANTE DE PAGO"
            strTextoObs = strTextoObs & Chr(13) & "2.1 Adjuntar:"
            strTextoObs = strTextoObs & Chr(13) & "    * Factura comercial original"
            strTextoObs = strTextoObs & Chr(13) & "    * Guía de remisión sellada por almacén."
            strTextoObs = strTextoObs & Chr(13) & "    * Orden de compra debidamente aprobada."
            strTextoObs = strTextoObs & Chr(13) & "    * Cotización del proveedor."
            strTextoObs = strTextoObs & Chr(13) & "    * Acta de conformidad (si la Orden de Compra corresponde a un servicio)."
            strTextoObs = strTextoObs & Chr(13) & "    * En caso de Comprobante Electrónico, enviar dos fotocopias, archivos en formato PDF y xml al correo contabilidad@ hispe.com.pe"
            strTextoObs = strTextoObs & Chr(13) & "2.2 Indicar el número de orden de compra."
            strTextoObs = strTextoObs & Chr(13) & "2.3 Descripción detallada del bien vendido, o tipo de servicio prestado indicando periodo, cantidad yunidad de medida."
            strTextoObs = strTextoObs & Chr(13) & "2.4 Indicar en el Comprobante de Pago si es agente de retención o buen contribuyente."
            strTextoObs = strTextoObs & Chr(13) & "2.5 De ser un servicio colocar tasa, código detracción y cuenta de detracciones."
            strTextoObs = strTextoObs & Chr(13) & "2.6 Horario y lugar de recepción de Comprobantes de Pago: Martes: 8:30am – 1pm / 2pm a 5:00pm Lt. 12A Mz E Asociacion Sumac Pacha - Lurin"
            strTextoObs2 = "3.- PARA EL PAGO"
            strTextoObs2 = strTextoObs2 & Chr(13) & "3.1 La condición rige desde la fecha de recepción del Comprobante de Pago."
            strTextoObs2 = strTextoObs2 & Chr(13) & "3.2 Atención de Tesorería: Jueves: 3:00pm - 5:30pm contabilidad@ hispe.com.pe"
            
            strTextoObs2 = strTextoObs2 & Chr(13) & "4.- CONDICIONES Y PENALIDADES"
            strTextoObs2 = strTextoObs2 & Chr(13) & "4.1 Nos reservamos el derecho de devolver la mercadería que no cumpla los estándares de calidad y/o no haya sido entregada en la fecha pactada."
            strTextoObs2 = strTextoObs2 & Chr(13) & "4.2 La factura debe estar conforme con lo descrito en la orden de compra; en caso contrario se procederá a la devolución de la misma."
    
            .Label130.Caption = strTextoObs
            .Label131.Caption = strTextoObs2
    
            .LblNroOC.Caption = "N° " & objAyudaOrden.NumeroOrden
            
            
            With objAyudaProveedor
                .Codigo = objAyudaOrden.CodProveedor
                
                .obtenerProveedor
            End With
            
            .F2NOMPROV.Text = objAyudaProveedor.NombreProveedor
            .F2DIRPROV.Text = objAyudaProveedor.DireccionProveedor
            .F2CONTACTO.Text = objAyudaProveedor.Contacto
            .F2CONTACTO.Visible = True
            
            .F4FECEMI.Text = objAyudaOrden.FechaEmision
            .F4FECEMI.Text = Format(objAyudaOrden.FechaEmision, "dd") & " de " & dev_mes(Month(objAyudaOrden.FechaEmision)) & " de " & Format(objAyudaOrden.FechaEmision, "yyyy")

            
            
            Select Case objAyudaOrden.CodTipoComprobante
                Case "02"
                    .LblImp.Caption = "Reten."
                Case Else
                    .LblImp.Caption = "I.G.V."
            End Select
            
            .FldObservaAll.Text = objAyudaOrden.Observacion
            .FldSon.Text = CADENANUM(objAyudaOrden.TotalFacturado, objAyudaOrden.CodMoneda, vbNullString)
            
                            
            .DataControl1.ConnectionString = cnn_dbbancos
            
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "DET.ITEM, "
            SqlCad = SqlCad & "DET.F3CODFAB AS F3CODPRO, "
            SqlCad = SqlCad & "DET.F5NOMPRO, "
            SqlCad = SqlCad & "COLOR.DESCRIPCION AS DESCOLOR, "
            SqlCad = SqlCad & "MED.F7SIGMED AS F3MEDIDA, "
            SqlCad = SqlCad & "SUM(DET.F3CANPRO * (1 + (F3PORCDEMASIA/100))) AS F3CANPRO, "
            SqlCad = SqlCad & "VAL(FORMAT(DET.F3PRENETO, '#0.00')) AS F3PREUNI, "
            SqlCad = SqlCad & "SUM(DET.F5VALVTA) AS F5VALVTA, "
            SqlCad = SqlCad & "SUM(DET.F3IGV) AS F3IGV, "
            SqlCad = SqlCad & "SUM(DET.F3PORDCT) AS F3PORDCT, "
            SqlCad = SqlCad & "SUM(DET.F3TOTAL) AS F3TOTAL "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "((IF3ORDEN AS DET "
            SqlCad = SqlCad & "LEFT JOIN CENTROS AS CC ON CC.F3COSTO = DET.F3CENCOS) "
            SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
            SqlCad = SqlCad & "LEFT JOIN EF2BIENCOLOR AS COLOR ON COLOR.CODIGO = DET.CODCOLOR "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
            SqlCad = SqlCad & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
            SqlCad = SqlCad & "DET.F3CANPRO <> 0 "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "DET.ITEM,  "
            SqlCad = SqlCad & "DET.F3CODFAB,  "
            SqlCad = SqlCad & "DET.F5NOMPRO, "
            SqlCad = SqlCad & "COLOR.DESCRIPCION, "
            SqlCad = SqlCad & "MED.F7SIGMED, "
            SqlCad = SqlCad & "VAL(FORMAT(DET.F3PRENETO, '#0.00'))"
            
            .DataControl1.Source = SqlCad
            
            .Caption = "OFERTA"
            
            .Show
        
        End With
        
    End If
    
    Exit Sub
    Resume
errImprimeOrdenV2:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub Imprime_Orden_ParaEnvioMail(Opcion)
    On Error GoTo CapturaError
    
    Dim sql As String
    Dim RSCONSULTA As New ADODB.Recordset
    Set RSCONSULTA = New ADODB.Recordset
    Dim RsPago As New ADODB.Recordset
    Set RsPago = New ADODB.Recordset
    Dim RsCTR_COM As New ADODB.Recordset
    Set RsCTR_COM = New ADODB.Recordset
    Dim nAnchoHoja As Double
    
    If Opcion = 1 Then
        If Dir(wrutatemp & "\ParaAtencionDeOrden.pdf", vbArchive) <> vbNullString Then
            Kill wrutatemp & "\ParaAtencionDeOrden.pdf"
        End If
        
        Dim rpt As New Acr_OrdenCompra
        
        With rpt 'Acr_OrdenCompra
            Set cImgInfo = New cImageInfo
       ' MsgBox "Acr_OrdenC_Otros"
            If Cmbmone.ListIndex = 0 Then
                .LblTotF.Caption = "Total " & "S/" '"S/"
            Else
                .LblTotF.Caption = "Total " & "US$"
    
            End If
            .flddirec1.Text = wf1direc1
            .FldTelf.Text = "Teléfono: " & wtelefono
            .LblCentroCosto.Text = PnlNomCosto.Caption
            '.flddirec2.Text = wf1direc2
            nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
            .fldruc.Text = "R.U.C. " & wrucempresa
            If FileExist(App.Path & "\Logo" & left(wempresa, 5) & ".bmp") = True Then 'wrucempresa & ".jpg") = True Then
'                .fldempresa.Visible = False
            With cImgInfo
                    .ReadImageInfo App.Path & "\Logo" & left(wempresa, 5) & ".bmp"
                End With
            Else
'                .fldempresa.Visible = True
'                .fldempresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
            End If
            
            '.IGV.Caption = wigv
            strNumeroOrden = Txt_NumOC.Text
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            sql = "SELECT A.F4NUMORD,A.F4NUMCOTIZA,A.F4ESTNUL, A.F4CODSOLICITUD,A.F4TIPDOC, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO, A.F4REGULARIZA,A.F4DIAPAGO,  B.F2TELPROV,  B.F2FAXPROV, " & _
                  "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4RND,A.F4OBSERVA, " & _
                  "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4ESTADO,A.F4FECENT,A.F4OBSERVA,A.F4CODPRV,A.F4TIPMON,A.F4REFERE,A.F4TIPCAM,A.F4FECGRA,A.F4USEGRA,A.F4FECMOD,A.F4USEMOD " & _
                  " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & strNumeroOrden & _
                  "' AND A.F4LOCAL='" & strTipoOrden & "' ORDER BY A.F4NUMORD DESC"
        
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
                
                '.F4NUMORD.Text = ("" & rsconsulta.Fields("F4NUMORD"))
                If left(strNumeroOrden, 2) = "OC" Then
                    .LblTitle.Caption = "ORDEN DE COMPRA"
'                    .Label120.Caption = "-Recepción de mercaderia y documentación sera en la Oficina Central de Lunes a Viernes en hrs. 09:00 a 12:00 y 14:00 – 18:00."
'                    .Label121.Caption = "-Se debe adjuntar Guía de Remisión Original (Destinatario y Sunat) con fecha y copia de Orden de Compra."
'                    .Label122.Caption = "-En caso de estar afecto a detracción, colocar el sello legible con el número de cuenta respectivo."
'                    .Label123.Caption = "-La fecha de vencimiento de las facturas, se consideran a partir de la fecha de recepción y conformidad del documento."
'                    .Label124.Caption = "-El pago de facturas se realizara el día Viernes de la semana correspondiente a la fecha de vencimiento."
'                    .Label125.Caption = "-Se recibirán facturas/guias emitidas en el mes anterior (regularización de guias) hasta el 2° día hábil del mes en curso."
'                    .Label126.Caption = "-La empresa se reserva el derecho de devolver la mercadería que no cumpla con las especificaciones solicitadas."
'                    .Label127.Caption = vbNullString
                Else
'                    .LblTitle.Caption = "ORDEN DE SERVICIO"
'                    .Label120.Caption = "-Recepción de facturas de CONTRATISTAS en Oficina Central el día Lunes entre las 08:15 – 13:00 y 14:00 – 18:00."
'                    .Label121.Caption = "-Los Contratistas deberán adjuntar la Valorización emitida por obra con fecha, sello, nombre y firma de autorización para programación pago y copia de la O/S. En caso de ser primer adelanto, adjuntar CONTRATO debidamente firmado"
'                    .Label122.Caption = "-En caso de estar afecto a detracción, colocar el sello legible con el número de cuenta respectivo."
'                    .Label123.Caption = "-Recepción de facturas de Proveedores en Oficina Central los días Lunes y Miércoles; entre las 08:15 – 13:00 y 14:00 – 18:00. "
'                    .Label124.Caption = "-Los PROVEEDORES deberán adjuntar copia de la O/S y la Guía de Remisión (Destinatario y Sunat) con sello, nombre y firma de almacén por la conformidad en la recepción de los productos en obra."
'                    .Label125.Caption = "-La fecha de  vencimiento de las facturas se consideran a partir de la fecha de recepción en Oficina Central."
'                    .Label126.Caption = "-El pago de facturas se realizará el día Viernes de la semana correspondiente a la fecha de vencimiento."
'                    .Label127.Caption = "-Se recibirán facturas emitidas en el mes anterior hasta el 2° día hábil del mes en curso."
'    .Field334.Text = "Formato LOG-F-08"
                End If
                .LblNroOC.Caption = "N° " & RSCONSULTA.Fields("F4NUMORD")
                '.fldsolicitud.Text = "" & RSCONSULTA.Fields("F4CODSOLICITUD")
                 
                If Trim(RSCONSULTA.Fields("F4CODSOLICITUD") & "") <> vbNullString Then
                    objAyudaOrden.TipoOrden = strTipoOrden
                    objAyudaOrden.NumeroOrden = strNumeroOrden
                    
                    .fldsolicitud.Text = objAyudaOrden.generarCadenaSolicitud()
                End If
                
                
                If Val(RSCONSULTA!F4ESTADO & "") = 1 Then
                    '.LblEstado.Visible = True
                Else
                    .LblEstado.Visible = False
                End If

                
                .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
                .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
                .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
                .F2CONTACTO.Visible = True
                .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
                .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
                .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
                .FldFchEntrega = "" & Format(RSCONSULTA.Fields("F4FECent"), "dd/mm/yyyy")
                .FldTipCam.Text = Format(Val("" & RSCONSULTA.Fields("F4tipcam")), "0.000")
                .FldTipDoc.Text = UCase("" & ObtenerCampo("DOCUMENTOS", "F2DESDOC", "F2CODDOC", RSCONSULTA!F4TIPDOC & "", "T", cnn_dbbancos))
                If RSCONSULTA!F4TIPDOC & "" = "02" Then
                   .LblImp.Caption = "Reten."
                Else
                    .LblImp.Caption = "I.G.V."
                End If
                .FldObservaAll.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
                .FldSon.Text = CADENANUM(Val("" & RSCONSULTA.Fields("F4MONTO")), "" & RSCONSULTA.Fields("F4TIPMON"), "")
                If RsPago.State = adStateOpen Then RsPago.Close
                RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsPago.EOF Then
                    .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
                End If
                RsPago.Close
                
                .F4COTIZACION.Text = "" & RSCONSULTA.Fields("F4NUMCOTIZA")
                .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
                If Rs.State = 1 Then Rs.Close
                
                Rs.Open "Select F2NOMUSER From EF2USERS where F2CODUSER = '" & wusuario & "'", cnn_dbbancos, _
                adOpenKeyset, adLockOptimistic
                
                '.LBLFIRMA.Caption = rs(0)  ' Trim("" & pnlnomsoli.Caption)
                '.lblcargo.Caption = traerCampo("EF2USERS", "F2CARGO", "F2CODUSER", wusuario & "")
                '.lblempresa.Caption = wnomcia
            End If
            .DataControl1.ConnectionString = cnn_dbbancos
            
            CadSql = vbNullString
            CadSql = CadSql & "SELECT "
            CadSql = CadSql & "DET.F3CODFAB AS F3CODPRO, "
            CadSql = CadSql & "DET.F5NOMPRO, "
            CadSql = CadSql & "COLOR.DESCRIPCION AS DESCOLOR, "
            CadSql = CadSql & "MED.F7SIGMED AS F3MEDIDA, "
            CadSql = CadSql & "SUM(DET.F3CANPRO * (1 + (F3PORCDEMASIA/100))) AS F3CANPRO, "
            'CadSql = CadSql & "SUM(DET.F3PREUNI) AS F3PREUNI, "
            'CadSql = CadSql & "SUM(DET.F3PRENETO) AS F3PREUNI, "
            CadSql = CadSql & "DET.F3PRENETO AS F3PREUNI, "
            CadSql = CadSql & "SUM(DET.F5VALVTA) AS F5VALVTA, "
            CadSql = CadSql & "SUM(DET.F3IGV) AS F3IGV, "
            CadSql = CadSql & "SUM(DET.F3PORDCT) AS F3PORDCT, "
            CadSql = CadSql & "SUM(DET.F3TOTAL) AS F3TOTAL "
            CadSql = CadSql & "FROM "
            CadSql = CadSql & "((IF3ORDEN AS DET "
            CadSql = CadSql & "LEFT JOIN CENTROS AS CC ON CC.F3COSTO = DET.F3CENCOS) "
            CadSql = CadSql & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
            CadSql = CadSql & "LEFT JOIN EF2BIENCOLOR AS COLOR ON COLOR.CODIGO = DET.CODCOLOR "
            CadSql = CadSql & "WHERE "
            CadSql = CadSql & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
            CadSql = CadSql & "DET.F4LOCAL = '" & strTipoOrden & "' "
            CadSql = CadSql & "GROUP BY "
            CadSql = CadSql & "DET.F3CODFAB,  "
            CadSql = CadSql & "DET.F5NOMPRO, "
            CadSql = CadSql & "COLOR.DESCRIPCION, "
            CadSql = CadSql & "MED.F7SIGMED, "
            CadSql = CadSql & "DET.F3PRENETO"
            'CadSql = CadSql & "ORDER BY "
            'CadSql = CadSql & "VAL(DET.ITEM) "
            
            .DataControl1.Source = CadSql
            .Caption = "ORDEN DE COMPRA NACIONAL"
            RSCONSULTA.Close
            
            .DescargarReporte = True
            .TipoOrden = Trim(Txt_TOC.Text)
            .NumeroOrden = Trim(Txt_NumOC.Text)
            
            .Show
            
            '.Visible = False
            
            
            
            '*********************************************
            
    '        .Lbl1.ForeColor = RGB(48, 112, 168)
    '        .Lbl2.LineColor = RGB(48, 112, 168)
    '        .Lbl3.ForeColor = RGB(48, 112, 168)
    '        .Lbl4.ForeColor = RGB(48, 112, 168)
            
            '.Show vbModal
        End With
    ElseIf Opcion = 2 Then
        With Acr_OrdenCompra
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            
            sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
                    "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA, " & _
                    "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT,A.CODPRV,A.F4TIPMON " & _
                    " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & strNumeroOrden & _
                    " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"
        
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            
            If Not RSCONSULTA.EOF Then
                '.F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
                .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
                .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").value
                .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
                .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
                .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
                .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
                .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
                .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
                If RSCONSULTA.Fields("F4TIPMON") = "S" Then
                  '.TIPMON.Text = "SOLES"
                  Else
                  If RSCONSULTA.Fields("F4TIPMON") = "D" Then
                  '.TIPMON.Text = "DOLARES"
                  End If
                End If
        '
                If RsPago.State = adStateOpen Then RsPago.Close
                RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsPago.EOF Then
                    .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
                End If
                RsPago.Close
        '
                If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
                RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
                If Not RsCTR_COM.EOF Then
                    '.F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
                    '.F4NOTA.Text = "" & RsCTR_COM.Fields("F1NOTA_OC")
                    '.F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                    '.F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
                End If
                RsCTR_COM.Close
                .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
        
            End If
            
            RSCONSULTA.Close
            
            .DataControl1.ConnectionString = cnn_form
            .DataControl1.Source = "select * from TMPORDENDECOMPRA"
            
            .Show 'vbModal
        End With
    End If
    
    Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Logistica"
    'Resume
    Resume Next
End Sub

Rem SK ADD:----------------------------------------

'Private Function generarCadenaSolicitud() As String
'
'    Screen.MousePointer = vbHourglass
'
'    Dim rstSolicitudOC As New ADODB.Recordset
'
'    generarCadenaSolicitud = vbNullString
'
'    If rstSolicitudOC.State = 1 Then rstSolicitudOC.Close
'
'    rstSolicitudOC.Open "SELECT COD_SOLICITUD FROM IF3ORDEN WHERE F4LOCAL = '" & Trim(Txt_TOC.Text) & "' AND F4NUMORD = '" & Trim(Txt_NumOC.Text) & "' GROUP BY COD_SOLICITUD", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstSolicitudOC.EOF Then
'        rstSolicitudOC.MoveFirst
'
'        Do While Not rstSolicitudOC.EOF
'            If generarCadenaSolicitud = vbNullString Then
'                generarCadenaSolicitud = Trim(rstSolicitudOC!COD_SOLICITUD & "") & "-" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", Trim(rstSolicitudOC!COD_SOLICITUD & ""), "T", "AND CS_DOCUMENTO = '" & Trim(Txt_TOC.Text) & "'")
'            Else
'                generarCadenaSolicitud = generarCadenaSolicitud & vbCrLf & Trim(rstSolicitudOC!COD_SOLICITUD & "") & "-" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", Trim(rstSolicitudOC!COD_SOLICITUD & ""), "T", "AND CS_DOCUMENTO = '" & Trim(Txt_TOC.Text) & "'")
'            End If
'
'            rstSolicitudOC.MoveNext
'        Loop
'    End If
'
'    Screen.MousePointer = vbDefault
'End Function

'--------------------------------------------------

Private Sub Imprime_Orden_Ditec(Opcion)
    Dim sql As String
    Dim RSCONSULTA As New ADODB.Recordset
    Set RSCONSULTA = New ADODB.Recordset
    Dim RsPago As New ADODB.Recordset
    Set RsPago = New ADODB.Recordset
    Dim RsCTR_COM As New ADODB.Recordset
    Set RsCTR_COM = New ADODB.Recordset
    
    If Opcion = 1 Then
        With Acr_OrdenC_Ditec
            If Cmbmone.ListIndex = 0 Then
                .lblmoneda4.Caption = "S/" '"S/"
                .lblmoneda2.Caption = "S/" '"S/"
                .lblmoneda1.Caption = "S/" '"S/"
                .lblmoneda3.Caption = "S/" '"S/"
            Else
                .lblmoneda4.Caption = "US$"
                .lblmoneda2.Caption = "US$"
                .lblmoneda1.Caption = "US$"
                .lblmoneda3.Caption = "US$"
            End If
            .flddirec1.Text = wf1direc1
            .flddirec2.Text = wf1direc2
            .fldruc.Text = wrucempresa
            .FldEmpresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
            
            '.IGV.Caption = wigv
            GOC = Txt_NumOC.Text
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
                  "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4RND,A.F4OBSERVA, " & _
                  "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT,A.F4OBSERVA " & _
                  " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & GOC & _
                  "' AND A.F4ESTNUL<>'S'AND A.f4local='" & TOC & "' ORDER BY A.F4NUMORD DESC;"
        
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
                '.F4NUMORD.Text = ("" & rsconsulta.Fields("F4NUMORD"))
                .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
                .F4CODSOLICITUD.Text = "" & RSCONSULTA.Fields("F4CODSOLICITUD")
                .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
                .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
                .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
                .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
                .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
                .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
                .F4IGV.Text = Format("" & RSCONSULTA.Fields("F4IGV"), "0.00")
                .F4RND.Text = Format("" & RSCONSULTA.Fields("F4RND"), "0.00")
                .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "0.00")
                .Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                '.F4OBSERVA.Text = "" & rsconsulta.Fields("F4OBSERVA")
                .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "0.00")
                .F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
                '.F4NOTA.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                .referencia.Text = "" & Txt_Referencia.Text
                .solicitado.Text = "" & pnlnomsoli.Caption
                If RsPago.State = adStateOpen Then RsPago.Close
                RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsPago.EOF Then
                    .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
                End If
                RsPago.Close
                If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
                RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
                If Not RsCTR_COM.EOF Then
                    .F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
    '                .F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                    .F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
                End If
                RsCTR_COM.Close
                .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
                .F4EMITIR.Text = "" & wnomcia
                .F4EMITIR1.Text = "R.U.C. " & wrucempresa
                .f4emitir2.Text = "" & wdireccion & " - Perú"
                .f4emitir3.Text = "Ph: " & wtelefono & "  Fax: " & wfax
                
                If Rs.State = 1 Then Rs.Close
                
                Rs.Open "Select F2NOMUSER From EF2USERS where F2CODUSER = '" & wusuario & "'", cnn_dbbancos, _
                adOpenKeyset, adLockOptimistic
                
                '.LBLFIRMA.Caption = rs(0)  ' Trim("" & pnlnomsoli.Caption)
                '.lblcargo.Caption = traerCampo("EF2USERS", "F2CARGO", "F2CODUSER", wusuario & "")
                '.lblempresa.Caption = wnomcia
            End If
            .DataControl1.ConnectionString = cnn_form
            .DataControl1.Source = "SELECT * FROM TMPORDENDECOMPRA"
            '.F3FECEN.Text = dxDBGrid1.Columns.ColumnByFieldName("F3ENTREGA").Value
            RSCONSULTA.Close
            .Lbl1.ForeColor = RGB(48, 112, 168)
            .Lbl2.LineColor = RGB(48, 112, 168)
            .Lbl3.ForeColor = RGB(48, 112, 168)
            .Lbl4.ForeColor = RGB(48, 112, 168)
            .Caption = "ORDEN DE COMPRA NACIONAL"
            .Show vbModal
        End With
        
    
    
    ElseIf Opcion = 2 Then
    With Acr_OrdenC_Ditec
        If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
        sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
        "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA, " & _
        "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT " & _
        " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & GOC & _
        " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"
    
        RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RSCONSULTA.EOF Then
            '.F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
            .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
            .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").value
            .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
            .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
            .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
            .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
            .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
            .F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
    '
            If RsPago.State = adStateOpen Then RsPago.Close
            RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsPago.EOF Then
                .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
            End If
            RsPago.Close
    '
            If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
            RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
            If Not RsCTR_COM.EOF Then
                .F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
                .F4NOTA.Text = "" & RsCTR_COM.Fields("F1NOTA_OC")
                .F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                .F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
            End If
            RsCTR_COM.Close
            .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
    
        End If
        RSCONSULTA.Close
    
        .DataControl1.ConnectionString = cnn_form
        .DataControl1.Source = "select * from TMPORDENDECOMPRA"
    
        .Show vbModal
    End With
    End If
End Sub

Private Sub Imprime_Orden_Electrica(Opcion)
    Dim sql As String
    Dim RSCONSULTA As New ADODB.Recordset
    Set RSCONSULTA = New ADODB.Recordset
    Dim RsPago As New ADODB.Recordset
    Set RsPago = New ADODB.Recordset
    Dim RsCTR_COM As New ADODB.Recordset
    Set RsCTR_COM = New ADODB.Recordset
    
    If Opcion = 1 Then
        With Acr_OrdenC_Electrica
            If Cmbmone.ListIndex = 0 Then
                .lblmoneda4.Caption = "S/" '"S/"
                .lblmoneda2.Caption = "S/" '"S/"
                .lblmoneda1.Caption = "S/" '"S/"
                .lblmoneda3.Caption = "S/" '"S/"
            Else
                .lblmoneda4.Caption = "US$"
                .lblmoneda2.Caption = "US$"
                .lblmoneda1.Caption = "US$"
                .lblmoneda3.Caption = "US$"
            End If
            .flddirec1.Text = wf1direc1
            .flddirec2.Text = wf1direc2
            .fldruc.Text = wrucempresa
            .FldEmpresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
            
            '.IGV.Caption = wigv
            GOC = Txt_NumOC.Text
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
                  "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4RND,A.F4OBSERVA, " & _
                  "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT,A.F4OBSERVA " & _
                  " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & GOC & _
                  "' AND A.F4ESTNUL<>'S'AND A.f4local='" & TOC & "' ORDER BY A.F4NUMORD DESC;"
        
            RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
                '.F4NUMORD.Text = ("" & rsconsulta.Fields("F4NUMORD"))
                .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
                .F4CODSOLICITUD.Text = "" & RSCONSULTA.Fields("F4CODSOLICITUD")
                .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
                .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
                .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
                .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
                .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
                .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
                .F4IGV.Text = Format("" & RSCONSULTA.Fields("F4IGV"), "0.00")
                .F4RND.Text = Format("" & RSCONSULTA.Fields("F4RND"), "0.00")
                .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "0.00")
                .Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                '.F4OBSERVA.Text = "" & rsconsulta.Fields("F4OBSERVA")
                .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "0.00")
                .F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
                .F4NOTA.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
                .referencia.Text = "" & Txt_Referencia.Text
                .solicitado.Text = "" & pnlnomsoli.Caption
                If RsPago.State = adStateOpen Then RsPago.Close
                RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsPago.EOF Then
                    .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
                End If
                RsPago.Close
                If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
                RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
                If Not RsCTR_COM.EOF Then
                    .F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
    '                .F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                    .F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
                End If
                RsCTR_COM.Close
                .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
                .F4EMITIR.Text = "" & wnomcia
                .F4EMITIR1.Text = "R.U.C. " & wrucempresa
                .f4emitir2.Text = "" & wdireccion & " - Perú"
                .f4emitir3.Text = "Ph: " & wtelefono & "  Fax: " & wfax
                
                If Rs.State = 1 Then Rs.Close
                
                Rs.Open "Select F2NOMUSER From EF2USERS where F2CODUSER = '" & wusuario & "'", cnn_dbbancos, _
                adOpenKeyset, adLockOptimistic
                
                '.LBLFIRMA.Caption = rs(0)  ' Trim("" & pnlnomsoli.Caption)
                '.lblcargo.Caption = traerCampo("EF2USERS", "F2CARGO", "F2CODUSER", wusuario & "")
                '.lblempresa.Caption = wnomcia
            End If
            .DataControl1.ConnectionString = cnn_form
            .DataControl1.Source = "SELECT * FROM TMPORDENDECOMPRA"
            '.F3FECEN.Text = dxDBGrid1.Columns.ColumnByFieldName("F3ENTREGA").Value
            RSCONSULTA.Close
            '.Lbl1.ForeColor = RGB(48, 112, 168)
            '.Lbl2.LineColor = RGB(48, 112, 168)
            '.Lbl3.ForeColor = RGB(48, 112, 168)
            '.Lbl4.ForeColor = RGB(48, 112, 168)
            .Caption = "ORDEN DE COMPRA NACIONAL"
            .Show vbModal
        End With
        
    
    
    ElseIf Opcion = 2 Then
    With Acr_OrdenC_Electrica
        If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
        sql = "SELECT A.F4NUMORD, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
        "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4OBSERVA, " & _
        "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT " & _
        " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & GOC & _
        " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"
    
        RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RSCONSULTA.EOF Then
            '.F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
            .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
            .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").value
            .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
            .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
            .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
            .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
            .F4FECEMI.Text = Format(RSCONSULTA.Fields("F4FECEMI"), "dd") & " de " & dev_mes(Month(RSCONSULTA.Fields("F4FECEMI"))) & " de " & Format(RSCONSULTA.Fields("F4FECEMI"), "yyyy")
            .F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
    '
            If RsPago.State = adStateOpen Then RsPago.Close
            RsPago.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsPago.EOF Then
                .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
            End If
            RsPago.Close
    '
            If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
            RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
            If Not RsCTR_COM.EOF Then
                .F4OBSFECHA.Text = "" & RsCTR_COM.Fields("F1OBSFECENT_OC")
                .F4NOTA.Text = "" & RsCTR_COM.Fields("F1NOTA_OC")
                .F4EMITIR.Text = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                .F4OBSGEN.Text = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
            End If
            RsCTR_COM.Close
            .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
    
        End If
        RSCONSULTA.Close
    
        .DataControl1.ConnectionString = cnn_form
        .DataControl1.Source = "select * from TMPORDENDECOMPRA"
    
        .Show vbModal
    End With
    End If
End Sub

Private Sub Email()

End Sub

Private Sub Calcula_PvtaTot()
    Dim Cantidad    As Double
    Dim totdcto     As Double
    Dim ValVta      As Double
    Dim IGV         As Double
    Dim preciounit  As Double
    Dim TOTAL       As Double
    Dim costo       As Double

    With Grid
        Cantidad = Val(Format(.Columns.ColumnByFieldName("F3CANPRO").value, "0.00"))
        If Cantidad > 0 Then
            'If cmbtipopera.ListIndex = 0 Then
                If .Columns.ColumnByFieldName("F5AFECTO").value = "*" Then     'Afecto
                    totdcto = (Val(Format("" & .Columns.ColumnByFieldName("F3canpro").value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PREcos").value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").value, "0.00"))) / 100
                    .Columns.ColumnByFieldName("F3TOTDCT").value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(Cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) - totdcto, "0.00"))
                    .Columns.ColumnByFieldName("F5VALVTA").value = Format$(ValVta, "###,##0.00")
                    IGV = ValVta * (wwigv / 100)
                    .Columns.ColumnByFieldName("F3IGV").value = Format$(IGV, "#,##0.00")
                    preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) + (Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) * (wwigv / 100))
                    .Columns.ColumnByFieldName("F3PREUNI").value = Format$(preciounit, "###,##0.00")
                    TOTAL = ValVta + IGV
                    .Columns.ColumnByFieldName("F3TOTAL").value = Format$(TOTAL, "###,##0.00")
                Else  'Inafecto
                    IGV = 0
                    .Columns.ColumnByFieldName("F3IGV").value = Format$(IGV, "0.00")
                    totdcto = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").value, "0.00")) / 100
                    .Columns.ColumnByFieldName("F3TOTDCT").value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(Cantidad * Val(Format(.Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) - totdcto, "0.00"))
                    .Columns.ColumnByFieldName("F5VALVTA").value = Format$(ValVta, "###,##0.00")
                    preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000"))
                    .Columns.ColumnByFieldName("F3PREUNI").value = Format$(preciounit, "###,##0.00")
                    TOTAL = ValVta + IGV
                    .Columns.ColumnByFieldName("F3TOTAL").value = Format$(TOTAL, "###,##0.00")
                End If
            'Else
            '    costo = Val(Format(.Columns.ColumnByFieldName("F3PRECOS").Value, "###,##0.00"))
            '    TOTAL = cantidad * costo                '
            '    .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
            'End If
        End If
    End With
End Sub

Private Sub Calcula_PvtaTotalGrid()
    Dim Cantidad    As Double
    Dim totdcto     As Double
    Dim ValVta      As Double
    Dim IGV         As Double
    Dim preciounit  As Double
    Dim TOTAL       As Double
    Dim costo       As Double
    Dim K As Integer

    With Grid
    .Dataset.First
        For K = 0 To .Dataset.RecordCount - 1
        If .Columns.ColumnByFieldName("F3PREUNI").value = 0 Then
            .Dataset.Edit
            Cantidad = Val(Format(.Columns.ColumnByFieldName("F3CANPRO").value, "0.00"))
            If Cantidad > 0 Then
                'If cmbtipopera.ListIndex = 0 Then
                    If .Columns.ColumnByFieldName("F5AFECTO").value = True Then     'Afecto
                        totdcto = (Val(Format("" & .Columns.ColumnByFieldName("F3canpro").value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PREcos").value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").value, "0.00"))) / 100
                        .Columns.ColumnByFieldName("F3TOTDCT").value = Format$(totdcto, "####,##0.00")
                        ValVta = Val(Format(Cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3sinigv").value, "0.0000")) - totdcto, "0.00"))
                        .Columns.ColumnByFieldName("F5VALVTA").value = Format$(ValVta, "###,##0.00")
                        IGV = ValVta * (wwigv / 100)
                        .Columns.ColumnByFieldName("F3IGV").value = Format$(IGV, "#,##0.00")
                        preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3sinigv").value, "0.0000")) + (Val(Format("" & .Columns.ColumnByFieldName("F3sinigv").value, "0.0000")) * (wwigv / 100))
                        .Columns.ColumnByFieldName("F3PREUNI").value = Format$(preciounit, "###,##0.00")
                        TOTAL = ValVta + IGV
                        .Columns.ColumnByFieldName("F3TOTAL").value = Format$(TOTAL, "###,##0.00")
                    Else  'Inafecto
                        IGV = 0
                        .Columns.ColumnByFieldName("F3IGV").value = Format$(IGV, "0.00")
                        totdcto = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").value, "0.00")) / 100
                        .Columns.ColumnByFieldName("F3TOTDCT").value = Format$(totdcto, "####,##0.00")
                        ValVta = Val(Format(Cantidad * Val(Format(.Columns.ColumnByFieldName("F3PRECOS").value, "0.0000")) - totdcto, "0.00"))
                        .Columns.ColumnByFieldName("F5VALVTA").value = Format$(ValVta, "###,##0.00")
                        preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").value, "0.0000"))
                        .Columns.ColumnByFieldName("F3PREUNI").value = Format$(preciounit, "###,##0.00")
                        TOTAL = ValVta + IGV
                        .Columns.ColumnByFieldName("F3TOTAL").value = Format$(TOTAL, "###,##0.00")
                    End If
            End If
            .Dataset.Post
        End If
        .Dataset.Next
        Next K
    End With
End Sub

Private Sub importarDatosRequerimiento()
    Dim sw_nuevo_temp   As Boolean
    Dim xnombre         As String
    Dim i               As Integer
    Dim entrega         As Date
    Dim J               As Integer
    
    csql = vbNullString
    csql = csql & "SELECT "
    csql = csql & "* "
    csql = csql & "FROM "
    csql = csql & "TB_CABSOLICITUD "
    csql = csql & "WHERE "
    csql = csql & "CS_DOCUMENTO = '" & objAyudaSolicitud.TipoDocumento & "' AND "
    csql = csql & "COD_SOLICITUD = '" & objAyudaSolicitud.Codigo & "'"
    
    Set rssolcab = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    With rssolcab
        If Not .EOF And Not .Bof Then
            Rem SK ADD:
            pnlnomprv.Caption = "Ruc es menor a 11 digitos"
            pnldireprv.Caption = "No tiene"
            
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2NEWRUC", Trim(!CS_PROVEEDOR & ""), "T") <> vbNullString Then
                Txt_Prove.Text = Trim(!CS_PROVEEDOR & "")
                pnlnomprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(!CS_PROVEEDOR & ""), "T") 'rst!F2NOMPROV
                pnldireprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(!CS_PROVEEDOR & ""), "T") 'IIf(IsNull(rst!F2DIRPROV), " ", rst!F2DIRPROV)
            End If
            
            Txt_NumSolComp.Text = Trim(!COD_SOLICITUD & "")
            txtobserva.Text = Trim("" & !CS_OBSERVACIONES)
            
            If Txt_TOC = "OS" Then
                lblAutorizado.Visible = True
                txtAutorizado.Visible = True
                txtAutorizado.Text = Trim("" & !DERIVADO & "")
            End If
            
            'txtAutorizado.Text = Trim("" & !DERIVADO & "")
            
            pnlnomsoli.Caption = vbNullString
            
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODUSER", "EF2USERS", "F2CODUSER", Trim(!CS_CODSOLICITANTE & ""), "T") <> vbNullString Then
                txtcodsoli.Text = Trim(!CS_CODSOLICITANTE & "")
                pnlnomsoli.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", Trim(!CS_CODSOLICITANTE & ""), "T")
            End If
            
            If Trim(!CS_MONEDA & "") <> vbNullString Then
                Cmbmone.ListIndex = ModUtilitario.seleccionarItem(Cmbmone, Trim(!CS_MONEDA & ""), "IZQ", 1)
            End If
            
            TxtCodCosto.Text = Trim(!CS_CODCOSTO & "")
            txtcodcosto_KeyPress 13
            
            txtlugar_entrega.Text = left(Trim("" & !CS_LUGENTR), 100)
        End If
        
        rssolcab.Close
    End With
    
    '*** detalle de solicitud de compra
    'Versión Nueva
    With Grid
'        csql = vbNullString
'        csql = csql & "SELECT "
'        csql = csql & "TB_DETSOLICITUD.cod_solicitud, "
'        csql = csql & "TB_DETSOLICITUD.observa, "
'        csql = csql & "TB_DETSOLICITUD.ITEM, "
'        csql = csql & "TB_CABSOLICITUD.cs_fecha, "
'        csql = csql & "TB_CABSOLICITUD.cs_codsolicitante, "
'        csql = csql & "EF2USERS.F2NOMUSER, "
'        csql = csql & "TB_CABSOLICITUD.cs_observaciones, "
'        csql = csql & "TB_DETSOLICITUD.F5CODCOSTO, "
'        csql = csql & "CENTROS.F3DESCRIP, "
'        csql = csql & "TB_DETSOLICITUD.cod_producto, "
'        csql = csql & "TB_DETSOLICITUD.precio, "
'        csql = csql & "TB_DETSOLICITUD.ds_descripcion, "
'        csql = csql & "TB_DETSOLICITUD.ds_unidmed, "
'        csql = csql & "EF7MEDIDAS.F7SIGMED, "
'        csql = csql & "TB_DETSOLICITUD.ds_cantidad, "
'        csql = csql & "IIf(IsNull(ORDENES.CANT_ORDEN),0,ORDENES.CANT_ORDEN) AS TOT_ORDEN, "
'        csql = csql & "TB_DETSOLICITUD.ds_cantidad-IIf(IsNull(ORDENES.CANT_ORDEN),0,ORDENES.CANT_ORDEN) AS SALDO, "
'        csql = csql & "TB_DETSOLICITUD.f5SINigv, "
'        csql = csql & "TB_DETSOLICITUD.f5CONigv, "
'        csql = csql & "TB_DETSOLICITUD.ruc_proveedor, "
'        csql = csql & "TB_DETSOLICITUD.F5AFECTO "
'        csql = csql & "FROM (TB_CABSOLICITUD LEFT JOIN EF2USERS ON TB_CABSOLICITUD.cs_codsolicitante = EF2USERS.F2CODUSER) "
'        csql = csql & "INNER JOIN (((TB_DETSOLICITUD LEFT JOIN ["
'        csql = csql & "SELECT IF3ORDEN.COD_SOLICITUD, IF3ORDEN.F3CENCOS, IF3ORDEN.F3CODPRO, Sum(IF3ORDEN.F3CANPRO) AS CANT_ORDEN "
'        csql = csql & "From IF3ORDEN "
'        csql = csql & "GROUP BY IF3ORDEN.COD_SOLICITUD, IF3ORDEN.F3CENCOS, IF3ORDEN.F3CODPRO "
'        csql = csql & "ORDER BY IF3ORDEN.COD_SOLICITUD, IF3ORDEN.F3CENCOS, IF3ORDEN.F3CODPRO"
'        csql = csql & "]. AS ORDENES "
'        csql = csql & "ON (TB_DETSOLICITUD.cod_producto = ORDENES.F3CODPRO) AND (TB_DETSOLICITUD.F5CODCOSTO = ORDENES.F3CENCOS) "
'        csql = csql & "AND (TB_DETSOLICITUD.cod_solicitud = ORDENES.COD_SOLICITUD)) "
'        csql = csql & "LEFT JOIN CENTROS ON TB_DETSOLICITUD.F5CODCOSTO = CENTROS.F3COSTO) LEFT JOIN EF7MEDIDAS ON "
'        csql = csql & "TB_DETSOLICITUD.ds_unidmed = EF7MEDIDAS.F7CODMED) ON TB_CABSOLICITUD.cod_solicitud = TB_DETSOLICITUD.cod_solicitud "
'        csql = csql & "Where ((([TB_DETSOLICITUD].[ds_cantidad] - IIf(IsNull([ORDENES].[CANT_ORDEN]), 0, [ORDENES].[CANT_ORDEN])) > 0) "
'        csql = csql & "And ((TB_CABSOLICITUD.cs_estado) >= '2')) "
'        csql = csql & "AND TB_DETSOLICITUD.cod_solicitud='" & num_solcomp & "'"
'        If item_solcomp > 0 Then
'            csql = csql & "and TB_DETSOLICITUD.item=" & item_solcomp & " "
'        End If
'
'        csql = csql & "ORDER BY TB_DETSOLICITUD.ITEM"
        
        csql = vbNullString
        csql = csql & "SELECT "
        csql = csql & "DET.CS_DOCUMENTO, "
        csql = csql & "DET.COD_SOLICITUD, "
        csql = csql & "DET.ITEM, "
        csql = csql & "DET.COD_PRODUCTO, "
        csql = csql & "DET.DS_DESCRIPCION, "
        csql = csql & "DET.DS_UNIDMED, "
        csql = csql & "MED.F7SIGMED, "
        csql = csql & "DET.DS_CANTIDAD, "
        csql = csql & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS SALDO, "
        csql = csql & "DET.OBSERVA "
        csql = csql & "FROM "
        csql = csql & "(TB_DETSOLICITUD AS DET "
        csql = csql & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.DS_UNIDMED) "
        csql = csql & "LEFT JOIN "
        csql = csql & "("
        csql = csql & "SELECT "
        csql = csql & "DET.COD_SOLICITUD, "
        csql = csql & "DET.F3CODPRO, "
        csql = csql & "SUM(DET.F3CANPRO) AS CANTIDAD "
        csql = csql & "FROM "
        csql = csql & "IF3ORDEN AS DET "
        csql = csql & "WHERE "
        csql = csql & "DET.COD_SOLICITUD <> '' "
        csql = csql & "GROUP BY "
        csql = csql & "DET.COD_SOLICITUD, DET.F3CODPRO) AS MOVPROD "
        csql = csql & "ON MOVPROD.COD_SOLICITUD = DET.COD_SOLICITUD AND  MOVPROD.F3CODPRO = DET.COD_PRODUCTO "
        csql = csql & "WHERE "
        csql = csql & "DET.CS_DOCUMENTO = '" & objAyudaSolicitud.TipoDocumento & "' AND "
        csql = csql & "DET.COD_SOLICITUD = '" & objAyudaSolicitud.Codigo & "' AND "
        
            If objAyudaSolicitud.CodProducto <> vbNullString Then
                csql = csql & "DET.COD_PRODUCTO = '" & objAyudaSolicitud.CodProducto & "' AND "
            End If
            
        csql = csql & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) > 0 "
        csql = csql & "ORDER BY "
        csql = csql & "DET.DS_DESCRIPCION"
        
        Set rsSolDet = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
        
        If Not (rsSolDet.EOF) Then
            sw_nuevo_temp = False
            sw_nuevo_item = False
            
            rsSolDet.MoveFirst
            
            J = 0
            
            Do While Not (rsSolDet.EOF)
                'If rsSolDet!COD_SOLICITUD = Trim(num_solcomp) Then
                    J = J + 1
                    
                    If J = 1 Then
                        Grid.Dataset.Edit
                    Else
                        Grid.Dataset.Append
                    End If
                    
                    .Dataset.FieldValues("f3codpro") = rsSolDet!COD_PRODUCTO & ""
                    .Dataset.FieldValues("f5nompro") = rsSolDet!ds_descripcion & ""
                    .Dataset.FieldValues("f5nompro_ing") = rsSolDet!ds_descripcion & ""
                    .Dataset.FieldValues("f3canpro") = rsSolDet!SALDO
                    '.Dataset.FieldValues("f3redondeo") = rsSolDet!Saldo
                    .Dataset.FieldValues("f3canproMAX") = rsSolDet!SALDO
                    .Dataset.FieldValues("f3codmedida") = rsSolDet!ds_unidmed & ""
                    .Dataset.FieldValues("f3desmedida") = rsSolDet!F7SIGMED & "" 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsSolDet!ds_unidmed & "", "T", cnn_dbbancos)
                    '.Dataset.FieldValues("f3sinigv") = Val(Format(rsSolDet!f5SINigv, "0.0000"))
                    .Dataset.FieldValues("f5afecto") = IIf(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5AFECTO", "IF5PLA", "F5CODPRO", Trim(rsSolDet!COD_PRODUCTO & ""), "T") = "*", True, False) 'rsSolDet!F5AFECTO
                    '.Dataset.FieldValues("f3conigv") = Val(Format(rsSolDet!f5CONigv, "0.0000"))
                    '.Dataset.FieldValues("f3valdesc") = 0
                    '.Dataset.FieldValues("f3pordesc") = 0
                    '.Dataset.FieldValues("f3total") = 0
                    .Dataset.FieldValues("cod_solicitud") = rsSolDet!COD_SOLICITUD & ""
                    .Dataset.FieldValues("f3observa") = rsSolDet!OBSERVA & ""
                    
                    .Dataset.FieldValues("F3PORCDEMASIA") = Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "PORCDEMASIA", "IF5PLA", "F5CODPRO", Trim(rsSolDet!COD_PRODUCTO & ""), "T"))
                    
                    .Dataset.Post
                'End If
                
                rsSolDet.MoveNext
            Loop
            
            sw_nuevo_item = False
        End If
        
        rsSolDet.Close
    End With
End Sub


Rem SK ADD:----------------------------------------------------------------------------------------------------------
Private Sub copiarSeleccionAyudaProductos()
    Dim rstProductoOC As New ADODB.Recordset
    Dim dblItem As Double
    Dim strUltimaDescripcion As String
    Dim dblUltimoPrecioSinIGv As Double
    'Dim dblUltimoDescuento As Double
    
    Me.MousePointer = vbHourglass
                
    Grid.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA WHERE TRIM(F3CODPRO & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "COD_SOLICITUD, F5CODPRO, F5CODFAB, F5NOMPRO, F5NOMPRO_ING, F7CODMED, "
    CadSql = CadSql & "IIF(TRIM(F5AFECTO & '') = '*', TRUE, FALSE) AS AFECTO, "
    CadSql = CadSql & "F2MONEDA, F5VTANET, F5VTANETDOL, F5FOB, F5FOBMAX "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPPRODUCTOS "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "F4PERINT = -1"
    
    If rstProductoOC.State = 1 Then rstProductoOC.Close
    
    rstProductoOC.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProductoOC.EOF Then
        rstProductoOC.MoveFirst
        
        Do While Not rstProductoOC.EOF
            'Obtener la Descripción del Producto en la Ultima Compra (Ordenes de Compra)
            With objAyudaOrden
                .CodProveedor = Trim(Txt_Prove.Text)
                .CodigoProducto = Trim(rstProductoOC!f5codpro & "")
                
                strUltimaDescripcion = .obtenerUltimaDescripcionProductoDeProveedor
            End With
            'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
            With objAyudaVale
                .CodigoProveedor = Trim(Txt_Prove.Text)
                .CodigoProducto = Trim(rstProductoOC!f5codpro & "")
                
                .obtenerUltimoPrecioSinIgvProductoDeProveedor
                
                Select Case .CodigoMoneda
                    Case "S"
                        dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text))), "#0.0000")
                    Case Else
                        dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text))), "#0.0000")
                End Select
            End With
            
            If ModUtilitario.ObtenerCampoV2(cnDBTemp, "F3CODPRO", "TMPORDENDECOMPRA", "F3CODPRO", Trim(rstProductoOC!f5codpro & ""), "T", "AND COD_SOLICITUD = '" & Trim(rstProductoOC!COD_SOLICITUD & "") & "'") = vbNullString Then
                dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPORDENDECOMPRA", vbNullString, vbNullString, vbNullString, "TRIM(F3CODPRO & '') <> '' ORDER BY ITEM DESC") & "") + 1
                
                'If dblItem = 1 Then
                '    cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA WHERE TRIM(F3CODPRO & '') = ''"
                'End If
                
                CadSql = vbNullString
                CadSql = CadSql & "INSERT INTO TMPORDENDECOMPRA(ITEM, COD_SOLICITUD, CLIENTE, F3CODPRO, F3CODFAB, F5NOMPRO, F5NOMPRO_ING, F3CODMEDIDA, F3DESMEDIDA, "
                CadSql = CadSql & "F5AFECTO, F3SINIGV, F3CONIGV, F3BASEIMP, F3MONINA, F3CANPRO, F3CANPROMAX, F3PORCDEMASIA) "
                CadSql = CadSql & "VALUES(" & dblItem & ", '" & Trim(rstProductoOC!COD_SOLICITUD & "") & "', "
                CadSql = CadSql & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_NOMREF", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(rstProductoOC!COD_SOLICITUD & ""), "N") & "', "
                CadSql = CadSql & "'" & Trim(rstProductoOC!f5codpro & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProductoOC!f5codfab & "") & "', "
                CadSql = CadSql & "'" & IIf(strUltimaDescripcion <> vbNullString, strUltimaDescripcion, Trim(rstProductoOC!F5NOMPRO & "")) & "', "
                CadSql = CadSql & "'" & Trim(rstProductoOC!F5NOMPRO_ING & "") & "', "
                CadSql = CadSql & "'" & Trim(rstProductoOC!f7codmed & "") & "', "
                CadSql = CadSql & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", Trim(rstProductoOC!f7codmed & ""), "T") & "', "
                CadSql = CadSql & IIf(CBool(rstProductoOC!Afecto), "TRUE", "FALSE") & ", "
                    
'                    If Val(rstProductoOC!F5VTANET & "") <> 0 Then
'                        CadSql = CadSql & Format(Val(rstProductoOC!F5VTANET & "") / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text)), "#0.0000") & ", "
'                    ElseIf Val(rstProductoOC!F5VTANETDOL & "") <> 0 Then
'                        CadSql = CadSql & Format(Val(rstProductoOC!F5VTANETDOL & "") * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text)), "#0.0000") & ", "
'                    Else
'                        CadSql = CadSql & "0, "
'                    End If
                    
                    CadSql = CadSql & dblUltimoPrecioSinIGv & ", "
                
                CadSql = CadSql & "0, "
                CadSql = CadSql & "0, "
                CadSql = CadSql & "0, "
                CadSql = CadSql & Val(rstProductoOC!F5FOB & "") & ", "
                CadSql = CadSql & Val(rstProductoOC!F5FOBMAX & "") & ", "
                CadSql = CadSql & Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "PORCDEMASIA", "IF5PLA", "F5CODPRO", Trim(rstProductoOC!f5codpro & ""), "T")) & ")"
            Else
                CadSql = vbNullString
                CadSql = CadSql & "UPDATE TMPORDENDECOMPRA "
                CadSql = CadSql & "SET "
                CadSql = CadSql & "F3CANPRO = " & Val(rstProductoOC!F5FOB & "") & ", "
                CadSql = CadSql & "F3CANPROMAX = " & Val(rstProductoOC!F5FOBMAX & "") & " "
                CadSql = CadSql & "WHERE "
                CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstProductoOC!COD_SOLICITUD & "") & "' AND "
                CadSql = CadSql & "F3CODPRO = '" & Trim(rstProductoOC!f5codpro & "") & "'"
            End If
            
            cnDBTemp.Execute CadSql
            
            rstProductoOC.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    dblItem = 0
    
    If rstProductoOC.State = 1 Then rstProductoOC.Close
    
    Set rstProductoOC = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub copiarSeleccionAyudaResumenRequerimiento(ByVal strNroPedido As String)
    Dim rstResumenReq As New ADODB.Recordset
    Dim dblItem As Double
    Dim strUltimaDescripcion As String
    Dim dblUltimoPrecioSinIGv As Double
    'Dim dblUltimoDescuento As Double
    
    Dim strCuentaContable As String
    
    Me.MousePointer = vbHourglass
                
    Grid.Dataset.Close
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA WHERE TRIM(F3CODPRO & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "NROPEDIDO, CODPRODUCTO, CANTIDADPC, UM "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPUTILRESUMENREQUERIMIENTO "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "PROCESAR = TRUE AND "
    CadSql = CadSql & "NROPEDIDO = '" & strNroPedido & "'"
    
    If rstResumenReq.State = 1 Then rstResumenReq.Close
    
    rstResumenReq.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstResumenReq.EOF Then
        rstResumenReq.MoveFirst
        
        Do While Not rstResumenReq.EOF
            If Trim(rstResumenReq!NroPedido & "") = strNroPedido Then
                'Obtener configuracion de Producto
                With objAyudaBien
                    .inicializarEntidades
                    
                    .Codigo = Trim(rstResumenReq!CodProducto & "")
                    
                    .obtenerConfigBien
                End With
                
                strCuentaContable = vbNullString
                
                Select Case strTipoOrden
                    Case "OC"
                        With objAyudaProveedor
                            .Codigo = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
                            
                            .obtenerConfigProveedor
                        End With
                        
                        Select Case objAyudaProveedor.OrigenProveedor
                            Case "N"
                                If objAyudaBien.CtaContable = vbNullString Then
'                                    MsgBox "Imposible adicionar el producto: " & vbNewLine & _
'                                            objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                            "ya que no tiene configurado su Cuenta Contable para Proveedores Nacionales." & vbNewLine & vbNewLine & _
'                                            "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
                                    
                                Else
                                    strCuentaContable = objAyudaBien.CtaContable
                                End If
                            Case "E"
                                If objAyudaBien.CtaContableImportacion = vbNullString Then
'                                    MsgBox "Imposible adicionar el producto: " & vbNewLine & _
'                                            objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                            "ya que no tiene configurado su Cuenta Contable para Proveedores Extranjeros." & vbNewLine & vbNewLine & _
'                                            "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
                                    
                                Else
                                    strCuentaContable = objAyudaBien.CtaContableImportacion
                                End If
                        End Select
                    Case "OS"
                        If objAyudaBien.CtaContable = vbNullString Then
'                            MsgBox "Imposible adicionar el servicio: " & vbNewLine & _
'                                    objAyudaBien.Descripcion & ";" & vbNewLine & _
'                                    "ya que no tiene configurado su Cuenta Contable." & vbNewLine & vbNewLine & _
'                                    "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
                                                    
                        Else
                            strCuentaContable = objAyudaBien.CtaContable
                        End If
                End Select
                
                If strCuentaContable <> vbNullString Then
                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T") = vbNullString Then
                        With objAyudaGasto
                            .inicializarEntidades
                            
                            .Codigo = vbNullString
                            .Base = "G"
                            .CuentaContable = strCuentaContable
                            .Descripcion = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", strCuentaContable, "T")
                            .TipoGasto = "P"
                            .Moneda = left(Cmbmone.Text, 1)
                            .GrupoFlujo = vbNullString
                            
                            If .guardarGasto Then
                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                            End If
                            
                            .inicializarEntidades
                        End With
                    End If
                    
                    'Obtener la Descripción del Producto en la Ultima Compra (Ordenes de Compra)
                    With objAyudaOrden
                        .CodProveedor = Trim(Txt_Prove.Text)
                        .CodigoProducto = Trim(rstResumenReq!CodProducto & "")
                        
                        strUltimaDescripcion = .obtenerUltimaDescripcionProductoDeProveedor
                    End With
                    'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                    With objAyudaVale
                        .CodigoProveedor = Trim(Txt_Prove.Text)
                        .CodigoProducto = Trim(rstResumenReq!CodProducto & "")
                        
                        .obtenerUltimoPrecioSinIgvProductoDeProveedor
                        
                        Select Case .CodigoMoneda
                            Case "S"
                                dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text))), "#0.0000")
                            Case Else
                                dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text))), "#0.0000")
                        End Select
                    End With
                    
                    If ModUtilitario.ObtenerCampoV2(cnDBTemp, "F3CODPRO", "TMPORDENDECOMPRA", "F3CODPRO", Trim(rstResumenReq!CodProducto & ""), "T", "AND COD_SOLICITUD = '" & Trim(rstResumenReq!NroPedido & "") & "'") = vbNullString Then
                        dblItem = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPORDENDECOMPRA", vbNullString, vbNullString, vbNullString, "TRIM(F3CODPRO & '') <> '' ORDER BY ITEM DESC") & "") + 1
                        
                        CadSql = vbNullString
                        CadSql = CadSql & "INSERT INTO TMPORDENDECOMPRA(ITEM, COD_SOLICITUD, CLIENTE, F3CODPRO, F5NOMPRO, F5NOMPRO_ING, F3CODMEDIDA, F3DESMEDIDA, "
                        CadSql = CadSql & "F5AFECTO, F3SINIGV, F3CANPRO, F3CANPROMAX, F3PORCDEMASIA, F3GASTO, F3CUENTA) "
                        CadSql = CadSql & "VALUES("
                        CadSql = CadSql & dblItem & ", "
                        CadSql = CadSql & "'" & Trim(rstResumenReq!NroPedido & "") & "', "
                        CadSql = CadSql & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_NOMREF", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(rstResumenReq!NroPedido & ""), "N") & "', "
                        CadSql = CadSql & "'" & Trim(rstResumenReq!CodProducto & "") & "', "
                        
                        CadSql = CadSql & "'" & IIf(strUltimaDescripcion <> vbNullString, strUltimaDescripcion, objAyudaBien.Descripcion) & "', "
                        
                        CadSql = CadSql & "'" & objAyudaBien.Descripcion & "', "
                        CadSql = CadSql & "'" & objAyudaBien.CodUM & "', "
                        CadSql = CadSql & "'" & Trim(rstResumenReq!um & "") & "', "
                        CadSql = CadSql & IIf(CBool(objAyudaBien.Afecto), "TRUE", "FALSE") & ", "
                                               
                        CadSql = CadSql & dblUltimoPrecioSinIGv & ", "
                        
                        CadSql = CadSql & Val(rstResumenReq!CANTIDADPC & "") & ", "
                        CadSql = CadSql & Val(rstResumenReq!CANTIDADPC & "") & ", "
                        CadSql = CadSql & objAyudaBien.PorcentajeDemasia & ", "
                        CadSql = CadSql & "'" & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T") & "', "
                        CadSql = CadSql & "'" & strCuentaContable & "')"
                    Else
                        CadSql = vbNullString
                        CadSql = CadSql & "UPDATE TMPORDENDECOMPRA "
                        CadSql = CadSql & "SET "
    '                    CadSql = CadSql & "F5NOMPRO = '" & IIf(strUltimaDescripcion <> vbNullString, strUltimaDescripcion, objAyudaBien.Descripcion) & "', "
    '                    CadSql = CadSql & "F5NOMPRO_ING = '" & objAyudaBien.Descripcion & "', "
    '                    CadSql = CadSql & "F3CODMEDIDA = '" & objAyudaBien.CodUM & "', "
    '                    CadSql = CadSql & "F3DESMEDIDA = '" & Trim(rstResumenReq!um & "") & "', "
                        CadSql = CadSql & "F3CANPRO = " & Val(rstResumenReq!CANTIDADPC & "") & ", "
                        CadSql = CadSql & "F3CANPROMAX = " & Val(rstResumenReq!CANTIDADPC & "") & " "
                        CadSql = CadSql & "WHERE "
                        CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstResumenReq!NroPedido & "") & "' AND "
                        CadSql = CadSql & "F3CODPRO = '" & Trim(rstResumenReq!CodProducto & "") & "'"
                    End If
                    
                    cnDBTemp.Execute CadSql
                End If
            End If
            
            rstResumenReq.MoveNext
        Loop
    End If
    
    If ModUtilitario.validarFormAbierto("frmUtilResumenRequerimiento") Then
        Unload frmUtilResumenRequerimiento
    End If
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTO"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTO"
    
    CadSql = vbNullString
    dblItem = 0
    strUltimaDescripcion = vbNullString
    dblUltimoPrecioSinIGv = 0
    
    If rstResumenReq.State = 1 Then rstResumenReq.Close
    
    Set rstResumenReq = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub recalcularItems()
'    Dim dblPorcentajeImpuesto As Double
'    Dim dblSignoImpuesto As Double
'    Dim dblCantidad As Double
'    Dim dblCantidadMaxima As Double
'    Dim dblPorcentajeDemasia As Double
'    Dim dblValorSinIGV As Double
'    Dim dblValorConIGV As Double
'    Dim dblPorcentajeDscto As Double
'    Dim dblValorDscto As Double
'    Dim bolAfecto As Boolean
'
'    Dim dblCantidadFinal As Double
'    Dim dblValorNetoSinIGV As Double
'    Dim dblBase As Double
'    Dim dblExonerado As Double
'    Dim dblImpuesto As Double
'    Dim dblTotal As Double
'
'    dblPorcentajeImpuesto = IIf(right(CmbTipDoc.Text, 2) = "02", gretenc, wwigv) / 100
'    dblSignoImpuesto = IIf(right(CmbTipDoc.Text, 2) = "02", -1, 1)
'
'    With Grid
'        If grid.Dataset.RecordCount = 0 Then Exit Sub
'
'        grid.Dataset.First
'
'        Do While Not grid.Dataset.EOF
'            'DATOS
'            dblCantidad = Val(grid.Dataset.FieldValues("F3CANPRO") & "")
'            dblCantidadMaxima = Val(grid.Dataset.FieldValues("F3CANPROMAX") & "")
'
'            If dblCantidadMaxima > 0 Then
'                If dblCantidad > dblCantidadMaxima Then
'                    dblCantidad = dblCantidadMaxima
'                End If
'            End If
'
'            dblPorcentajeDemasia = Val(grid.Dataset.FieldValues("F3PORCDEMASIA") & "") / 100
'
'            dblValorSinIGV = Val(grid.Dataset.FieldValues("F3SINIGV") & "")
'            dblValorConIGV = Val(grid.Dataset.FieldValues("F3CONIGV") & "")
'
'            dblPorcentajeDscto = Val(grid.Dataset.FieldValues("F3PORDESC") & "") / 100
'            dblValorDscto = Val(grid.Dataset.FieldValues("F3VALDESC") & "")
'
'            bolAfecto = CBool(grid.Dataset.FieldValues("F5AFECTO"))
'
'            'LIMPIAR VARIABLES
'            dblCantidadFinal = 0
'            dblValorNetoSinIGV = 0
'
'            dblBase = 0
'            dblExonerado = 0
'            dblImpuesto = 0
'            dblTotal = 0
'
'            'CALCULOS
'            If dblValorSinIGV <> 0 Or dblValorConIGV <> 0 Then
'                If dblValorSinIGV <> 0 Then
'                    dblValorConIGV = Val(Format((dblValorSinIGV * (1 + (dblPorcentajeImpuesto * dblSignoImpuesto))) * IIf(bolAfecto, 1, 0), "#0.000000"))
'                Else
'                    dblValorSinIGV = Val(Format(dblValorConIGV / (1 + (dblPorcentajeImpuesto * dblSignoImpuesto)), "#.000000"))
'                    dblValorConIGV = dblValorConIGV * IIf(bolAfecto, 1, 0)
'                End If
'
'                If dblPorcentajeDscto <> 0 Then
'                    dblValorDscto = Val(Format(dblValorSinIGV * dblPorcentajeDscto, "#0.000000"))
'                Else
'                    dblPorcentajeDscto = Val(Format((dblValorDscto) / dblValorSinIGV, "#0.00"))
'                End If
'            End If
'
'            'RESULTADOS
'            dblCantidadFinal = dblCantidad * (1 + dblPorcentajeDemasia)
'            dblValorNetoSinIGV = dblValorSinIGV - dblValorDscto
'
'            dblBase = (dblCantidadFinal * dblValorNetoSinIGV) * IIf(bolAfecto, 1, 0)
'            dblExonerado = (dblCantidadFinal * dblValorNetoSinIGV) * IIf(bolAfecto, 0, 1)
'            dblImpuesto = (dblBase * dblPorcentajeImpuesto) * IIf(bolAfecto, 1, 0)
'            dblTotal = dblBase + dblExonerado + (dblImpuesto * dblSignoImpuesto)
'
'            grid.Dataset.Edit
'
'            grid.Dataset.FieldValues("F3SINIGV") = dblValorSinIGV
'            grid.Dataset.FieldValues("F3CONIGV") = dblValorConIGV
'            grid.Dataset.FieldValues("F3PORDESC") = Val(Format(dblPorcentajeDscto * 100, "#0.00"))
'            grid.Dataset.FieldValues("F3CANPROFINAL") = dblCantidadFinal
'            grid.Dataset.FieldValues("F3VALDESC") = dblValorDscto
'
'            grid.Dataset.FieldValues("F3NETO") = dblValorNetoSinIGV
'
'            grid.Dataset.FieldValues("F3BASEIMP") = dblBase
'            grid.Dataset.FieldValues("F3MONINA") = dblExonerado
'            grid.Dataset.FieldValues("F3IGV") = dblImpuesto
'            grid.Dataset.FieldValues("F3TOTAL") = dblTotal
'
'            grid.Dataset.Post
'
'            grid.Dataset.Next
'        Loop
'            mostrarTotales
'    End With
    
    With objAyudaOrden
        If Grid.Dataset.RecordCount = 0 Then Exit Sub

        Grid.Dataset.First

        Do While Not Grid.Dataset.EOF
            .inicializarEntidadesDetalle
            
            'Entregar Datos a Clase
            .PorcentajeImpuesto = IIf(right(CmbTipDoc.Text, 2) = "02", gretenc, wwigv) / 100
            .SignoImpuesto = IIf(right(CmbTipDoc.Text, 2) = "02", -1, 1)
            
            .Cantidad = Val(Grid.Dataset.FieldValues("F3CANPRO") & "")
            .CantidadMaxima = Val(Grid.Dataset.FieldValues("F3CANPROMAX") & "")

            .PorcentajeDemasia = Val(Grid.Dataset.FieldValues("F3PORCDEMASIA") & "") / 100

            .PrecioSinImpuesto = Val(Grid.Dataset.FieldValues("F3SINIGV") & "")
            .PrecioConImpuesto = Val(Grid.Dataset.FieldValues("F3CONIGV") & "")

            .PorcentajeDscto = Val(Grid.Dataset.FieldValues("F3PORDESC") & "") / 100
            .TotalDscto = Val(Grid.Dataset.FieldValues("F3VALDESC") & "")

            .Afecto = CBool(Grid.Dataset.FieldValues("F5AFECTO"))
            
            'Calcular
            .calculosPorItem
            
            'Copiar Resultados
            Grid.Dataset.Edit

            Grid.Dataset.FieldValues("F3SINIGV") = .PrecioSinImpuesto
            Grid.Dataset.FieldValues("F3CONIGV") = .PrecioConImpuesto
            Grid.Dataset.FieldValues("F3PORDESC") = Val(Format(.PorcentajeDscto * 100, "#0.00"))
            Grid.Dataset.FieldValues("F3CANPROFINAL") = .CantidadFinal
            Grid.Dataset.FieldValues("F3VALDESC") = .TotalDscto
            
            Grid.Dataset.FieldValues("F3NETO") = .PrecioNetoSinImpuesto
            
            Grid.Dataset.FieldValues("F3BASEIMP") = .BasePorItem
            Grid.Dataset.FieldValues("F3MONINA") = .ExoneradoPorItem
            Grid.Dataset.FieldValues("F3IGV") = .ImpuestoPorItem
            Grid.Dataset.FieldValues("F3TOTAL") = .TotalPorItem
            
            Grid.Dataset.Post

            Grid.Dataset.Next
        Loop
            mostrarTotales
    End With
End Sub

Private Sub verificarCostoProductoPorCambioMoneda()
    Dim rstProducto As New ADODB.Recordset
    Dim dblUltimoPrecioSinIGv As Double
    
    Me.MousePointer = vbHourglass
    
'    If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
'        Grid.Dataset.Post
'    Else
'        Grid.Dataset.Edit
'
'        Grid.Dataset.Post
'    End If
    
    Grid.Dataset.Close
    
    'abrirCnTemporal
    
    'cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA WHERE TRIM(F3CODPRO & '') = ''"
    
    abrirCnTemporal
    
    CadSql = vbNullString
    CadSql = CadSql & "SELECT "
    CadSql = CadSql & "* "
    CadSql = CadSql & "FROM "
    CadSql = CadSql & "TMPORDENDECOMPRA "
    CadSql = CadSql & "WHERE "
    CadSql = CadSql & "TRIM(F3CODPRO & '') <> ''"
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    rstProducto.Open CadSql, cnDBTemp, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProducto.EOF Then
        rstProducto.MoveFirst
        
        Do While Not rstProducto.EOF
            'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
            With objAyudaVale
                .CodigoProveedor = Trim(Txt_Prove.Text)
                .CodigoProducto = Trim(rstProducto!F3CODPRO & "")
                
                .obtenerUltimoPrecioSinIgvProductoDeProveedor
                
                Select Case .CodigoMoneda
                    Case "S"
                        dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text))), "#0.0000")
                    Case Else
                        dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text))), "#0.0000")
                End Select
            End With
            
            CadSql = vbNullString
            CadSql = CadSql & "UPDATE TMPORDENDECOMPRA "
            CadSql = CadSql & "SET "
            CadSql = CadSql & "F3SINIGV = " & dblUltimoPrecioSinIGv & ", "
            CadSql = CadSql & "F3CONIGV = 0 "
            CadSql = CadSql & "WHERE "
            CadSql = CadSql & "TRIM(COD_SOLICITUD & '') = '" & Trim(rstProducto!COD_SOLICITUD & "") & "' AND "
            CadSql = CadSql & "F3CODPRO = '" & Trim(rstProducto!F3CODPRO & "") & "'"
            
            cnDBTemp.Execute CadSql
            
            rstProducto.MoveNext
        Loop
    'Else
    '    listarGrillaOrden
    End If
    
    CadSql = vbNullString
    dblUltimoPrecioSinIGv = 0
    
    If rstProducto.State = 1 Then rstProducto.Close
    
    Set rstProducto = Nothing
    
    listarGrillaOrden
    
    recalcularItems
    
    Me.MousePointer = vbDefault
End Sub

Rem SK ADD:----------------------------------------------------------------------------------------------------------

Private Sub abofechaentrega_Click()
    abofechaentrega.value = Now
End Sub

Private Sub abofechaentrega_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub abofechaentrega_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'        dxDBGrid1.Columns.FocusedIndex = 1
'    End If

End Sub

Private Sub atbmenu_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error Resume Next
    
    Dim resp    As Integer
    
    Select Case Tool.Id
        Case "ID_Nuevo":
'            Me.MousePointer = vbHourglass
'
'            inicio = True
'            Wnuevo = True
'
'            If swGrabacion = True Then
'                resp = MsgBox("La Orden no ha sido grabada. ¿Desea grabarla ahora?", vbYesNo + vbQuestion, "Sistema de Logística")
'                If resp = vbYes Then
'                    MODIFICAR_OC
'                End If
'            End If
'
'            sw_nuevo_documento = False
'            Limpia_Orden
'            limpiarCajas
'            'AdicionaItem
'            AdicionaItemGrid
'            sw_nuevo_documento = True
'            ModUtilitario.pulsarTecla vbKeyTab  'ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'            loc = 1
'
'            Me.MousePointer = vbDefault
'            'atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = False
'            atbmenu.Tools.ITEM("ID_Anular").Visible = False
'            atbmenu.Tools.ITEM("ID_Eliminar").Visible = False
'            atbmenu.Tools.ITEM("ID_Imprimir").Visible = False
''            atbmenu.Tools.ITEM("ID_Aprobacion").Visible = False
'
'            Rem SK ADD:
'            If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                Unload ayuda_productos
'            End If
            
            If MsgBox("¿Desea crear un nuevo registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            'strTipoOrden = vbNullString
            strNumeroOrden = vbNullString
            
            consultarOrden
            
            Me.MousePointer = vbDefault
        Case "ID_Grabar":
'            If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
'                Grid.Dataset.Post
'                sw_detalle = True
'            End If
'
'
'
'            If MsgBox("¿Desea Grabar la " & Me.Caption & "?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
'                Me.MousePointer = vbHourglass
'
'                GrabarOC
'                'ActualizarNumOrd
'
'                fraSeguimiento.Enabled = True
'
'                Me.MousePointer = vbDefault
'            End If
            
            Me.MousePointer = vbHourglass
            
            validarCajas
            
            Me.MousePointer = vbDefault
        Case "ID_Eliminar"
'            Dim strReq As String
'            If MsgBox("¿Desea Eliminar la Orden de Compra?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
'                eliminar_sin_preguntar
'
'                strReq = ""
'                For i = 1 To Grid.Dataset.RecordCount
'                    Grid.Dataset.RecNo = i
'                    If strReq <> Grid.Columns.ColumnByFieldName("cod_solicitud").value & "" Then
'                        strReq = Grid.Columns.ColumnByFieldName("cod_solicitud").value & ""
'                        VerificaAtencionDeRequerimiento (strReq)
'                    End If
'
'                Next
'                csql = "delete * from if4orden where f4local='" & TOC & "' and f4numord='" & Txt_NumOC.Text & "'"
'                cnn_dbbancos.Execute csql
'                'AlmacenaQuery_sql csql, cnn_dbbancos
'                Actualiza_Log csql, cnn_dbbancos.ConnectionString
'                Me.Hide
'                lista_oc.dxDBGrid1.Dataset.ADODataset.Requery
'            End If
            
            Me.MousePointer = vbHourglass
            
            eliminarOrden
            
            Me.MousePointer = vbDefault
        Case "ID_Anular":
'            If Trim$(Txt_NumOC.Text) = "" Then
'                MsgBox "No existe Orden de Compra", vbInformation, "Sistema de Logística"
'                Exit Sub
'            Else
'                eliminar
'            End If
            
            Me.MousePointer = vbHourglass
            
            anularOrden
            
            Me.MousePointer = vbDefault
        Case "ID_Cerrar"
            'MsgBox "Opción no Disponible por el momento.", vbInformation + vbOKOnly, App.ProductName
            
            Me.MousePointer = vbHourglass
            
            cerrarOrden
            
            Me.MousePointer = vbDefault
        Case "ID_Imprimir":
'            If Len(Trim(Txt_NumOC.Text)) > 0 Then
'                Select Case wrucempresa
'                    Case "20381208835"
'                        Imprime_Orden_Electrica 1
'                    Case "20434047171"
'                        Imprime_Orden_Ditec 1
'                    Case Else
'                        Imprime_Orden 1
'                End Select
'            Else
'                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
'            End If
            
            Me.MousePointer = vbHourglass
            
            If MsgBox("Guarde los cambios antes de imprimir." & vbNewLine & _
                        "¿Desea imprimir la orden?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                
                'Imprime_Orden 1
                imprimeOrdenV2 strTipoOrden, strNumeroOrden
                
            End If
            
            Me.MousePointer = vbDefault
        Case "ID_Email":
            Me.MousePointer = vbHourglass
            
            enviarViaMailOrden
            
            Me.MousePointer = vbDefault
        Case "ID_RenovarOrden"
'            'MsgBox "renovar"
'            Me.MousePointer = vbHourglass
'            If Rs.State = 1 Then Rs.Close
'            sql = "select * from if4orden where f4numord='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
'            Rs.Open sql, cnn_dbbancos, 3, 1
'            If Rs.RecordCount > 0 Then
'                If Rs!F4ESTNUL = "S" Then
'                    MsgBox "No se puede renovar una Orden de Compra Anulada", vbExclamation, "Sistema de Logística"
'                    SwRenovar = False
'                    Exit Sub
'                End If
'            Else
'                MsgBox "Error, no existe la orden que quiere renovar", vbCritical, wnomcia
'            End If
'            SwRenovar = True
'            If MsgBox("¿Desea Renovar la Orden de Compra?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
'                Me.MousePointer = vbHourglass
'                GrabarOC
'                'ActualizarNumOrd
'                'ANULANDO ANTERIOR
'                If Trim$(Txt_NumOC.Text) = "" Then
'                    MsgBox "No existe Orden de Compra", vbInformation, "wnomcia"
'                    Exit Sub
'                Else
'                    eliminar_sin_preguntar
'                End If
'                SwRenovar = False
'                Txt_TOC.Text = TOC
'                Me.Txt_NumOC.Text = wNumOc
'                Call Txt_NumOC_KeyPress(13)
'
'            End If
'            Me.MousePointer = vbDefault
        Case "ID_CtasxPagar"
'            If Len(Trim(Txt_NumOC.Text)) > 0 Then
'                If rsif4orden.State = adStateOpen Then rsif4orden.Close
'                rsif4orden.Open "SELECT F4CORRELA FROM IF4ORDEN WHERE F4NUMORD=" & Txt_NumOC.Text & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                If Not rsif4orden.EOF Then
'                    If Val("" & rsif4orden.Fields("F4CORRELA")) > 0 Then
'                        MsgBox "La orden de compra ya fue trasladada a cuentas por pagar.", vbInformation, "Atención"
'                    Else
'                        If MsgBox("Está seguro(a) de trasladar la Orden de Compra a Cuentas por Pagar ?", vbYesNo, "Atención") = vbYes Then
'                            TRASLADA_CTASXPAGAR Txt_NumOC.Text
'                        End If
'                    End If
'                End If
'                rsif4orden.Close
'            Else
'                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
'            End If
        Case "ID_Aprobacion"
'            If MsgBox("¿Desea solicitar la aprobación de la Orden de Compra " & Txt_NumOC.Text & "?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
'                Call EnviaMail(Txt_NumOC.Text)
'            End If
        Case "Reposicion"
            
        Case "ID_ModificarOrden"
            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0016'") = vbNullString Then
                MsgBox "Ud. no cuenta con permisos para usar esta Opción.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If MsgBox("¿Desea guardar las modificaciones realizadas sobre la Orden?", vbInformation + vbYesNo, App.ProductName) = vbYes Then
                atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
                atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
            Else
                atbmenu.Tools.ITEM("ID_Grabar").Enabled = False
                atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = True
            End If
        Case "ID_Salir":
            Unload Me
    End Select
End Sub
Private Sub ActualizarNumOrd(psolicitud As String)
    Dim prSol As New Recordset
    Dim cadena As String
    
    sql = "Select * from TB_CABSOLICITUD where Cod_Solicitud='" & Trim(psolicitud) & "'"
    
    prSol.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
    
    If Not (prSol.EOF Or prSol.Bof) Then
        If Trim(prSol!NUMORDEN) <> "" And Trim(prSol!NUMORDEN) <> Trim(ordendecompra.Txt_NumOC.Text) Then
            cadena = "" & prSol!NUMORDEN & " , " & Trim(ordendecompra.Txt_NumOC.Text)
        Else
            cadena = Trim(ordendecompra.Txt_NumOC.Text)
        End If
        sql = "Update TB_CABSOLICITUD set NumOrden='" & left(Trim(cadena), 255) & "' where Cod_Solicitud='" & Trim(psolicitud) & "'"
        cnn_dbbancos.Execute sql
        'AlmacenaQuery_sql sql, cnn_dbbancos
        Actualiza_Log sql, cnn_dbbancos.ConnectionString
    End If
    
    prSol.Close
End Sub
Private Sub cmbtipopera_Change()
    
    
    wgraba = 0
    If Not inicio Then swGrabacion = True
    
End Sub

'Private Sub cmbtipopera_Click()
'
'    If cmbtipopera.ListIndex = 1 Then
'        dxDBGrid1.Columns.ColumnByFieldName("check").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3PORDCT").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("ds_unidmed").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3TOTDCT").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3IGV").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("CANT_ANT").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Visible = True
'    Else
'        dxDBGrid1.Columns.ColumnByFieldName("check").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("F3PREUNI").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3PORDCT").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("f3medida").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("F3TOTDCT").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("F5AFECTO").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("F3IGV").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("CANT_ANT").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Visible = True
'        dxDBGrid1.Columns.ColumnByFieldName("F3PRECOS").Caption = "Costo Unit."
'        If wf1visualiza_dctos = "*" Then
'            dxDBGrid1.Columns.ColumnByFieldName("f3pordct").Visible = False
'            dxDBGrid1.Columns.ColumnByFieldName("f3totdct").Visible = False
'            dxDBGrid1.Columns.ColumnByFieldName("f5valvta").Visible = False
'        End If
'    End If
'
'End Sub

'Private Sub cmbtipopera_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        If txtcodcosto.Visible = True Then
'            txtcodcosto.SetFocus
'        Else
'            If txtuupp.Visible = True Then
'                txtuupp.SetFocus
'            Else
'                dxDBGrid1.SetFocus
'            End If
'        End If
'    End If
'
'End Sub

'Private Sub cmbtipopera_LostFocus()
'
'    If cmbtipopera.ListIndex = 1 Then
'        Forma_Imp
'    Else
'        Forma_Loc
'    End If
'
'End Sub

Sub Forma_Loc()
    
    Visib
    
End Sub

Sub Visib()
    
    Label9.Visible = True
    Label10.Visible = True
    'Label11.Visible = True
    lblmoneda(0).Visible = True
    lblmoneda(1).Visible = True
    lblmoneda(2).Visible = True
    txtmonto.Visible = True
    txtbase.Visible = True
    txtigv.Visible = True
    
End Sub

Sub Invisi()

    Cmbmone.ListIndex = 1
    Label9.Visible = False
    Label10.Visible = False
    'Label11.Visible = False
    Label12.left = 5000
    lblmoneda(0).Visible = False
    lblmoneda(1).Visible = False
    lblmoneda(2).left = 5600
    txtmonto.Visible = False
    txtbase.Visible = False
    txtigv.Visible = False
    txttotal.left = 4905
    
End Sub

Sub Forma_Imp()

    Invisi
    
End Sub

Private Sub cmdcerrar_Click()
'pnlcosto.Visible = False
End Sub

Private Sub chkDetraccionAplicar_Click()
    If CBool(chkDetraccionAplicar.value) Then
        txtDetraccionPorc.Text = Format(IIf(Val(txtDetraccionPorc.Text) > 0, Val(txtDetraccionPorc.Text), infoPlusParConta.dblDetraccionPorcentaje), "#0.00")
    Else
        txtDetraccionPorc.Text = "0.00"
    End If
End Sub

Private Sub chkOrdenEnviada_Click()
    txtEnviadoPor.Enabled = CBool(chkOrdenEnviada.value)
    txtEnviadoPor.BackColor = IIf(CBool(chkOrdenEnviada.value), HA, DH)
    dtpFechaEnvio.Enabled = CBool(chkOrdenEnviada.value)
End Sub

Private Sub chkOrdenRecepcionada_Click()
    txtRecepcionadoPor.Enabled = CBool(chkOrdenRecepcionada.value)
    txtRecepcionadoPor.BackColor = IIf(CBool(chkOrdenRecepcionada.value), HA, DH)
    dtpFechaRecepcion.Enabled = CBool(chkOrdenRecepcionada.value)
End Sub

Private Sub chkSinProveedorEsp_Click()
    Txt_Prove.Enabled = Not CBool(chkSinProveedorEsp.value)
    txtcontacto.Enabled = Not CBool(chkSinProveedorEsp.value)
    
    If CBool(chkSinProveedorEsp.value) Then
        Txt_Prove.Text = ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "CodigoProveedorComprasVarias", "l")
        
        Txt_Prove.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2CODPROV", Trim(Txt_Prove.Text), "T")
    Else
        Txt_Prove.Text = vbNullString
        pnlnomprv.Caption = vbNullString
        pnldireprv.Caption = vbNullString
        txtcontacto.Text = vbNullString
    End If
    
    Txt_Prove_LostFocus
End Sub

Private Sub chkSinProveedorEsp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmbTipDoc_Click()
    With Grid.Columns
        If right(CmbTipDoc.Text, 2) = "02" Then
            .ColumnByFieldName("F3SINIGV").Caption = "Precio s/Ret"
            .ColumnByFieldName("F3CONIGV").Caption = "Precio c/Ret"
            .ColumnByFieldName("F3IGV").Caption = "Retención"
            
            lblImpuesto.Caption = "Retención"
        Else
            .ColumnByFieldName("F3SINIGV").Caption = "Precio s/IGV"
            .ColumnByFieldName("F3CONIGV").Caption = "Precio c/IGV"
            .ColumnByFieldName("F3IGV").Caption = "I.G.V."
            '.ColumnByFieldName("F3NETO").Caption = "Precio Neto"
            lblImpuesto.Caption = "I.G.V."
        End If
    End With
    
    recalcularItems
'            Dim nPorc As Double
'            If right(CmbTipDoc.Text, 2) = "02" Then
'                nPorc = gretenc
'                If Grid.Columns.ColumnByFieldName("f3sinigv").Value >= 700 Then
'                    Grid.Dataset.Edit
'                    Grid.Columns.ColumnByFieldName("F5afecto").Value = True
'                    Grid.Dataset.Post
'                Else
'                    Grid.Dataset.Edit
'                    Grid.Columns.ColumnByFieldName("F5afecto").Value = False
'                    Grid.Dataset.Post
'                End If
'
'            Else
'                nPorc = wwigv
'            End If
'            Grid.Dataset.Edit
'            'Grid.Columns.ColumnByFieldName("F3COLMOD").Value = UCase(Grid.Columns.FocusedColumn.FieldName)
'            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
'                'Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value * (1 + (nPorc / 100))
'                If right(CmbTipDoc.Text, 2) = "02" Then
'                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
'                Else
'                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value * (1# + (nPorc / 100))
'                End If
'                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
'
'            Else
'                Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
'
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
'                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
'            End If
'            Grid.Dataset.Post
'            CalculaTotal
End Sub

Private Sub CmbTipDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub cmdFirmaAprob_Click()
    
'''    frmaccesocompras.Show 1
'''    txtTempo = wusuario
'''    If txtTempo <> "" Then
'''        If xcentro = "08" Then
'''            If txtAprobadoX <> "" Then MsgBox "La orden ya fue firmada por el Jefe de Compras", 48, "Sistema de Logística"
'''            txtAprobadoX = txtTempo
'''            FileCopy Devuelve_Path("BMP") & wusuario & ".bmp", Devuelve_Path("") & "firma.bmp"
'''            ImgAprobadoX.Picture = LoadPicture(Devuelve_Path("") & "firma.bmp")
'''        Else
'''            If txtAprobadoX = "" Then MsgBox "Debe firmar antes El Jefe de Compras!!!", 48, "Sistema de Logística": Exit Sub
'''            If txtAprobadoY <> "" Then MsgBox "Ya firmó el Jefe del Area!!!", 48, "Sistema de Logística": Exit Sub
'''            If txtAprobadoY = txtTempo Then
'''                FileCopy Devuelve_Path("BMP") & wusuario & ".bmp", Devuelve_Path("") & "firma.bmp"
'''                ImgAprobadoY.Picture = LoadPicture(Devuelve_Path("") & "firma.bmp")
'''            Else
'''                If txtAprobadoX = "" Then MsgBox "Debe firmar antes El Jefe de Compras!!!", 48, "Sistema de Logística": Exit Sub
'''                If txtAprobadoY = "" Then MsgBox "Debe firmar antes el Jefe del Area!!!", 48, "Sistema de Logística": Exit Sub
'''                If txtAprobadoZ <> "" Then MsgBox "Ya firmó el Gerente de Logística!!!", 48, "Sistema de Logística": Exit Sub
'''                txtAprobadoZ = txtTempo
'''                FileCopy Devuelve_Path("BMP") & wusuario & ".bmp", Devuelve_Path("") & "firma.bmp"
'''                ImgAprobadoY.Picture = LoadPicture(Devuelve_Path("") & "firma.bmp")
'''            End If
'''        End If
'''    Else
'''        Exit Sub
'''    End If

End Sub

Private Sub Cmbmone_Click()
    Select Case Cmbmone.ListIndex
        Case 0:
            lblmoneda(0).Caption = "S/" '"S/"
            lblmoneda(1).Caption = "S/" '"S/"
            lblmoneda(2).Caption = "S/" '"S/"
            lblmoneda(3).Caption = "S/" '"S/"
            lblmoneda(4).Caption = "S/" '"S/"
            Me.txttotal.BackColor = &HC0FFFF
            Me.txtigv.BackColor = &HC0FFFF
            Me.txtbase.BackColor = &HC0FFFF
            Me.txtmonto.BackColor = &HC0FFFF
            Me.TxtRnd.BackColor = &HC0FFFF
            
            Me.txtDetraccionPorc.BackColor = &HC0FFFF
        Case 1:
            lblmoneda(0).Caption = "US$"
            lblmoneda(1).Caption = "US$"
            lblmoneda(2).Caption = "US$"
            lblmoneda(3).Caption = "US$"
            lblmoneda(4).Caption = "US$"
            Me.txttotal.BackColor = &HC0FFC0
            Me.txtigv.BackColor = &HC0FFC0
            Me.txtbase.BackColor = &HC0FFC0
            Me.txtmonto.BackColor = &HC0FFC0
            Me.TxtRnd.BackColor = &HC0FFC0
            
            Me.txtDetraccionPorc.BackColor = &HC0FFC0
    End Select
    If Not inicio Then swGrabacion = True
    Grid.Dataset.Refresh
    
    If Trim(Txt_Prove.Text) <> vbNullString Then
        If Grid.Dataset.RecordCount > 0 Then
            verificarCostoProductoPorCambioMoneda
        End If
    End If
End Sub

Private Sub Cmbmone_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

Private Sub calcula()

End Sub

Private Sub mostrarTotales()
    With Grid.Columns
        txtbase.Text = Format(.ColumnByFieldName("F3BASEIMP").SummaryFooterValue, "#,0.00")
        txtmonto.Text = Format(.ColumnByFieldName("F3MONINA").SummaryFooterValue, "#,0.00")
        txtigv.Text = Format(.ColumnByFieldName("F3IGV").SummaryFooterValue, "#,0.00")
        txttotal.Text = Format(.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "#,0.00")
    End With
End Sub

Private Sub dxCheckBox1_Click()
    If dxCheckBox1.Checked Then
        Grid.Columns.ColumnByFieldName("F3GASTO").Visible = True
        Grid.Columns.ColumnByFieldName("F3CUENTA").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("F3GASTO").Visible = False
        Grid.Columns.ColumnByFieldName("F3CUENTA").Visible = False
    End If
End Sub

Private Sub dxCheckBox2_Click()
    If dxCheckBox2.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f3sinigv").DecimalPlaces = 6
    Else
        Grid.Columns.ColumnByFieldName("f3sinigv").DecimalPlaces = 2
    End If
End Sub

Private Sub dxCheckBox3_Click()
    If dxCheckBox3.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f3baseimp").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("f3baseimp").Visible = False
    End If
End Sub

Private Sub dxCheckBox4_Click()
    If dxCheckBox4.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("F3PORCDEMASIA").Visible = True
        Grid.Columns.ColumnByFieldName("F3CANPROFINAL").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("F3PORCDEMASIA").Visible = False
        Grid.Columns.ColumnByFieldName("F3CANPROFINAL").Visible = False
    End If
End Sub


Private Sub dxCheckBox5_Click()
    If dxCheckBox5.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f3pordesc").Visible = True
        Grid.Columns.ColumnByFieldName("f3valdesc").Visible = True
        Grid.Columns.ColumnByName("colAyudaPorcDscto").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("f3pordesc").Visible = False
        Grid.Columns.ColumnByFieldName("f3valdesc").Visible = False
        Grid.Columns.ColumnByName("colAyudaPorcDscto").Visible = False
    End If
End Sub

Private Sub dxCheckBox6_Click()
    If dxCheckBox6.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f3observa").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("f3observa").Visible = False
    End If
End Sub

Private Sub dxCheckBox7_Click()
    If dxCheckBox7.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f5nompro_ing").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("f5nompro_ing").Visible = False
    End If
End Sub

Private Sub dxCheckBox8_Click()
    If dxCheckBox8.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("CLIENTE").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("CLIENTE").Visible = False
    End If
End Sub


Private Sub dxCheckBox9_Click()
    If dxCheckBox9.Checked = 1 Then
        If MsgBox("¿Dese efectuar el Cierre Parcial de Orden?" & vbNewLine & vbNewLine & _
                    "ATENCIÓN: Usar el Cierre Parcial de Orden solo en caso se desestime la entrega, en coordinación con el proveedor, de algunos Items y necesariamente otros queden Pendiente de Entrega; caso contrario use el Cierre Total de la Orden, ya que de esto depende la debida actualización del Estado de la Orden.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            
            Me.MousePointer = vbHourglass
            

           actualizarSaldosPorEntregarDeProductos
            listarGrillaOrden
            
            atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
            
            Grid.Columns.ColumnByFieldName("PORENTREGAR").Visible = True
            Grid.Columns.ColumnByFieldName("CERRAR").Visible = True
            
            Me.MousePointer = vbDefault
        Else
            atbmenu.Tools.ITEM("ID_Cerrar").Enabled = True
            
            Grid.Columns.ColumnByFieldName("PORENTREGAR").Visible = False
            Grid.Columns.ColumnByFieldName("CERRAR").Visible = False
        End If
    Else
        atbmenu.Tools.ITEM("ID_Cerrar").Enabled = True
        
        Grid.Columns.ColumnByFieldName("PORENTREGAR").Visible = False
        Grid.Columns.ColumnByFieldName("CERRAR").Visible = False
    End If
End Sub

'Private Sub cmdopera_Click(Index As Integer)
'If Index = 0 Then
'Me.MousePointer = vbHourglass
'Rem NSE IMPRIMIR
'If Len(Trim(Txt_NumOC.Text)) > 0 Then
'    If optcosto(0).Value Then
'        Imprime_Orden 1
'    Else
'        Imprime_Orden 2
'    End If
'Else
'    MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
'End If
'Me.MousePointer = vbDefault
'Else
'    pnlcosto.Visible = False
'End If
'End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case Column.Caption
   Case "Codigo", "Fab.", "Descripción", "U.M.", "Cant": Text = Format(Text, "#,###,###0.0000")
   Case "Costo Unit.": Text = Format(Text, "#,###,###0.0000")
   Case "Costo Unit.": Text = Format(Text, "#,###,###0.0000")
    Case "Costo Unit.": Text = Format(Text, "#,###,###0.0000")
End Select
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Dim sql         As String
    If KeyCode = 113 Then
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3codpro", "f5codfab":
        
            If pnlnomprv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
        
            wcodproducto = ""
            wrucprov = Trim(Txt_Prove.Text)
            wnomprov = Trim(pnlnomprv.Caption)
            Con_Ayu = 3
            ayuda_productos.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").value = wdesproducto
              '  dxDBGrid1.Columns.ColumnByFieldName("f5c").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").value = wcodfab
                
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").value = wmarca
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").value = wafecto

                Rem EMB dxDBGrid1.Dataset.FieldValues("f5valvta") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.value, "dd/mm/yyyy"))
                dxDBGrid1.Columns.FocusedIndex = 5
            End If
    End Select

    End If
    
    If KeyCode = 115 Or KeyCode = 46 Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
            sw_nuevo_item = False
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Desea salir del registro?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
        Cancel = 1
    Else
        Rem SK ADD:
        If ModUtilitario.validarFormAbierto("ayuda_productos") Then
            Unload ayuda_productos
        End If
        
        ModUtilitario.sWrtIni strFichero, "ConfigCP", "OrdenCompraAbierta", "0"
    End If
End Sub

'Private Sub dxDBGrid1_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal y As Single)
'
'    If dxDBGrid1.Columns.FocusedIndex = 1 Then
'        If Len(Trim("" & dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value)) > 0 Then
'            lbldescripcion.Visible = True
'            lbldescripcion.Caption = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value)
'        Else
'            lbldescripcion.Caption = ""
'            lbldescripcion.Visible = False
'        End If
'    Else
'        lbldescripcion.Caption = ""
'        lbldescripcion.Visible = False
'    End If
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim strReq As String
'
'    strReq = ""
'    For i = 1 To Grid.Dataset.RecordCount
'        Grid.Dataset.RecNo = i
'        If strReq <> Grid.Columns.ColumnByFieldName("cod_solicitud").Value & "" Then
'            strReq = Grid.Columns.ColumnByFieldName("cod_solicitud").Value & ""
'            '******************* NUEVO CAMBIO ****************
'            'VerificaAtencionDeRequerimiento (strReq)
'            '*************************************************
'        End If
'
'    Next


    sw_nuevo_item = True
    dxDBGrid1.Dataset.Close
    Grid.Dataset.Close
'    lista_oc.dxDBGrid1.Dataset.Active = False
'    lista_oc.dxDBGrid1.Dataset.Refresh
'    lista_oc.dxDBGrid1.Dataset.Active = True
    
    
    With frmListaOrden
        .listarOrden
        
        .Show
    End With
End Sub
Private Sub DesHabilitar()
        atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
        atbmenu.Tools.ITEM("ID_Email").Enabled = True
        atbmenu.Tools.ITEM("ID_Eliminar").Enabled = True
End Sub
Private Sub define_cabecera()

    lblmoneda(0).left = 8580
    Label9.left = 7575
    txtbase.left = 7350
    
End Sub

Private Sub Form_Load()
'    'Dim fec     As Date
'
'    SwRenovar = False
'    Me.MousePointer = vbHourglass
'
'    num_solcomp = vbNullString
'
''    If wTipoOC = 1 Or TOC = "OC" Then
''        Me.Caption = "Orden de Compra"
''    Else
''        Me.Caption = "Orden de Servicio"
''    End If
'
'    Set rst = New ADODB.Recordset
'    Set rsOrdenCab = New ADODB.Recordset
'    Set rsOrdenDet = New ADODB.Recordset
'    Set rsproductos = New ADODB.Recordset
'    Set rssolcab = New ADODB.Recordset
'    Set rsSolDet = New ADODB.Recordset
'    Set rstaux = New ADODB.Recordset
'
'    sw_ayuda = False
'    inicio = True
'    swGrabacion = False
'    sw_activate = False
'
'    CargaDocumentos CmbTipDoc
'
'    Rem SK ADD: Deshabilitado condicional, innecesaria.
'    loc = 1
'
''    If loc = 2 Then
''        Call define_cabecera
''        txtmonto.Visible = False
''        txtigv.Visible = False
''        txttotal.Visible = False
''        Label10.Visible = False
''        Label11.Visible = False
''        Label12.Visible = False
''        lblmoneda(0).Visible = False
''        lblmoneda(1).Visible = False
''        lblmoneda(2).Visible = False
''    Else
''        loc = 1
''    End If
'
'    'txt_fecha.Value = Format(Date, "dd/MM/yyyy")
'    'fec = txt_fecha.Value
'    Wnuevo = True
'    flawigv = False
'    SWcondipago = 0
'
'    Set rst = Af.OpenSQLForwardOnly("select F1IGV, F1RETENC from param_com where f1codemp='" & UCase(wempresa) & "'", cconex_ctrcom)
'
'    If Not (rst.EOF) Then
'         wwigv = rst.Fields("F1IGV")
'         gretenc = rst.Fields("F1RETENC")
'    End If
'
'    rst.Close
'
'    'Txt_Prove.Enabled = True
'
'    If FlagGeneraOC = False Then
'        Wnuevo = True
'    End If
'
'    jc = 0
'
'    sw_nuevo_item = False
'
'    'If Dir(wrutatemp & "tmp_logistica_" & wempresa & ".mdb") <> "" Then
'    '    cnombase = "tmp_logistica_" & wempresa & ".mdb"
'    'Else
'        cnombase = "templus.mdb" '"tmp_bancos.MDB"
'    'End If
'
'    cnomtabla = "TMPORDENDECOMPRA"
'
''    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
''    If cnn_form.State = adStateOpen Then cnn_form.Close
''    cnn_form.Open cconex_form
'
'    abrirCnTemporal
'
'    StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ";Persist Security Info=False"
'
'    If CnTmp.State = 1 Then CnTmp.Close
'
'    CnTmp.Open StrCn
'
'    configurarGrilla
'
'    'limpiarCajas
'
'    'If sw_nuevo_documento = True Then
'    '    sw_nuevo_documento = False
'    '
'    '    AdicionaItemGrid
'    '
'    '    sw_nuevo_documento = True
'    'Else
'    '    inicio = True
'    '
'        MODIFICAR_OC
'    '
'    '    sw_nuevo_documento = False
'    '    inicio = False
'    'End If
'
'    atbmenu.Tools.ITEM("ID_Grabar").Visible = True
'
'    'importarDatosRequerimiento
'
'    Screen.MousePointer = vbDefault
'    Me.MousePointer = vbDefault
    
    Me.MousePointer = vbHourglass
    
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPORDENDECOMPRA"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTO"
    
    abrirCnTemporal
    
    cnDBTemp.Execute "DELETE FROM TMPUTILRESUMENREQUERIMIENTO"
    
    abrirCnContaTabla
    
    configuraGrilla
    
    listarEstadoEnImageCombo
    
    listarTipoComprobanteEnCombo
    
    consultarOrden
    
    'Activar Control de Apertura de Formulario
    '(Para evitar abrir mas de una vez, el mismo formulario en diferentes Instancias del Programa)
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "OrdenCompraAbierta", "1"
    
    Me.MousePointer = vbDefault
End Sub

Private Sub CargaDocumentos(pCombo As ComboBox)
    Dim TbDocumento1 As New ADODB.Recordset
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT * FROM DOCUMENTOS WHERE F2TIPO IN ('P', 'A') ORDER BY F2DESDOC"
    
    If TbDocumento1.State = adStateOpen Then TbDocumento1.Close
    
    TbDocumento1.Open SqlCad, cconex_dbbancos, adOpenDynamic, adLockOptimistic
    
    pCombo.Clear
    
    If Not TbDocumento1.EOF Then
        Do While Not TbDocumento1.EOF
            pCombo.AddItem TbDocumento1.Fields("F2DESDOC") + Space(100) + TbDocumento1.Fields("F2CODDOC")
            
            TbDocumento1.MoveNext
        Loop
    End If
    
    pCombo.ListIndex = 0
    
    TbDocumento1.Close
    
    Set TbDocumento1 = Nothing
End Sub


Sub limpiarCajas()
    'SWcondipago = 0
    
    Rem SK ADD:
    Select Case strTipoOrden
        Case "OC"
            Me.Caption = "Orden de Compra"
        Case "OS"
            Me.Caption = "Orden de Servicio"
    End Select
    
    Txt_TOC.Text = strTipoOrden
    Txt_NumOC.Text = vbNullString
    
    'Txt_NumSolComp.Text = vbNullString
    
    txt_fecha.value = Format(Date, "dd/mm/yyyy")
    'abofechaentrega.CheckBox = True
    abofechaentrega.value = Format(Date, "dd/mm/yyyy")
    
    'aBoHoraEntrega.Value = Time
    
    'FrameOC.Caption = ""
    txtcontacto.Text = vbNullString
    txtcodsoli.Text = wusuario
        pnlnomsoli.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", wusuario, "T")
    Cmbmone.ListIndex = 0
    TxtCodCosto.Text = vbNullString
    PnlNomCosto.Caption = vbNullString
    txtcodforma.Text = vbNullString
    pnlnomforma.Caption = vbNullString
           
    'txt_tc.Text = Format(traerCampo("CAMBIOS", "CAMBIO", "FECHA", Me.txt_fecha.Value, ""), "0.000")
    txt_tc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", Trim(txt_fecha.value), "F")), "0.000")
    
    Txt_Referencia.Text = vbNullString
    
    txtbase.Text = "0.00"
    txtmonto.Text = "0.00"
    txtigv.Text = "0.00"
    txttotal = "0.00"
      
    'SWcondipago = 0
    txtempresa.Text = UCase(wnomcia)
    
    txtCotizacion.Text = vbNullString
    txtlugar_entrega.Text = wdireccion
    
    
    Rem SK ADD:
    fraSeguimiento.Enabled = False
    chkOrdenEnviada.value = vbUnchecked
    txtEnviadoPor.Text = vbNullString: txtEnviadoPor.BackColor = DH
    dtpFechaEnvio.value = Date
    
    chkOrdenRecepcionada.value = vbUnchecked
    txtRecepcionadoPor.Text = vbNullString: txtRecepcionadoPor.BackColor = DH
    dtpFechaRecepcion.value = Date
End Sub

Private Sub Limpia_Orden()

'    pnlnomcosto.Caption = ""
    Txt_Prove.Text = ""
    pnlnomprv.Caption = ""
    txtcontacto.Text = ""
    txtcodsoli.Text = ""
    Txt_NumSolComp.Text = ""
    pnlnomsoli.Caption = ""
    txtcodforma.Text = ""
    pnlnomforma.Caption = ""
'    txtcodcosto.Text = ""
    pnldireprv.Caption = ""
    Txt_Referencia.Text = ""
    txtobserva.Text = ""
    Txt_NumOC = ""
    
    txt_tc.Text = Format(traerCampo("CAMBIOS", "CAMBIO", "FECHA", Me.txt_fecha.value, ""), "0.000")
    txttotal.Text = "0.00"
    txtigv.Text = "0.00"
    txtbase.Text = "0.00"
    txtmonto.Text = "0.00"
    
    wgraba = 1
    
    
End Sub

Sub Visi()

    txtbase.Visible = True
    txtigv.Visible = True
    txttotal.Visible = True

End Sub

'Sub LLENA_TEMPCAB()
'Dim cnn         As ADODB.Connection
'Dim tempocompra As ADODB.Recordset
'Dim X           As Integer
'Dim rsprod      As New ADODB.Recordset
'
'    'Nueva Versión
'    Set cnn = New ADODB.Connection
'    Set tempocompra = New ADODB.Recordset
'    If Dir(wrutatemp & "tmp_logistica_" & wempresa & ".mdb") <> "" Then
'        cnombase = "tmp_logistica_" & wempresa & ".mdb"
'    Else
'        cnombase = "templus.mdb" '"tmp_bancos.MDB"
'    End If
'
'    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ";Persist Security Info=False"
'
'    sql = "delete * from tmpocompra"
'    cnn.Execute sql
'    ''AlmacenaQuery_sql sql, cnn
'
'    If tempocompra.State = adStateOpen Then tempocompra.Close
'    tempocompra.Open "tmpocompra", cnn, adOpenStatic, adLockOptimistic
'
'    With dxDBGrid1
'        If .Dataset.RecordCount = 0 Then
'            tempocompra.Close
'            cnn.Close
'            Exit Sub
'        End If
'        .Dataset.First
'        If Not (.Dataset.EOF) Then
'            .Dataset.First
'            Do While Not (.Dataset.EOF)
'                If Val(IIf(IsNull(.Dataset.FieldValues("f3precos")), 0, _
'                .Dataset.FieldValues("f3precos"))) > 0 Then
'                    tempocompra.AddNew
'                    tempocompra!Orden = Format(Txt_NumOC.Text, "0000000")
'                    tempocompra!PROVEEDOR = pnlnomprv.Caption
'                    tempocompra!direccion = pnldireprv.Caption
'                    tempocompra!ruc = Txt_Prove.Text
'                    tempocompra!CLIENTE = txtFechaPago.Text
'                    tempocompra!CODCONTA = IIf(ChK_regularizacion.Checked = True, 1, 0)
'                    'tempocompra!CONTACTO = txtcontacto.Text
'                    tempocompra!Fecha = txt_fecha.Value
'                    tempocompra!FORPAG = pnlnomforma.Caption
'                    tempocompra!Moneda = Cmbmone.Text
'                    tempocompra!referencia = Txt_Referencia.Text
'                    'tempocompra!Centro = txtcodcosto.Text
'                    'tempocompra!nomcentro = pnlnomcosto.Caption
'                    tempocompra!OBSERVA = txtobserva.Text
'                    tempocompra!SUBTOTAL = txtbase.Text
'                    tempocompra!MONTOINA = txtmonto.Text
'                    tempocompra!IGV = txtigv.Text
'                    tempocompra!TOTAL = txttotal.Text
'                    tempocompra!Empresa = txtempresa.Text
'                    tempocompra!ss = Txt_NumSolComp.Text
'                    tempocompra!Codigo = "" & .Dataset.FieldValues("f3codpro")
'                    tempocompra!Descripcion = "" & .Dataset.FieldValues("f5nompro")
'                    tempocompra!Cantidad = .Dataset.FieldValues("f3canpro")
'                    tempocompra!costo = .Dataset.FieldValues("f3precos")
'                    tempocompra!descuento = .Dataset.FieldValues("f3pordct")
'                    tempocompra!Precio = .Dataset.FieldValues("f3preuni")
'
'
'                    If rsprod.State = adStateOpen Then rsprod.Close
'                    rsprod.Open "SELECT F7CODMED from if5pla where f5codpro='" & .Dataset.FieldValues("f3codpro") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                    If Not (rsprod.EOF) Then
'                        tempocompra!unidad = rsprod.Fields("F7CODMED") & ""
'                    End If
'                    rsprod.Close
'
'                    tempocompra.Update
'                End If
'                .Dataset.Next
'            Loop
'            .Dataset.First
'        End If
'        tempocompra.Close
'        cnn.Close
'    End With
'
'End Sub



Private Sub Grid_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
   'If sw_nuevo_item = False Then
        If Action = daInsert Then
            Grid.Dataset.Edit
            
            Grid.Columns.ColumnByFieldName("ITEM").value = Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "TOP 1 ITEM", "TMPORDENDECOMPRA", vbNullString, vbNullString, vbNullString, "ITEM <> 0 ORDER BY ITEM DESC") & "") + 1 'Grid.Dataset.RecordCount + 1
            
            Grid.Dataset.Post
            
            Grid.Columns.FocusedIndex = 1
        End If
    'End If
    
    If Action = daEdit Then
        If Grid.Columns.ColumnByFieldName("F3CANPROMAX").value > 0 And Grid.Columns.ColumnByFieldName("F3CANPRO").value > Grid.Columns.ColumnByFieldName("F3CANPROMAX").value Then
            MsgBox "No puede poner una cantidad mayor a la del requerimiento", vbCritical
            
            Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("F3CANPRO").value = Grid.Columns.ColumnByFieldName("F3CANPROMAX").value
            Grid.Dataset.Post
            Grid.Dataset.Edit
        End If
    End If
    
    mostrarTotales
End Sub

Private Sub Grid_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    'If sw_nuevo_item = False Then
        If Action = daInsert Then
            If Grid.Dataset.RecordCount > 0 Then
                If Len(Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "")) = 0 Then
                    Allow = False
                Else
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3CODPRO").ColIndex
                End If
            End If
        End If
        
        If Action = daDelete Then
            sw_detalle = True
            
            Grid.Dataset.Refresh
        End If
    'End If
End Sub

Private Sub Grid_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
'    Dim nPorc As Double
'
'    If right(CmbTipDoc.Text, 2) = "02" Then
'        nPorc = gretenc
'    Else
'        nPorc = wwigv
'    End If
'
'    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
'        Case "F5AFECTO"
'            Grid.Dataset.Edit
'
'            If Grid.Columns.ColumnByFieldName("F5AFECTO").Value = True Then
'                Grid.Columns.ColumnByFieldName("F5AFECTO").Value = False
'                'If UCase(Grid.Columns.ColumnByFieldName("F3COLMOD").Value & "") = "F3SINIGV" Then
'                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
'                'Else
'                '    Grid.Columns.ColumnByFieldName("F3sinIGV").Value = Grid.Columns.ColumnByFieldName("F3conIGV").Value
'                'End If
'                Grid.Columns.ColumnByFieldName("F3baseimp").Value = 0
'                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3SINIGV").Value * Grid.Columns.ColumnByFieldName("F3canpro").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
'            Else
'                Grid.Columns.ColumnByFieldName("F5AFECTO").Value = True
'                'If UCase(Grid.Columns.ColumnByFieldName("F3COLMOD").Value & "") = "F3SINIGV" Then
'                    Grid.Columns.ColumnByFieldName("F3CONIGV").Value = Grid.Columns.ColumnByFieldName("F3SINIGV").Value * (1 + (nPorc / 100))
'                'Else
'                    'Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value / (1 + (nPorc / 100))
'                '    Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value * nPorc / 9
'                'End If
'                Grid.Columns.ColumnByFieldName("F3baseimp").Value = Grid.Columns.ColumnByFieldName("F3SINIGV").Value * Grid.Columns.ColumnByFieldName("F3canpro").Value
'                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
'                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3baseimp").Value * nPorc / 100
'            End If
'
'            Grid.Dataset.Post
'    End Select
'
'    CalculaTotal
'    importarDatosRequerimiento

    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
        Case "F5AFECTO"
            If Grid.Dataset.State = dsEdit Then
                Grid.Dataset.Post
            End If
        Case "CERRAR"
            With Grid
                If .Dataset.State = dsEdit Then
                    If Val(.Columns.ColumnByFieldName("PORENTREGAR").value & "") <= 0 Then
                        MsgBox "Producto sin Saldo por Entregar, verifique.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If MsgBox("¿Desea cerrar el Item seleccionado?" & vbNewLine & IIf(Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "") = vbNullString, "STOCK LIBRE", Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")) & " - " & left(Trim(Grid.Columns.ColumnByFieldName("F5NOMPRO_ING").value & ""), 80) & vbNewLine & vbNewLine & _
                                "ATENCIÓN: " & vbNewLine & _
                                "Verifique que los Ingresos por Compra se encuentren actualizados antes de proceder con el cierre de la Orden.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                        
                        With objAyudaOrden
                            .TipoOrden = strTipoOrden
                            .NumeroOrden = strNumeroOrden
                            
                            .Requerimiento = Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
                            .CodigoProducto = Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "")
                            
                            If .cerrarOrden(True) Then
                                
                                Grid.Dataset.Post
                                
                                Grid.Dataset.Edit
                                
                                Grid.Columns.ColumnByFieldName("PORENTREGAR").value = 0
                                Grid.Columns.ColumnByFieldName("CERRAR").value = False
                                
                                Grid.Dataset.Post
                                
                                MsgBox "Item seleccionado cerrado.", vbInformation + vbOKOnly, App.ProductName
                            Else
                                Grid.Dataset.Cancel
                            End If
                        End With
                    Else
                        .Dataset.Cancel
                    End If
                End If
            End With
    End Select
End Sub

Private Sub CalculaTotal()
'    On Error GoTo Errores:
'
'    With Grid
'        If .Dataset.Active = True Then
'
'     'Grid.Dataset.Edit
'        If .Columns.ColumnByFieldName("F5afecto").Value = True Then
'            .Dataset.Edit
'            .Columns.ColumnByFieldName("F3VALDESC").Value = (.Columns.ColumnByFieldName("F3BASEIMP").Value + .Columns.ColumnByFieldName("f3igv").Value) * .Columns.ColumnByFieldName("F3PORDESC").Value / 100
'        Else
'            .Dataset.Edit
'            .Columns.ColumnByFieldName("F3VALDESC").Value = .Columns.ColumnByFieldName("F3monina").Value * .Columns.ColumnByFieldName("F3PORDESC").Value / 100
'        End If
'
'        'Grid.Dataset.Post
'        If right(CmbTipDoc.Text, 2) = "02" Then
'            .Columns.ColumnByFieldName("f3total").Value = .Columns.ColumnByFieldName("F3baseimp").Value + .Columns.ColumnByFieldName("F3monina").Value - .Columns.ColumnByFieldName("F3igv").Value - .Columns.ColumnByFieldName("f3VALDESC").Value
'        Else
'            .Columns.ColumnByFieldName("f3total").Value = .Columns.ColumnByFieldName("F3baseimp").Value + .Columns.ColumnByFieldName("F3monina").Value + .Columns.ColumnByFieldName("F3igv").Value - .Columns.ColumnByFieldName("f3VALDESC").Value
'        End If
'
'        .Dataset.Post
'        End If
'    End With
'
'    Exit Sub
'Errores:
'    Err.Clear
End Sub

Private Sub Grid_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "F3PORCDEMASIA", "F3BASEIMP", "F3MONINA", "F3REDONDEO", "F3IGV", "F3TOTAL", "F3PORDESC"
            'Text = Format(Text, "###,###,##0.00")
            Text = Format(Text, "#,0.00")
        Case "F3SINIGV", "F3CONIGV", "F3NETO", "F3VALDESC"
            If dxCheckBox2.Checked = False Then
                'Text = Format(Text, "###,###,##0.0000")
                Text = Format(Text, "#,0.0000")
            Else
                Text = Format(Text, "#,0.000000")
            End If
        Case "PORENTREGAR"
            If Val(Text) > 0 Then
                Font.Bold = True
                FontColor = vbRed
                Color = vbYellow
            End If
    End Select
End Sub

Private Sub Grid_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "F3VALDESC", "F3PORDESC", "F3BASEIMP", "F3MONINA", "F3IGV", "F3REDONDEO", "F3TOTAL"
            If Mid(Cmbmone.Text, 1, 1) = "S" Then
                Color = &HC0FFFF
            Else
                Color = &HC0FFC0
            End If
            
            Font.Bold = True
            
            Text = Format(Text, "#,0.00")
    End Select
End Sub

Private Sub Grid_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case UCase(Column.FieldName)
        Case "COD_SOLICITUD"
            If ModUtilitario.validarFormAbierto("ayuda_solicitudes_OC") Then
                Unload ayuda_solicitudes_OC
            End If

            objAyudaSolicitud.inicializarEntidades

            With ayuda_solicitudes_OC
                .TipoDocumento = Trim(Txt_TOC.Text)

                .Show 1
            End With

            If objAyudaSolicitud.Codigo <> vbNullString Then
                importarDatosRequerimiento

                listarGrilla

                recalcularItems
            End If
        Case "F3CODPRO"
            If Trim(pnlnomprv.Caption) = vbNullString Then
                MsgBox "Debe Seleccionar un Proveedor.", vbInformation, App.ProductName
                
                Txt_Prove.SetFocus
                
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
'            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").Value & ""), "N") = vbNullString Then
'                MsgBox "No. de Requerimiento no existe, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                Me.MousePointer = vbDefault
'
'                Exit Sub
'            End If
            
            Select Case Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
                Case Is <> vbNullString
'                    If ModUtilitario.validarFormAbierto("frmUtilResumenRequerimiento") Then
'                        Unload frmUtilResumenRequerimiento
'                    End If
'
'                    With frmUtilResumenRequerimiento
'                        .NroPedido = Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
'                        .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
'                        .CodigoProducto = vbNullString
'
'                        .Show 1
'                    End With
'
'                    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILRESUMENREQUERIMIENTO", "PROCESAR", "TRUE", "N") & "") <> 0 Then
'                        copiarSeleccionAyudaResumenRequerimiento Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "")
'                    End If
                Case Else
'                    If ModUtilitario.validarFormAbierto("ayuda_productos") Then
'                        Unload ayuda_productos
'                    End If
'
'                    With ayuda_productos
'                        .CodigoAuxiliar = Trim(Txt_Prove.Text)
'                        .CodigoRequerimiento = vbNullString
'                        .CodigoProducto = vbNullString
'                        .CadenaCorte = vbNullString
'
'                        If .CodigoRequerimiento = vbNullString Then
'                            .CadenaCorte = InputBox("Ingrese cadena de texto a buscar:", App.ProductName, vbNullString)
'                        End If
'
'                        .Show 1
'                    End With
'                    abrirCnTemporal
'
'                    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPPRODUCTOS", "F4PERINT", "-1", "N") & "") <> 0 Then
'                        copiarSeleccionAyudaProductos
'                    End If
            End Select
            
            listarGrilla

            recalcularItems
            
            Me.MousePointer = vbDefault
        Case "F5DESCOSTO":
            wcodcosto = vbNullString: wdescosto = vbNullString: wunicosto = vbNullString
            
            Ayuda_Centros.Show 1
            
            If Len(Trim(wcodcosto)) > 0 Then
                Grid.Dataset.Edit
                Grid.Columns.ColumnByFieldName("F5CODCOSTO").value = wcodcosto
                Grid.Columns.ColumnByFieldName("F5DESCOSTO").value = wdescosto
                Grid.Dataset.Post
            End If
        Case "F5CODCTA"
            Dim gassto As String
            Dim rsgasto As New ADODB.Recordset
            Dim amovs(0 To 3)  As a_grabacion
            
            wctacont = vbNullString: wnomctacont = vbNullString
            
            Ayuda_PlanCta.Show 1
            
            If Len(Trim(wctacont)) > 0 Then
                gassto = ObtenerCampo("BF9GIN", "CODIGO", "CUENTA", wctacont, "T", cnn_dbbancos)
                If Len(Trim(gassto)) = 0 Then
                    csql = "SELECT TOP 1 Val(CODIGO) AS num From BF9GIN ORDER BY Val(CODIGO) DESC"
                    If rsgasto.State = 1 Then rsgasto.Close
                    rsgasto.Open csql, cnn_dbbancos, 3, 1
                    If rsgasto.RecordCount > 0 Then
                        gassto = Format(CStr(rsgasto.Fields("num").value + 1), "000")
                    End If
                    amovs(0).campo = "CODIGO": amovs(0).valor = gassto: amovs(0).Tipo = "T"
                    amovs(1).campo = "BASE": amovs(1).valor = "G": amovs(1).Tipo = "T"
                    amovs(2).campo = "NOMBRE": amovs(2).valor = wnomctacont: amovs(2).Tipo = "T"
                    amovs(3).campo = "CUENTA": amovs(3).valor = wctacont: amovs(3).Tipo = "T"
                    GRABA_REGISTRO amovs(), "BF9GIN", "A", 3, StrConexDbBancos, ""
                End If
                Grid.Dataset.Edit
                Grid.Columns.ColumnByFieldName("F3GASTO").value = gassto
                Grid.Columns.ColumnByFieldName("F5CODCTA").value = wctacont
                Grid.Dataset.Post
            End If
        Case "DESCOLOR"
            
        Case "F3DESMEDIDA":
            If Len(Grid.Columns.ColumnByFieldName("F3CODPRO").value) > 0 Then
                If Val(txt_tc.Text) = 0 Then
                    MsgBox "Tipo de Cambio no puede ser cero, verifique.", vbInformation + vbOKOnly, App.ProductName
                    
                    Exit Sub
                End If
                
                wprodfactor = Grid.Columns.ColumnByFieldName("F3CODPRO").value
                wcodmed = vbNullString
                
                ayuda_um_factor.Show 1
                
                'rstvp.Open "select * from MEDIVENTAS where F5CODPRO = '" & Grid.Columns.ColumnByFieldName("F3CODPRO").Value & "' and F7CODMED = '" & wcodmed & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
                
                sw_nuevo_item = True
                
                Grid.Dataset.Edit
                
                sw_nuevo_item = False
                
                If Len(wcodmed & "") <> 0 Then
                    sw_nuevo_item = True
                    
                    'Grid.Dataset.Edit
                    
                    sw_nuevo_item = False
                    
                    Grid.Columns.ColumnByFieldName("F3CODMEDIDA").value = wcodmed
                    Grid.Columns.ColumnByFieldName("F3DESMEDIDA").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", wcodmed, "T")
                    
                    If Val(Grid.Columns.ColumnByFieldName("CANT_ANT").value & "") = 0 Then
                        If MsgBox("¿La cantidad ingresada esta convertida a la U.M. " & Trim(Grid.Columns.ColumnByFieldName("F3DESMEDIDA").value & "") & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            Grid.Columns.ColumnByFieldName("CANT_ANT").value = Val(Format(Grid.Columns.ColumnByFieldName("F3CANPRO").value * Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", wprodfactor, "T", "AND F7CODMED = '" & wcodmed & "'")), "#0.00"))
                        Else
                            Grid.Columns.ColumnByFieldName("CANT_ANT").value = Val(Format(Val(Grid.Columns.ColumnByFieldName("F3CANPRO").value & ""), "#0.00"))
                        End If
                    End If
                    
                    If Val(Grid.Columns.ColumnByFieldName("F3CANPROMAX").value & "") > 0 Then
                        Grid.Columns.ColumnByFieldName("F3CANPROMAX").value = Val(Format(Grid.Columns.ColumnByFieldName("CANT_ANT").value / Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", wprodfactor, "T", "AND F7CODMED = '" & wcodmed & "'")), "#0.00"))
                    End If
                    
                    Grid.Columns.ColumnByFieldName("F3CANPRO").value = Val(Format(Grid.Columns.ColumnByFieldName("CANT_ANT").value / Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", wprodfactor, "T", "AND F7CODMED = '" & wcodmed & "'")), "#0.00"))
                    
                    'If Val(Grid.Columns.ColumnByFieldName("F3SINIGV").value & "") = 0 Then
                        'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                        With objAyudaVale
                            .CodigoProveedor = Trim(Txt_Prove.Text)
                            .CodigoProducto = wprodfactor
                            
                            .obtenerUltimoPrecioSinIgvProductoDeProveedor
                            
                            Select Case .CodigoMoneda
                                Case "S"
                                    Grid.Columns.ColumnByFieldName("F3SINIGV").value = Format(Val(.ValorVenta / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text))), "#0.0000")
                                Case Else
                                    Grid.Columns.ColumnByFieldName("F3SINIGV").value = Format(Val(.ValorVentaDol * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text))), "#0.0000")
                            End Select
                        End With
                    'End If
                    
                    Grid.Columns.ColumnByFieldName("F3SINIGV").value = Grid.Columns.ColumnByFieldName("F3SINIGV").value * Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", wprodfactor, "T", "AND F7CODMED = '" & wcodmed & "'"))
                    Grid.Columns.ColumnByFieldName("F3CONIGV").value = 0
                    
                    sw_nuevo_item = True
                    
                    Grid.Dataset.Post
                    
                    recalcularItems
                    
                    sw_nuevo_item = False
                    
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("f3canpro").ColIndex
                Else
                    sw_nuevo_item = True
                    
                    Grid.Dataset.Post
                    
                    sw_nuevo_item = False
                    
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3MEDIDA").ColIndex
                End If
            Else
                MsgBox "Seleccione primero el Producto", vbInformation, "Sistema de Logística"
                
                Grid.Columns.FocusedIndex = 1
            End If
    End Select

'----------------------------------- 2014
        
   
'''    If dxDBGrid1.Columns.FocusedIndex = 1 Then
'''        wdestino = "E"
'''        wgastos = ""
'''        Sw_AyuCodProv = False
'''        ayuda_gastos.TipoConcepto = "E"
'''        ayuda_gastos.Show 1
'''        dxDBGrid1.Dataset.Edit
'''        If Len(Trim(wgastos)) > 0 Then
'''            dxDBGrid1.Columns.ColumnByFieldName("F3GASTO").Value = wgastos
'''            dxDBGrid1.Columns.ColumnByFieldName("F3CTACON").Value = wctacont
'''            If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").Value & "")) = 0 Then
'''                dxDBGrid1.Columns.ColumnByFieldName("F3CONCEPTO").Value = wnomgasto
'''            End If
'''        End If
'''        dxDBGrid1.Dataset.Post
'''        PROCESO_CUENTA
'''
'''        If dxDBGrid1.Columns.ColumnByFieldName("F3CENCOS").ReadOnly = False Then
'''            dxDBGrid1.Columns.FocusedIndex = 3
'''        End If
'''
'''    End If
'----------------------------------- 2014
 
    If Grid.Columns.FocusedColumn.ObjectName = "COLUMNELIMINAR" Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If Grid.Dataset.RecordCount = 1 Then
                Grid.Dataset.Delete
                AdicionaItemGrid
                sw_detalle = False
                'atbmenu.Tools("ID_Grabar").Enabled = False
            Else
                Grid.Dataset.Delete
            End If
            'calcula
            mostrarTotales
            
            sw_nuevo_item = False
        End If
    End If
    
    Select Case UCase(Grid.Columns.FocusedColumn.Caption)
        Case "?"
            'Codigo_producto = Grid.Columns.ColumnByFieldName("F3CODPRO").value
            If Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "") = vbNullString Then
                MsgBox "Producto no especificado, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If ModUtilitario.validarFormAbierto("ayuda_prov_prod") Then
                Unload ayuda_prov_prod
            End If
            
            With ayuda_prov_prod
                objAyudaOrden.inicializarEntidades
                objAyudaOrden.inicializarEntidadesDetalle
                
                .CodigoProducto = Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "")
                
                .Show 1
            End With
            
            With objAyudaOrden
                If .PrecioSinImpuesto > 0 Then
                    'Grid.Dataset.Edit
                    
                    Select Case .CodMoneda
                        Case "S"
                            .PrecioSinImpuesto = Val(Format(.PrecioSinImpuesto / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text)), "#0.0000"))
                        Case Else
                            .PrecioSinImpuesto = Val(Format(.PrecioSinImpuesto * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text)), "#0.0000"))
                    End Select
                    
                    'Grid.Columns.ColumnByFieldName("F3SINIGV").value = Grid.Columns.ColumnByFieldName("F3SINIGV").value * Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & ""), "T", "AND F7CODMED = '" & Trim(Grid.Columns.ColumnByFieldName("F3CODMEDIDA").value & "") & "'"))
                    'Grid.Columns.ColumnByFieldName("F3PORDCT").value = .PorcentajeDscto
                    
                    'Grid.Dataset.Post
                    
                    Dim rstTemporalEditButton As New ADODB.Recordset
                    
                    If rstTemporalEditButton.State = 1 Then rstTemporalEditButton.Close
                    
                    rstTemporalEditButton.Open "SELECT * FROM TMPORDENDECOMPRA WHERE F3CODPRO = '" & Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "") & "'", cnDBTemp, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rstTemporalEditButton.EOF Then
                        rstTemporalEditButton.MoveFirst
                        
                        Grid.Dataset.Close
                        
                        Do While Not rstTemporalEditButton.EOF
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE "
                            SqlCad = SqlCad & "TMPORDENDECOMPRA "
                            SqlCad = SqlCad & "SET "
                            SqlCad = SqlCad & "F3SINIGV = " & .PrecioSinImpuesto * Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5FACTOR", "MEDIVENTAS", "F5CODPRO", Trim(rstTemporalEditButton!F3CODPRO & ""), "T", "AND F7CODMED = '" & Trim(rstTemporalEditButton!F3CODMEDIDA & "") & "'")) & ", "
                            SqlCad = SqlCad & "F3PORDCT = " & .PorcentajeDscto & " "
                            SqlCad = SqlCad & "WHERE "
                            SqlCad = SqlCad & "F3CODPRO = '" & Trim(rstTemporalEditButton!F3CODPRO & "") & "'"
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            rstTemporalEditButton.MoveNext
                        Loop
                            listarGrilla
                    
                            recalcularItems
                    End If
                End If
            End With
            
'            If wvv_prod <> 0 And wpv_prod <> 0 Then
'                Grid.Dataset.Edit
'
'                Grid.Columns.ColumnByFieldName("F3sinigv").value = wvv_prod
'                Grid.Columns.ColumnByFieldName("F3conIGV").value = wpv_prod
'
'                Grid.Dataset.Post
'
'                recalcularItems
'            End If
'            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
'                If right(CmbTipDoc.Text, 2) = "02" Then
'                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
'                End If
'
'                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (wwigv / 100)
'            Else
'                Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
'                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
'            End If
'
'            Grid.Columns.ColumnByFieldName("f3total").Value = Grid.Columns.ColumnByFieldName("F3baseimp").Value + Grid.Columns.ColumnByFieldName("F3monina").Value + Grid.Columns.ColumnByFieldName("F3igv").Value - Grid.Columns.ColumnByFieldName("f3VALDESC").Value
        Case "%"
            With frmListaProvDscto
                If Trim(Txt_Prove.Text) = vbNullString Then
                    MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                    
                    Txt_Prove.SetFocus
                    
                    Exit Sub
                End If
                
                If Grid.Columns.ColumnByFieldName("F3CODPRO").value = vbNullString Then
                    MsgBox "Seleccione el Producto.", vbInformation + vbOKOnly, App.ProductName
                    
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3CODPRO").ColIndex
                    
                    Exit Sub
                End If
                
                If Grid.Columns.ColumnByFieldName("F3CODMEDIDA").value = vbNullString Then
                    MsgBox "Seleccione la U.M. del Producto.", vbInformation + vbOKOnly, App.ProductName
                    
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3DESMEDIDA").ColIndex
                    
                    Exit Sub
                End If
                
                If Val(Grid.Columns.ColumnByFieldName("F3CANPRO").value & "") = 0 Then
                    MsgBox "Ingrese la cantidad solicitada del producto.", vbInformation + vbOKOnly, App.ProductName
                    
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3CANPRO").ColIndex
                    
                    Exit Sub
                End If
                
                If Val(Grid.Columns.ColumnByFieldName("F3SINIGV").value & "") = 0 Then
                    MsgBox "Ingrese el precio del producto.", vbInformation + vbOKOnly, App.ProductName
                    
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3SINIGV").ColIndex
                    
                    Exit Sub
                End If
                
                objAyudaProvDscto.inicializarEntidades
                
                .Ayuda = True
                
                .CodigoProveedor = Trim(Txt_Prove.Text)
                .CodigoProducto = Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "")
                .CodigoUM = Trim(Grid.Columns.ColumnByFieldName("F3CODMEDIDA").value & "")
                .Cantidad = Val(Grid.Columns.ColumnByFieldName("F3CANPRO").value & "")
                
                .Show vbModal
                
                If objAyudaProvDscto.Porcentaje <> 0 Then
                    With Grid.Dataset
                        .Edit
                        
                        .FieldValues("F3PORDESC") = objAyudaProvDscto.Porcentaje
                        
                        .Post
                        
                        Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3VALDESC").ColIndex
                    End With
                    
                    objAyudaProvDscto.inicializarEntidades
                End If
                
                .Ayuda = False
            End With
        Case "ATE."
            If Trim(pnlnomprv.Caption) = vbNullString Then
                MsgBox "Debe Seleccionar un Proveedor.", vbInformation, App.ProductName
                
                Txt_Prove.SetFocus
                
                Exit Sub
            End If
            
            If Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "") = vbNullString Then
                MsgBox "Debe Seleccionar y/o Ingresar un Codigo de Producto.", vbInformation, App.ProductName
                
                Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("F3CODPRO").ColIndex
                
                Exit Sub
            End If
            
''            If ModUtilitario.validarFormAbierto("ayuda_productos") Then
''                Unload ayuda_productos
''            End If
''
''            With ayuda_productos
''                .CodigoAuxiliar = Trim(Txt_Prove.Text)
''                .CodigoRequerimiento = vbNullString
''                .CodigoProducto = Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").Value & "")
''
''                .Show 1
''            End With
''
''            abrirCnTemporal
''
''            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPPRODUCTOS", "F4PERINT", "-1", "N") & "") = 0 Then
''                MsgBox "No se registro ninguna selección de Productos.", vbInformation + vbOKOnly, App.ProductName
''            Else
''                copiarSeleccionAyudaProductos
''            End If
            Me.MousePointer = vbHourglass
            
            If ModUtilitario.validarFormAbierto("frmUtilResumenRequerimiento") Then
                Unload frmUtilResumenRequerimiento
            End If
            
            With frmUtilResumenRequerimiento
                .NroPedido = vbNullString 'Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").Value & "")
                .CodigoProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
                .CodigoProducto = Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "")
                
                .Show 1
            End With
            
            If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPUTILRESUMENREQUERIMIENTO", "PROCESAR", "TRUE", "N") & "") <> 0 Then
                copiarSeleccionAyudaResumenRequerimiento vbNullString
            End If
            
            listarGrilla
            
            recalcularItems
            
            Me.MousePointer = vbDefault
    End Select
End Sub

Private Sub Grid_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Rem SK ADD:
    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
        Case "F3CANPRO", "F3PORCDEMASIA", "F3SINIGV", "F3CONIGV", "F3PORDESC", "F3VALDESC", "F5AFECTO"
            If Grid.Dataset.State = dsEdit Then
                With objAyudaOrden
                    'DATOS
                    .PorcentajeImpuesto = IIf(right(CmbTipDoc.Text, 2) = "02", gretenc, wwigv) / 100
                    .SignoImpuesto = IIf(right(CmbTipDoc.Text, 2) = "02", -1, 1)
                    
                    .Cantidad = Val(Grid.Columns.ColumnByFieldName("F3CANPRO").value & "")
                    .CantidadMaxima = Val(Grid.Columns.ColumnByFieldName("F3CANPROMAX").value & "")
                    .PorcentajeDemasia = Val(Grid.Columns.ColumnByFieldName("F3PORCDEMASIA").value & "") / 100
                    
                    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
                        Case "F3SINIGV"
                            .PrecioSinImpuesto = Val(Grid.Columns.ColumnByFieldName("F3SINIGV").value & "")
                            .PrecioConImpuesto = 0
                        Case "F3CONIGV"
                            .PrecioSinImpuesto = 0
                            .PrecioConImpuesto = Val(Grid.Columns.ColumnByFieldName("F3CONIGV").value & "")
                        Case Else
                            .PrecioSinImpuesto = Val(Grid.Columns.ColumnByFieldName("F3SINIGV").value & "")
                            .PrecioConImpuesto = Val(Grid.Columns.ColumnByFieldName("F3CONIGV").value & "")
                    End Select
                    
                    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
                        Case "F3PORDESC"
                            .PorcentajeDscto = Val(Grid.Columns.ColumnByFieldName("F3PORDESC").value & "") / 100
                            .TotalDscto = 0
                        Case "F3VALDESC"
                            .PorcentajeDscto = 0
                            .TotalDscto = Val(Grid.Columns.ColumnByFieldName("F3VALDESC").value & "")
                        Case Else
                            .PorcentajeDscto = Val(Grid.Columns.ColumnByFieldName("F3PORDESC").value & "") / 100
                            .TotalDscto = Val(Grid.Columns.ColumnByFieldName("F3VALDESC").value & "")
                    End Select
                    
                    If Not IsNull(Grid.Columns.ColumnByFieldName("F5AFECTO").value) Then
                        .Afecto = CBool(Grid.Columns.ColumnByFieldName("F5AFECTO").value)
                    End If
                    
                    Grid.Dataset.Edit
                    
                    If .CantidadMaxima > 0 Then
                        If .Cantidad > .CantidadMaxima Then
                            MsgBox "La cantidad no puede exceder al saldo requerido, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            Grid.Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    'CALCULOS
                    .calculosPorItem
                    
                    Grid.Columns.ColumnByFieldName("F3SINIGV").value = .PrecioSinImpuesto
                    Grid.Columns.ColumnByFieldName("F3CONIGV").value = .PrecioConImpuesto
                    Grid.Columns.ColumnByFieldName("F3PORDESC").value = Val(Format(.PorcentajeDscto * 100, "#0.00"))
                    Grid.Columns.ColumnByFieldName("F3CANPROFINAL").value = .CantidadFinal
                    Grid.Columns.ColumnByFieldName("F3VALDESC").value = .TotalDscto
                    
                    Grid.Columns.ColumnByFieldName("F3NETO").value = .PrecioNetoSinImpuesto
                    
                    Grid.Columns.ColumnByFieldName("F3BASEIMP").value = .BasePorItem
                    Grid.Columns.ColumnByFieldName("F3MONINA").value = .ExoneradoPorItem
                    Grid.Columns.ColumnByFieldName("F3IGV").value = .ImpuestoPorItem
                    Grid.Columns.ColumnByFieldName("F3TOTAL").value = .TotalPorItem
                    
                    Grid.Dataset.Post
                End With
            End If
            
            mostrarTotales
    End Select


'    Dim nPorc As Double
'    Dim nValor As Double
'
'    If right(CmbTipDoc.Text, 2) = "02" Then
'        If left(Cmbmone.Text, 1) = "D" Then
'            nValor = Val(Grid.Columns.ColumnByFieldName("f3sinigv").Value & "") * Val(txt_tc.Text)
'        Else
'            nValor = Val(Grid.Columns.ColumnByFieldName("f3sinigv").Value & "")
'        End If
'
'        nPorc = gretenc
'        'If nValor >= 700 Then
'        '    Grid.Dataset.Edit
'        '    Grid.Columns.ColumnByFieldName("F5afecto").Value = True
'        '    Grid.Dataset.Post
'        'Else
'        '    Grid.Dataset.Edit
'        '    Grid.Columns.ColumnByFieldName("F5afecto").Value = False
'        '    Grid.Dataset.Post
'        'End If
'    Else
'        nPorc = wwigv
'    End If
'
'    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
'        Case "F3CANPRO"
'            Grid.Dataset.Edit
'
'            If Val(Grid.Columns.ColumnByFieldName("F3CANPROMAX").Value & "") > 0 Then
'                If Val(Grid.Columns.ColumnByFieldName("F3CANPRO").Value & "") > Val(Grid.Columns.ColumnByFieldName("F3CANPROMAX").Value & "") Then
'                    MsgBox "La cantidad no puede exceder al saldo pendiente de atención del requerimiento, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                    Grid.Dataset.Cancel
'
'                    Exit Sub
'                End If
'            End If
'
'            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
'                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
'            Else
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
'                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
'            End If
'            Grid.Dataset.Post
'        Case "F3CONIGV"
'            Grid.Dataset.Edit
'
'            Grid.Columns.ColumnByFieldName("F3COLMOD").Value = UCase(Grid.Columns.FocusedColumn.FieldName)
'
'            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
'                If right(CmbTipDoc.Text, 2) = "02" Then
'                    Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value * nPorc / 9
'                Else
'                    Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value / (1# + (nPorc / 100))
'                End If
'                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
'            Else
'                Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
'                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
'            End If
'
'            Grid.Dataset.Post
'
'        Case "F3SINIGV"
'            Grid.Dataset.Edit
'
'            Grid.Columns.ColumnByFieldName("F3COLMOD").Value = UCase(Grid.Columns.FocusedColumn.FieldName)
'
'            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
'                If right(CmbTipDoc.Text, 2) = "02" Then
'                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
'                Else
'                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value * (1# + (nPorc / 100))
'                End If
'                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
'            Else
'                Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
'                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
'                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
'                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
'            End If
'
'            Grid.Dataset.Post
'        Case "F3PORDESC"
'
'    End Select
'
'    CalculaTotal
    
End Sub





Private Sub grid_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    On Error GoTo errOnKeyDown
    
    For d = 0 To 25
        nSaveRecNo = Grid.Dataset.RecNo
    Next
    
    Dim strCodigoProducto As String
            
    strCodigoProducto = Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "")
    
    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
        Case "COD_SOLICITUD"
'            Select Case KeyCode
'                Case vbKeyReturn
'                    If Grid.Dataset.State = dsEdit Then Grid.Dataset.Post
'
'                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & ""), "N") <> vbNullString Then
'                        With Grid.Dataset
'                            'If .State = dsEdit Then .Post
'
'                            .Edit
'
'                            Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "COD_SOLICITUD", "TB_CABSOLICITUD", "VAL(COD_SOLICITUD)", Val(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & ""), "N")
'
'                            .Post
'                        End With
'
'                        Grid_OnEditButtonClick Grid.Columns.ColumnByFieldName("F3CODPRO"), Nothing
'
'                        If Grid.Dataset.State = dsEdit Then Grid.Dataset.Post
'                    End If
'            End Select
'        Case "F3CODPRO"
'            Select Case KeyCode
'                Case vbKeyF1
'                    Grid_OnEditButtonClick Grid.Columns.FocusedColumn, Nothing
'            End Select
        Case "F5NOMPRO"
            Select Case KeyCode
                Case vbKeyTab
                    If Grid.Dataset.State = dsEdit Then
                        Grid.Dataset.Post
                    End If
                    
                    If ModUtilitario.validarFormAbierto("frmListaBien") Then
                        Unload frmListaBien
                    End If
                    
                    With frmListaBien
                        '.Ayuda = True
                        '.TieneMovimientoAlmacen = IIf(strTipoOrden = "OC", True, False)
                        '.SoloServicios = IIf(strTipoOrden = "OC", False, True)
                        '.InsumoOP = False
                        '.CadenaCorte = Trim(Grid.Columns.ColumnByFieldName("F5NOMPRO").value & "")
                        
                        .Ayuda = True
                        .InsumoOP = False
                        .ParaVenta = False
                        .TieneMovimientoAlmacen = IIf(strTipoOrden = "OC", True, False)
                        .CadenaCorte = Trim(Grid.Columns.ColumnByFieldName("F5NOMPRO").value & "")
                        .FiltroAdicional = vbNullString
                        .TipoBienMostrar = IIf(strTipoOrden = "OC", "P", "S")
                        
                        .ParaVenta = False
                        
                        objAyudaBien.inicializarEntidades
                        
                        .Show 1
                        
                        If objAyudaBien.Codigo <> vbNullString Then
                            objAyudaBien.obtenerConfigBien
                            
                            If Trim(Grid.Columns.ColumnByFieldName("F3CODPRO").value & "") <> objAyudaBien.Codigo Then
                                If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(*)", "TMPORDENDECOMPRA", "F3CODPRO", objAyudaBien.Codigo, "T", "AND TRIM(COD_SOLICITUD & '') = '" & Trim(Grid.Columns.ColumnByFieldName("COD_SOLICITUD").value & "") & "'")) > 0 Then
                                    MsgBox "Imposible adicionar el producto seleccionado, ya se encuentra registrado en la Orden, verifique.", vbInformation + vbOKOnly, App.ProductName
                                    
                                    Exit Sub
                                End If
                            End If
                            
                            Dim strCuentaContable As String
                            
                            Select Case strTipoOrden
                                Case "OC"
                                    With objAyudaProveedor
                                        .Codigo = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
                                        
                                        .obtenerConfigProveedor
                                    End With
                                    
                                    Select Case objAyudaProveedor.OrigenProveedor
                                        Case "N"
                                            If objAyudaBien.CtaContable = vbNullString Then
'                                                MsgBox "Imposible adicionar el producto seleccionado ya que no tiene configurado su Cuenta Contable para Proveedores Nacionales." & vbNewLine & vbNewLine & _
'                                                        "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
'
'                                                Exit Sub
                                            Else
                                                strCuentaContable = objAyudaBien.CtaContable
                                            End If
                                        Case "E"
                                            If objAyudaBien.CtaContableImportacion = vbNullString Then
'                                                MsgBox "Imposible adicionar el producto seleccionado ya que no tiene configurado su Cuenta Contable para Proveedores Extranjeros." & vbNewLine & vbNewLine & _
'                                                        "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
'
'                                                Exit Sub
                                            Else
                                                strCuentaContable = objAyudaBien.CtaContableImportacion
                                            End If
                                    End Select
                                Case "OS"
                                    If objAyudaBien.CtaContable = vbNullString Then
'                                        MsgBox "Imposible adicionar el producto seleccionado ya que no tiene configurado su Cuenta Contable." & vbNewLine & vbNewLine & _
'                                                "Comuniquese con el área de Contabilidad para la asignación de Cuenta Contable correspondiente.", vbInformation + vbOKOnly, App.ProductName
'
'                                        Exit Sub
                                    Else
                                        strCuentaContable = objAyudaBien.CtaContable
                                    End If
                            End Select
                            
                            If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T") = vbNullString Then
                                With objAyudaGasto
                                    .inicializarEntidades
                                    
                                    .Codigo = vbNullString
                                    .Base = "G"
                                    .CuentaContable = strCuentaContable
                                    .Descripcion = ModUtilitario.ObtenerCampoV2(cnDBContaTabla, "F5NOMCTA", "CF5PLA", "F5CODCTA", strCuentaContable, "T")
                                    .TipoGasto = "P"
                                    .Moneda = left(Cmbmone.Text, 1)
                                    .GrupoFlujo = vbNullString
                                    
                                    If .guardarGasto Then
                                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                                    End If
                                    
                                    .inicializarEntidades
                                End With
                            End If
                            
                            Dim strUltimaDescripcion As String
                            Dim dblUltimoPrecioSinIGv As Double
                            
                            
                                With objAyudaOrden
                                    .CodProveedor = Trim(Txt_Prove.Text)
                                    .CodigoProducto = objAyudaBien.Codigo
                                    
                                    strUltimaDescripcion = .obtenerUltimaDescripcionProductoDeProveedor
                                End With
                                'Obtener el Precio  en la Ultima Compra Ingresada (Ingreso por Compra)
                                With objAyudaVale
                                    .CodigoProveedor = Trim(Txt_Prove.Text)
                                    .CodigoProducto = objAyudaBien.Codigo
                                    
                                    .obtenerUltimoPrecioSinIgvProductoDeProveedor
                                    
                                    Select Case .CodigoMoneda
                                        Case "S"
                                            dblUltimoPrecioSinIGv = Format(Val(.ValorVenta / IIf(left(Cmbmone.Text, 1) = "S", 1, Val(txt_tc.Text))), "#0.0000")
                                        Case Else
                                            dblUltimoPrecioSinIGv = Format(Val(.ValorVentaDol * IIf(left(Cmbmone.Text, 1) = "D", 1, Val(txt_tc.Text))), "#0.0000")
                                    End Select
                                End With
                            
                            With Grid
                                .Dataset.Edit
                                
                                .Columns.ColumnByFieldName("F3GASTO").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CODIGO", "BF9GIN", "CUENTA", strCuentaContable, "T")
                                .Columns.ColumnByFieldName("F3CUENTA").value = strCuentaContable
                                .Columns.ColumnByFieldName("F3CODPRO").value = objAyudaBien.Codigo
                                .Columns.ColumnByFieldName("F5NOMPRO").value = IIf(strUltimaDescripcion <> vbNullString, strUltimaDescripcion, objAyudaBien.Descripcion)
                                .Columns.ColumnByFieldName("F5NOMPRO_ING").value = objAyudaBien.Descripcion
                                .Columns.ColumnByFieldName("F3CODMEDIDA").value = objAyudaBien.CodUM
                                .Columns.ColumnByFieldName("F3DESMEDIDA").value = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", objAyudaBien.CodUM, "T")
                                .Columns.ColumnByFieldName("F3CANPRO").value = IIf(strTipoOrden = "OC", 0, 1)
                                .Columns.ColumnByFieldName("F3CANPROMAX").value = 0
                                .Columns.ColumnByFieldName("F3CANPROFINAL").value = 0
                                .Columns.ColumnByFieldName("F3SINIGV").value = dblUltimoPrecioSinIGv
                                .Columns.ColumnByFieldName("F5AFECTO").value = objAyudaBien.Afecto
                                
                                .Dataset.Post
                            End With
                        End If
                    End With
                    
                    listarGrilla
                    
                    recalcularItems
                Case vbKeyF3
                    Dim strNuevaDescripcion As String
                    
                    If Grid.Dataset.State = dsEdit Then
                        Grid.Dataset.Post
                    End If
                    
                    strNuevaDescripcion = Trim(InputBox("Edite la Descripción del Producto para el Proveedor:", "Reemplazar Descripción", Trim(Grid.Columns.ColumnByFieldName("F5NOMPRO").value & "")))
                    
                    If strNuevaDescripcion <> vbNullString Then
                        If MsgBox("¿Desea completar la acción de Reemplazo?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            Grid.Dataset.Close
                            
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE TMPORDENDECOMPRA SET F5NOMPRO = '" & Replace(UCase(strNuevaDescripcion), "'", "' & CHR(39) & '", 1) & "' WHERE F3CODPRO = '" & strCodigoProducto & "'"
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            SqlCad = vbNullString
                            
                            listarGrillaOrden
                        End If
                    End If
            End Select
        Case "F3SINIGV"
            Select Case KeyCode
                Case vbKeyF3
                    Dim dblNuevoPrecioSinIgv As Double
                    
                    If Grid.Dataset.State = dsEdit Then
                        Grid.Dataset.Post
                    End If
                    
                    dblNuevoPrecioSinIgv = Val(InputBox("Ingrese el Precio S/Igv del Producto para el Proveedor:", "Reemplazar Precio S/Igv", Trim(Grid.Columns.ColumnByFieldName("F3SINIGV").value & "")))
                    
                    If dblNuevoPrecioSinIgv > 0 Then
                        If MsgBox("¿Desea completar la acción de Reemplazo?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            Grid.Dataset.Close
                            
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE TMPORDENDECOMPRA SET F3SINIGV = " & dblNuevoPrecioSinIgv & ", F3CONIGV = 0 WHERE F3CODPRO = '" & strCodigoProducto & "'"
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            SqlCad = vbNullString
                            
                            listarGrillaOrden
                            
                            recalcularItems
                        End If
                    End If
            End Select
        Case "F3CONIGV"
            Select Case KeyCode
                Case vbKeyF3
                    Dim dblNuevoPrecioConIgv As Double
                    
                    dblNuevoPrecioConIgv = Val(InputBox("Ingrese el Precio C/Igv del Producto para el Proveedor:", "Reemplazar Precio C/Igv", Trim(Grid.Columns.ColumnByFieldName("F3CONIGV").value & "")))
                    
                    If dblNuevoPrecioConIgv > 0 Then
                        If MsgBox("¿Desea completar la acción de Reemplazo?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE TMPORDENDECOMPRA SET F3SINIGV = 0, F3CONIGV = " & dblNuevoPrecioConIgv & " WHERE F3CODPRO = '" & strCodigoProducto & "'"
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            SqlCad = vbNullString
                            
                            listarGrillaOrden
                            
                            recalcularItems
                        End If
                    End If
            End Select
        Case "F3PORDESC"
            Select Case KeyCode
                Case vbKeyF3
                    Dim dblNuevoPorcentajeDscto As Double
                    
                    dblNuevoPorcentajeDscto = Val(InputBox("Ingrese el Porcentaje de Descuento del Producto:", "Reemplazar Porcentaje Descuento", Trim(Grid.Columns.ColumnByFieldName("F3PORDESC").value & "")))
                    
                    If dblNuevoPorcentajeDscto > 0 Then
                        If MsgBox("¿Desea completar la acción de Reemplazo?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            SqlCad = vbNullString
                            SqlCad = SqlCad & "UPDATE TMPORDENDECOMPRA SET F3PORDESC = " & dblNuevoPorcentajeDscto & " WHERE F3CODPRO = '" & strCodigoProducto & "'"
                            
                            abrirCnTemporal
                            
                            cnDBTemp.Execute SqlCad
                            
                            SqlCad = vbNullString
                            
                            listarGrillaOrden
                            
                            recalcularItems
                        End If
                    End If
            End Select
    End Select
    
    If Grid.Dataset.RecordCount >= nSaveRecNo Then
        Grid.Dataset.RecNo = nSaveRecNo
        
        Grid.SetFocus
    End If
    
    Exit Sub
errOnKeyDown:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    listarGrillaOrden
    
    If Grid.Dataset.RecordCount = 0 Then
        adicionarItemOrden
    End If
    
    Me.MousePointer = vbDefault
    
    Err.Clear
End Sub

Private Sub txt_fecha_CloseUp()
    txt_tc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", Trim(txt_fecha.value), "F")), "0.000")
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txt_fecha_LostFocus()

    If IsDate(txt_fecha.value) Then
        If Val(txt_tc.Text & "") = 0# Then
            If rscambios.State = adStateOpen Then rscambios.Close
            If ctipoadm_bd = "M" Then
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE FECHA='" & txt_fecha.value & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            Else
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & txt_fecha.value & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            End If
            If Not rscambios.EOF Then
                txt_tc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
            Else
                txt_tc.Text = Format(3.643, "0.000")
            End If
            rscambios.Close
        End If
    Else
        MsgBox "Fecha incorrecta. Verifique.", vbCritical, "Atención"
        txt_fecha.SetFocus
    End If

End Sub

Private Sub Txt_NumOC_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_NumOC_GotFocus()
    Txt_NumOC.SelStart = 0: Txt_NumOC.SelLength = Len(Txt_NumOC)

End Sub

Private Sub Txt_NumSolComp_Change()

    If Not inicio Then swGrabacion = True
        atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
        atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
        'atbmenu.Tools.ITEM("IDEmail").Enabled = True
        atbmenu.Tools.ITEM("ID_Anular").Enabled = True
End Sub

Private Sub Txt_Prove_GotFocus()
    ModUtilitario.seleccionarTextoCaja Txt_Prove
End Sub

Private Sub Txt_Prove_LostFocus()

'    If sw_ayuda = False Then
'        If Len(Trim(Txt_Prove.Text)) > 0 Then
'            If rst.State = adStateOpen Then rst.Close
'            rst.Open "SELECT F2NOMPROV,F2DIRPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & Trim(Txt_Prove.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not rst.EOF Then
'                pnlnomprv.Caption = "" & rst.Fields("F2NOMPROV")
'                pnldireprv.Caption = "" & rst.Fields("F2DIRPROV")
'                GRABA_GRID Trim(Txt_Prove.Text)
'            Else
'                MsgBox "El proveedor no existe. Verifique.", vbInformation, "Atención"
'                Txt_Prove.SetFocus
'            End If
'            If rst.State = adStateOpen Then rst.Close
'        End If
'    End If
    
    If Trim(Txt_Prove.Text) <> vbNullString And Trim(pnlnomprv.Caption) = vbNullString Then
        Txt_Prove.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NEWRUC", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
        pnlnomprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
        pnldireprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
        txtcontacto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CONTACTO", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
        
        CmbTipDoc.ListIndex = ModUtilitario.seleccionarItem(CmbTipDoc, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TIPDOC", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T"), "DER", 2)
        
        txtcodforma.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text), "T")
        pnlnomforma.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", Trim(txtcodforma.Text), "T")
    End If
End Sub

Private Sub Txt_Referencia_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txt_tc_Change()

    If Not inicio Then swGrabacion = True
    
    If Val(txt_tc.Text) = 0 Then
        txt_tc.Text = "3.640"
    End If
    
End Sub

Private Sub txtcodcosto_Change()

    If Not inicio Then swGrabacion = True
    If Len(Trim(TxtCodCosto.Text)) Mod 3 = 0 Then
        txtcodcosto_KeyPress 13
    End If
End Sub



Private Sub txtcodcosto_DblClick()
wcodcosto = "": wdescosto = "": wunicosto = "":
            
            Ayuda_Centros.Show 1
                      
            If Len(Trim(wcodcosto)) > 0 Then
                
                TxtCodCosto.Text = wcodcosto
                
               
            End If
End Sub

Private Sub txtcodcosto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        txtcodcosto_DblClick
    End If
    Select Case KeyCode
        Case vbKeyReturn
            ModUtilitario.pulsarTecla vbKeyTab
    End Select
End Sub

Private Sub txtcodcosto_KeyPress(KeyAscii As Integer)
Set rst = New ADODB.Recordset
    If KeyAscii = 13 Then
        If Len(Trim(TxtCodCosto.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "select f3descrip,F2DIRECION from centros where f3costo='" & TxtCodCosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                PnlNomCosto.Caption = Trim(rst.Fields("f3descrip") & "")
                If TxtCodCosto.Text = "998" Then
                    dxDBGrid1.Columns.ColumnByName("f3cencos").Visible = True
                    Grid.Columns.ColumnByFieldName("f5DESCOSTO").Visible = True
                Else
'                    dxDBGrid1.Columns.ColumnByName("f3cencos").Visible = False
'                    Grid.Columns.ColumnByFieldName("f5DESCOSTO").Visible = False
                End If
                txtlugar_entrega.Text = Trim(rst.Fields("F2DIRECION") & "")
                'ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
            Else
                MsgBox "Centro de costo no existe. Verifique.", vbInformation, "Atenciòn"
                'TxtCodCosto.SetFocus
            End If
            rst.Close
        End If
    End If

End Sub

Private Sub txtcodcosto_LostFocus()

   txtcodcosto_KeyPress 13

End Sub

Private Sub txtcodforma_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodpag = ""
        sw_ayuda = True
        ayuda_formapago.Show 1
'        ayu_f_p_c.Show 1
        sw_ayuda = False
        If Len(wcodpag) > 0 Then
            txtcodforma = wcodpag
            pnlnomforma = wnompag
            txtcodforma_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtcodforma_LostFocus()

'    If sw_ayuda = False Then
'        If Len(Trim(txtcodforma.Text)) > 0 Then
'            If rst.State = adStateOpen Then rst.Close
'            rst.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG='" & Trim(txtcodforma.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not rst.EOF Then
'                pnlnomforma.Caption = Trim("" & rst!F2DESPAG)
'            Else
'                pnlnomforma.Caption = ""
'                MsgBox "Còdigo de forma de pago no existe. Verifique.", vbInformation, "Atenciòn"
'                txtcodforma.SetFocus
'            End If
'            rst.Close
'        End If
'    End If
    
    If Trim(txtcodforma.Text) <> vbNullString And Trim(pnlnomforma.Caption) = vbNullString Then
        txtcodforma.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2FORPAG", "EF2FORPAG", "F2FORPAG", Trim(txtcodforma.Text), "T")
        pnlnomforma.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", Trim(txtcodforma.Text), "T")
    End If
End Sub

Private Sub txtcodsoli_Change()
    'txtcodsoli_LostFocus
End Sub

Private Sub txtcodsoli_DblClick()
    Call txtcodsoli_KeyDown(113, 0)
End Sub

Private Sub txtcodsoli_GotFocus()

    If Len(Trim(txtcodsoli.Text)) = 0 Then
        txtcodsoli.Text = wusuario
    End If

End Sub

Private Sub txtcodsoli_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 113
    wcodusuario = ""
    ayuda_usuarios.Show 1
    If wcodusuario <> "" Then
        txtcodsoli.Text = wcodusuario
        Me.pnlnomsoli.Caption = UCase(wnomusuario)
    End If
End Select
End Sub

Private Sub txtcodsoli_LostFocus()

    If Trim(txtcodsoli.Text) <> vbNullString Then
'        If rst.State = adStateOpen Then rst.Close
'        rst.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(txtcodsoli.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'        If Not rst.EOF Then
'            pnlnomsoli.Caption = "" & rst.Fields("f2nomuser")
'        Else
'            pnlnomsoli.Caption = ""
'            MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
'            'txtcodsoli.SetFocus
'        End If
'        rst.Close
        pnlnomsoli.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", Trim(txtcodsoli.Text), "T")
    End If

End Sub

Private Sub txtcontacto_GotFocus()
    txtcontacto.SelStart = 0: txtcontacto.SelLength = Len(txtcontacto)

End Sub

Private Sub txtContacto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub txtCotizacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
End Sub

Private Sub txtempresa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub txtlugar_entrega_DblClick()
    txtlugar_entrega_KeyDown vbKeyF2, 0
End Sub

Private Sub txtlugar_entrega_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_almacen
                If ModUtilitario.validarFormAbierto("ayuda_almacen") Then
                    Unload ayuda_almacen
                End If
                
                wcod_alm = vbNullString
                
                .Show 1
                
                If wcod_alm <> vbNullString Then
                    txtlugar_entrega.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRALM", "EF2ALMACENES", "F2CODALM", wcod_alm, "T")
                End If
            End With
    End Select
End Sub


Private Sub txtlugar_entrega_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Grid.Columns.FocusedIndex = 2
            
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
        Case Else
            KeyAscii = ModUtilitario.validarCajaAlfaNumerica(KeyAscii)
    End Select
End Sub

Private Sub txtobserva_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    Else
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
    
End Sub

Private Sub Txt_Fecha_Change()
    
    wgraba = 0
    If Not inicio Then swGrabacion = True
    'txt_tc.Text = Format(ObtenerCampo("CAMBIOS", "CAMBIO", "FECHA", txt_fecha.Value, "F", cnn_dbbancos), "0.000")
    txt_tc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", Trim(txt_fecha.value), "F")), "0.000")

End Sub

Private Sub Txt_Fecha_GotFocus()
    
    'txt_fecha.FocusSelect = True
    
End Sub

Private Sub Txt_Fecha_KeyPress(KeyAscii As Integer)
    
'    If KeyAscii = 13 Then
'        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'    End If
    
End Sub

Private Sub Txt_NumOC_KeyPress(KeyAscii As Integer)
    
Select Case KeyAscii
Case 13
    
        'Txt_NumOC.Text = Format(Txt_NumOC.Text, "0000000")
        If Len(Txt_NumOC.Text) > 0 Then
            flagwin = True
            Wnuevo = False
            sw_nuevo_documento = False
            GOC = Trim(Txt_NumOC.Text)
            MODIFICAR_OC
            If ExisteOrdenCompra Then
                ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
            Else
                MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " no existe", vbInformation, "Sistema de Logística"
                Txt_NumOC.SetFocus
            End If
        End If
Case 65 To 90
    KeyAscii = 0
End Select
    
End Sub

Private Sub Txt_NumSolComp_DblClick()

    Call Txt_NumSolComp_KeyDown(113, 0)
        atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
        atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
        atbmenu.Tools.ITEM("ID_Email").Enabled = True
        atbmenu.Tools.ITEM("ID_Anular").Enabled = True
    
End Sub

Private Sub Txt_NumSolComp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        FlagAcceso = False
        flagwin = True
        whelp_solicitud = 4
        FlagAcceso = False
    
        'hlp_solicitudes.Show vbModal
        num_solcomp = ""
        ayuda_solicitudes_OC.Show 1
        
        If Len(Trim(num_solcomp)) > 0 Then
'            Txt_NumSolComp = num_solcomp
            Txt_Prove.Enabled = True
        
            Call importarDatosRequerimiento
            'Txt_Prove.Text = ""
            'pnlnomprv.Caption = ""
            'pnldireprv.Caption = ""
            Grid.Dataset.ADODataset.Requery
            txt_fecha.SetFocus
            
        End If
    End If
    
End Sub

Private Sub Txt_NumSolComp_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        num_solcomp = Txt_NumSolComp.Text
        
        If num_solcomp <> "" Then
        Txt_Prove.Enabled = True
        Call importarDatosRequerimiento
        Txt_Prove.Text = ""
        pnlnomprv.Caption = ""
        pnldireprv.Caption = ""
        
        End If
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
        atbmenu.Tools.ITEM("ID_Grabar").Visible = True
        'atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
        'atbmenu.Tools.ITEM("ID_Email").Enabled = True
        'atbmenu.Tools.ITEM("ID_Anular").Enabled = True
    End If
    
End Sub

Sub MostrarDatosOC()
Dim sw_nuevo_temp   As Boolean
Dim sql             As String
Dim i               As Integer
    
    If Loc = 1 Then
        With rsOrdenCab
            If Not (.EOF) Then
                txtempresa = !F4EMPRESA & ""
                If Txt_NumOC = "" Then
                    !F4NUMORD = " "
                Else
                    Txt_NumOC = (!F4NUMORD & "")
                End If
                Txt_NumSolComp = !F4CODSOLICITUD & ""
                txt_fecha.value = !F4FECEMI
                txtobserva.Text = rsOrdenCab!F4OBSERVA & ""
                txtcontacto.Text = rsOrdenCab!F4CONTACTO & ""
                If !F4TIPMON = "S" Then
                    Cmbmone.ListIndex = 0
                Else
                    Cmbmone.ListIndex = 1
                End If
                If !F4PAGOPARCIAL = 1 Then
                    Chk_pagoparcial.Checked = True
                Else
                    Chk_pagoparcial.Checked = False
                End If
                txt_tc = Format$(!F4TIPCAM, "0.000") & ""
                txtcodforma = !F4FORPAG & ""
                Txt_Referencia = !F4REFERE & ""
                txtcodsoli = !F4CODSOL & ""
                txtCotizacion.Text = !F4NUMCOTIZA & ""
                TxtCodCosto.Text = !F4CENTRO & ""
                'txtcodcosto_KeyPress 13
                abofechaentrega.value = Format(!F4FECENT, "DD/MM/YYYY")
                
                If Loc = 2 Then
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                Else
                    txtigv = Format$(!F4IGV & "", "#,##0.00")
                    txtmonto = Format$(!F4MONINA & "", "#,##0.00")
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                    TxtRnd = Format$(!F4RND & "", "#,##0.00")
                    txttotal = Format$(!F4MONTO & "", "#,##0.00")
                End If
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV from EF2PROVEEDORES where F2newruc='" & !F4CODPRV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    Txt_Prove.Text = "" & rst!F2NEWRUC
                    pnlnomprv.Caption = rst!F2NOMPROV
                    pnldireprv.Caption = IIf(IsNull(rst!F2DIRPROV), " ", rst!F2DIRPROV)
                    wgraba = 0
                Else
                    pnlnomprv.Caption = "Ruc es menor a 11 digitos"
                    pnldireprv.Caption = "No tiene "
                End If
                rst.Close
                
                xnombre = rsOrdenCab!F4CODSOL
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2NOMUSER from ef2userS where f2coduser='" & UCase(Trim(xnombre)) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    txtcodsoli = UCase(xnombre)
                    pnlnomsoli.Caption = rst!F2NOMUSER & ""
                End If
                rst.Close
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2DESPAG from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    pnlnomforma.Caption = "" & rst.Fields("F2DESPAG")
                    wgraba = 0
                End If
                rst.Close
                
                
                txtlugar_entrega.Text = left(Trim("" & !F4LUGAR_ENTREGA), 100)
        
            Else
                MsgBox "La Solicitud de Compra no existe", vbInformation, "Atención"
                Txt_NumSolComp.Enabled = True
                Txt_NumSolComp.SetFocus
                Exit Sub
            End If
        End With
    Else
    End If
          
    With rsOrdenDet
        sql = "SELECT * from if3orden where f4numord='" & GOC & "' AND F4local = '" & TOC & "'"
        If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
        rsOrdenDet.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (.EOF) Then
            existe = True
            If sw_nuevo_documento = False Then
                DELETEREC_LOG cnomtabla, cnn_form
                AdicionaItem
                sw_nuevo_documento = True
            End If
            
            dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
            dxDBGrid1.Dataset.Active = True
        
            dxDBGrid1.Dataset.Close
            dxDBGrid1.Dataset.Open
            dxDBGrid1.OptionEnabled = False
            dxDBGrid1.Dataset.DisableControls
            'dxDBGrid1.Dataset.EnableControls
            'dxDBGrid1.Dataset.Close
            'dxDBGrid1.Dataset.Open
            'dxDBGrid1.OptionEnabled = True
            
            sw_nuevo_temp = False
            sw_nuevo_item = True
            
            .MoveFirst
            
            Do While Not .EOF
                i = i + 1
                If Loc = 1 Then
                    If rsOrdenDet.Fields("f4numord") = GOC Then
                        If sw_nuevo_temp = False Then
                            If sw_nuevo_documento = True Then
                                dxDBGrid1.Dataset.Edit
                            Else
                                dxDBGrid1.Dataset.Append
                            End If
                            sw_nuevo_temp = True
                        Else
                            dxDBGrid1.Dataset.Append
                        End If
                
                        dxDBGrid1.Dataset.FieldValues("item") = i
                        dxDBGrid1.Dataset.FieldValues("f3codpro") = .Fields("f3codpro") & ""
                        dxDBGrid1.Dataset.FieldValues("f3cencos") = .Fields("F3CENCOS") & ""
                        If rst.State = adStateOpen Then rst.Close
                        rst.Open "SELECT P.f5nompro,P.f5codfab,P.F7codmed,M.F2DESMAR from if5pla P, EF2MARCAS M where P.f5codpro='" & rsOrdenDet!F3CODPRO & "' AND P.F5MARCA=M.F2CODMAR", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not (rst.EOF) Then
                            dxDBGrid1.Dataset.FieldValues("f5nompro") = rst.Fields("f5nompro") & ""
                            dxDBGrid1.Dataset.FieldValues("f5codfab") = rst!f5codfab & ""
                            dxDBGrid1.Dataset.FieldValues("f3medida") = rst!f7codmed & ""
                            dxDBGrid1.Dataset.FieldValues("f5marca") = rst!f2desmar & ""
                            
                        End If
                        
                        rst.Close
                        
                        If Len(Trim(.Fields("f5nompro") & "")) > 0 Then
                            dxDBGrid1.Dataset.FieldValues("f5nompro") = Trim(.Fields("f5nompro") & "")
                        End If
                            
                        dxDBGrid1.Dataset.FieldValues("f3canpro") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("f3redondeo") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("f3precos") = Format$(Val("" & .Fields("f3precos")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3pordct") = Format$(Val("" & .Fields("f3pordct")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3totdct") = Format$(Val("" & .Fields("f3totdct")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f5valvta") = Format$(Val("" & .Fields("f5valvta")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f5afecto") = .Fields("f5afecto")
                        dxDBGrid1.Dataset.FieldValues("f3igv") = Format$(Val("" & .Fields("f3igv")), "#,##0.00")
                        dxDBGrid1.Dataset.FieldValues("f3preuni") = Format$(Val("" & .Fields("f3preuni")), "#,##0.0000")
                        dxDBGrid1.Dataset.FieldValues("f3total") = Format$(Val("" & .Fields("f3total")), "###,##0.00")
                        If Not (IsDate(rsOrdenDet!f3fentrega)) Then
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.value, "dd/mm/yyyy"))
                        Else
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = Format(rsOrdenDet!f3fentrega, "dd/mm/yyyy")
                        End If
                        dxDBGrid1.Dataset.FieldValues("check") = True
                        dxDBGrid1.Dataset.FieldValues("cant_ant") = .Fields("f3canpro")
                        dxDBGrid1.Dataset.FieldValues("cod_solicitud") = .Fields("cod_solicitud")
                    Else
                        Exit Do
                    End If
                    
                End If
                .MoveNext
            Loop
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = False
            jc = 1
        End If
        rsOrdenDet.Close
    End With
    
    If existe = True Then
       dxDBGrid1.Dataset.EnableControls
       dxDBGrid1.Dataset.Open
       dxDBGrid1.OptionEnabled = True
       existe = False
    Else
       Exit Sub
    End If
End Sub
Sub MostrarDatosOC_Grid()
    Dim Amov(0 To 40) As a_grabacion
    Dim RsMed As New ADODB.Recordset
    Dim RsCC As New ADODB.Recordset
    Dim sw_nuevo_temp   As Boolean
    Dim sql             As String
    Dim i               As Integer
    Dim RsU As New ADODB.Recordset
    
    Set RsMed = Af.OpenSQLForwardOnly("select * from ef7medidas", cconex_dbbancos)

    Set RsCC = Af.OpenSQLForwardOnly("select * from centros", cconex_dbbancos)
    
    'If loc = 1 Then
    '    With rsOrdenCab
    '        If Not (.EOF) Then
        With objAyudaOrden
            .inicializarEntidades
            
            .TipoOrden = strTipoOrden
            .NumeroOrden = strNumeroOrden
            
            .obtenerConfigOrden
            
            Txt_TOC.Text = .TipoOrden
            Txt_NumOC.Text = .NumeroOrden
            
            txtempresa = .Empresa ' !F4EMPRESA & ""
'            If Txt_NumOC = "" Then
'                !F4NUMORD = " "
'            Else
'                Txt_NumOC = (!F4NUMORD & "")
'            End If

            
            'Txt_NumSolComp = !F4CODSOLICITUD & ""
            
            txt_fecha.value = .FechaEmision  '!F4FECEMI
            txtobserva.Text = .Observacion 'rsOrdenCab!F4OBSERVA & ""
            txtcontacto.Text = .ContactoProveedor 'rsOrdenCab!F4CONTACTO & ""
            
'            If !F4TIPMON = "S" Then
'                Cmbmone.ListIndex = 0
'            Else
'                Cmbmone.ListIndex = 1
'            End If
            
            Cmbmone.ListIndex = ModUtilitario.seleccionarItem(Cmbmone, .CodMoneda, "IZQ", "1")
            
            FrameOC.FontBold = True

            atbmenu.Tools.ITEM("ID_Nuevo").Visible = True
            atbmenu.Tools.ITEM("ID_Grabar").Visible = True
            atbmenu.Tools.ITEM("ID_Anular").Visible = True
            atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
            atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
            
            txt_tc.Text = Format(.TipoCambio, "0.000")   'Format$(!F4TIPCAM, "0.000") & ""
            txtcodforma = .CodFormaPago  '!F4FORPAG & ""
                pnlnomforma.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", .CodFormaPago, "T")
                
            Txt_Referencia = .referencia  '!F4REFERE & ""
            txtcodsoli.Text = .CodigoSolicitante '!F4CODSOL & ""
                pnlnomsoli.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", .CodigoSolicitante, "T")
                
            txtCotizacion.Text = .NumeroCotizacion '!F4NUMCOTIZA & ""
            TxtCodCosto.Text = .CentroCosto '!F4CENTRO & ""
                PnlNomCosto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", .CentroCosto, "T")
                
'            If Not IsNull(!F4FECENT) Then
'                abofechaentrega.Value = Format(!F4FECENT, "DD/MM/YYYY")
'            Else
'                abofechaentrega.CheckBox = True
'                abofechaentrega.Value = Empty
'            End If
            
            If .FechaEntrega <> vbNullString Then
                abofechaentrega.value = Format(.FechaEntrega, "Short Date")
            Else
                abofechaentrega.CheckBox = True
                abofechaentrega.value = Empty
            End If
            
            'aBoHoraEntrega.Value = Format(!F4FECENT, "hh:mm:ss")
            'txtFechaPago.Text = !F4DIAPAGO & ""
            
            ChK_regularizacion.Checked = .OrdenRegularizada
            
'            If !F4REGULARIZA = "1" Then
'                ChK_regularizacion.Checked = True
'            Else
'                ChK_regularizacion.Checked = False
'            End If
            
            Chk_pagoparcial.Checked = .PagoParcial
            
'            If !F4PAGOPARCIAL = True Then
'                Chk_pagoparcial.Checked = True
'            Else
'                Chk_pagoparcial.Checked = False
'            End If
            
            'If loc = 2 Then
            '    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
            'Else
                txtbase.Text = Format(.SUBTOTAL, "#,##0.00")
                txtmonto.Text = Format(.TotalInafecto, "#,##0.00")
                txtigv.Text = Format(.TotalImpuesto, "#,##0.00")
                TxtRnd.Text = Format(.TotalRedondeo & "", "#,##0.00")
                txttotal.Text = Format(.TotalFacturado & "", "#,##0.00")
            'End If
            
            'If rst.State = adStateOpen Then rst.Close
            'rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV from EF2PROVEEDORES where F2newruc='" & !F4CODPRV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            'If Not (rst.EOF) Then
            '    Txt_Prove.Text = "" & rst!F2NEWRUC
            '    pnlnomprv.Caption = rst!F2NOMPROV
            '    pnldireprv.Caption = IIf(IsNull(rst!F2DIRPROV), " ", rst!F2DIRPROV)
            '    wgraba = 0
            'Else
            '    pnlnomprv.Caption = "Ruc es menor a 11 digitos"
            '    pnldireprv.Caption = "No tiene "
            'End If
            'rst.Close
            
            pnlnomprv.Caption = "Ruc es menor a 11 digitos"
            pnldireprv.Caption = "No tiene "
            
            If .RucProveedor <> vbNullString Then
                Txt_Prove.Text = .RucProveedor
                pnlnomprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2NEWRUC", .RucProveedor, "T")
                pnldireprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2NEWRUC", .RucProveedor, "T")
            End If
            
'            xnombre = rsOrdenCab!F4CODSOL & ""
'            csql = "SELECT F2NOMUSER from ef2userS where f2coduser='" & UCase(Trim(xnombre)) & "'"
'            Set RsU = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
'            If Not (RsU.EOF) Then
'                txtcodsoli = UCase(xnombre)
'                pnlnomsoli.Caption = RsU!F2NOMUSER & ""
'            End If
'            RsU.Close
            
'            If rst.State = adStateOpen Then rst.Close
'            rst.Open "SELECT F2DESPAG from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not (rst.EOF) Then
'                pnlnomforma.Caption = "" & rst.Fields("F2DESPAG")
'                wgraba = 0
'            End If
'            rst.Close
            
            'SeleccionaEnComboRight rsOrdenCab!F4TIPDOC & "", CmbTipDoc
            CmbTipDoc.ListIndex = ModUtilitario.seleccionarItem(CmbTipDoc, .CodTipoComprobante, "DER", 2)
            
            txtlugar_entrega.Text = .LugarEntrega 'left(Trim("" & !F4LUGAR_ENTREGA), 100)
            
            Rem SK ADD:
            fraSeguimiento.Enabled = True
            chkOrdenEnviada.value = IIf(.Colocada, vbChecked, vbUnchecked)
            
            If .Colocada Then
                txtEnviadoPor.Text = .ColocadaUsuario 'Trim(rsOrdenCab!F4COLOCADAUSER & "")
                dtpFechaEnvio.value = .ColocadaFecha 'Format(Trim(rsOrdenCab!F4COLOCADAFECHA & ""), "Short Date")
            End If
            
            chkOrdenRecepcionada.value = IIf(.Atendida, vbChecked, vbUnchecked)
            
            If .Atendida Then
                txtRecepcionadoPor.Text = .AtendidaUsuario ' Trim(rsOrdenCab!F4ATENDIDAUSER & "")
                dtpFechaRecepcion.value = .AtendidaFecha ' Format(Trim(rsOrdenCab!F4ATENDIDAFECHA & ""), "Short Date")
            End If
            
            chkOrdenEnviada_Click
            chkOrdenRecepcionada_Click
            'Else
            '    MsgBox "La Solicitud de Compra no existe", vbInformation, "Atención"
            '    Txt_NumSolComp.Enabled = True
            '    Txt_NumSolComp.SetFocus
            '
            '    fraSeguimiento.Enabled = False
            '
            '    Exit Sub
        '    End If
        End With
    'Else
    'End If
    
    With rsOrdenDet
        sql = "SELECT * from if3orden where f4numord='" & strNumeroOrden & "' AND F4local = '" & strTipoOrden & "' ORDER BY val(ITEM)"
        If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
        rsOrdenDet.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (.EOF) Then
            existe = True
            If sw_nuevo_documento = False Then
                DELETEREC_LOG cnomtabla, CnTmp
                'AdicionaItem
                sw_nuevo_documento = True
            End If
            
            sw_nuevo_temp = False
            sw_nuevo_item = True
            
            .MoveFirst
            
            Grid.Dataset.Edit
            
            Do While Not .EOF
                i = i + 1
                
                If Loc = 1 Then
                    If rsOrdenDet.Fields("f4numord") = strNumeroOrden Then
                        Amov(0).campo = "item": Amov(0).valor = i & "": Amov(0).Tipo = "N"
                        Amov(1).campo = "f3codpro": Amov(1).valor = .Fields("f3codpro") & "": Amov(1).Tipo = "T"
                        Amov(2).campo = "f5codcosto": Amov(2).valor = .Fields("F3CENCOS") & "": Amov(2).Tipo = "T"
                        
                        RsCC.Filter = adFilterNone
                        RsCC.Filter = "f3costo='" & .Fields("F3CENCOS") & "" & "'"
                        If RsCC.RecordCount > 0 Then
                            Amov(15).campo = "f5descosto": Amov(15).valor = RsCC!F3ABREV & "": Amov(15).Tipo = "T"
                        Else
                            Amov(15).campo = "f5descosto": Amov(15).valor = "": Amov(15).Tipo = "T"
                        End If
                        'If rst.State = adStateOpen Then rst.Close
                        'rst.Open "SELECT P.f5nompro,P.f5codfab,P.F7codmed,M.F2DESMAR from if5pla P, EF2MARCAS M where P.f5codpro='" & rsOrdenDet!f3codpro & "' AND P.F5MARCA=M.F2CODMAR", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        'If Not (rst.EOF) Then
                         '   Amov(3).campo = "f5nompro": Amov(3).valor = rst.Fields("f5nompro") & "": Amov(3).TIPO = "T"
                          '  Amov(4).campo = "f3codmedida": Amov(4).valor = rst!f7codmed & "": Amov(4).TIPO = "T"
                        'Else
                        Amov(3).campo = "f5nompro": Amov(3).valor = .Fields("f5nompro") & "": Amov(3).Tipo = "T"
                        Amov(4).campo = "f3codmedida": Amov(4).valor = .Fields("UNIDAD") & "": Amov(4).Tipo = "T"
                        RsMed.Filter = adFilterNone
                        RsMed.Filter = "f7codmed='" & .Fields("UNIDAD") & "" & "'"
                        If RsMed.RecordCount > 0 Then
                            Amov(14).campo = "f3desmedida": Amov(14).valor = RsMed!F7SIGMED & "": Amov(14).Tipo = "T"
                        Else
                            Amov(14).campo = "f3desmedida": Amov(14).valor = "": Amov(14).Tipo = "T"
                        End If
                        'End If
                        'rst.Close
                                                                           
                        Amov(5).campo = "f3canpro": Amov(5).valor = .Fields("f3canpro") & "": Amov(5).Tipo = "N"
                        Amov(6).campo = "f3sinigv": Amov(6).valor = Val("" & .Fields("f3precos")): Amov(6).Tipo = "N"
                        Amov(7).campo = "f3conigv": Amov(7).valor = Val("" & .Fields("f3preuni")): Amov(7).Tipo = "N"
                        Amov(8).campo = "f5afecto": Amov(8).valor = IIf(.Fields("f5afecto") & "" = "*", -1, 0): Amov(8).Tipo = "N"
                                                                           
                        If Trim(.Fields("f5afecto") & "") = "*" Then
                            Amov(9).campo = "f3baseimp": Amov(9).valor = Val("" & .Fields("f5valvta")): Amov(9).Tipo = "N"
                            Amov(10).campo = "f3monina": Amov(10).valor = "0": Amov(10).Tipo = "N"
                        Else
                            Grid.Dataset.FieldValues("f3baseimp") = 0
                            Grid.Dataset.FieldValues("f3monina") = Val("" & .Fields("f5valvta"))
                            Amov(9).campo = "f3baseimp": Amov(9).valor = 0: Amov(9).Tipo = "N"
                            Amov(10).campo = "f3monina": Amov(10).valor = Val("" & .Fields("f5valvta")): Amov(10).Tipo = "N"
                        End If
                        
                        Amov(11).campo = "f3igv": Amov(11).valor = Val("" & .Fields("f3igv")): Amov(11).Tipo = "N"
                        Amov(12).campo = "f3total": Amov(12).valor = Val("" & .Fields("f3total")): Amov(12).Tipo = "N"
                        Amov(13).campo = "f3colmod": Amov(13).valor = ("" & .Fields("f3backorder")): Amov(13).Tipo = "T"
                        Amov(16).campo = "f3valdesc": Amov(16).valor = Val("" & .Fields("f3totdct")): Amov(16).Tipo = "N"
                        Amov(17).campo = "f3pordesc": Amov(17).valor = Val("" & .Fields("f3pordct")): Amov(17).Tipo = "N"
                        Amov(18).campo = "f3observa": Amov(18).valor = "" & .Fields("f3observa"): Amov(18).Tipo = "T"
                        Amov(19).campo = "cod_solicitud": Amov(19).valor = "" & .Fields("cod_solicitud"): Amov(19).Tipo = "T"
                        
                        Rem SK ADD:
                        Amov(20).campo = "CODCOLOR": Amov(20).valor = Trim(.Fields("CODCOLOR") & ""): Amov(20).Tipo = "T"
                        Amov(21).campo = "DESCOLOR": Amov(21).valor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "DESCRIPCION", "EF2BIENCOLOR", "CODIGO", Trim(.Fields("CODCOLOR") & ""), "T"): Amov(21).Tipo = "T"
                        Amov(22).campo = "F3CODFAB": Amov(22).valor = Trim(.Fields("F3CODFAB") & ""): Amov(22).Tipo = "T"
                        Amov(23).campo = "F3PORCDEMASIA": Amov(23).valor = Val(.Fields("F3PORCDEMASIA") & ""): Amov(23).Tipo = "N"
                        Amov(24).campo = "F3CANPROMAX": Amov(24).valor = .Fields("F3CANPRO2") & "": Amov(24).Tipo = "N"
                        Amov(25).campo = "F3NETO": Amov(25).valor = .Fields("F3PRENETO") & "": Amov(25).Tipo = "N"
                        Amov(26).campo = "f5nompro_ING": Amov(26).valor = .Fields("f5nompro_ING") & "": Amov(26).Tipo = "T"
                        Amov(27).campo = "F3CANPROFINAL": Amov(27).valor = Val(.Fields("f3canpro") & "") * (1 + (Val(.Fields("F3PORCDEMASIA") & "") / 100)): Amov(27).Tipo = "N"
                    Else
                        Exit Do
                    End If
                End If
                
                GRABA_REGISTRO_noenvia Amov, "TMPORDENDECOMPRA", "A", 27, CnTmp, ""
                
                .MoveNext
            Loop
            Grid.Dataset.Post
            sw_nuevo_item = False
            jc = 1
        End If
        rsOrdenDet.Close
    End With
    If RsMed.State = 1 Then RsMed.Close
    Set RsMed = Nothing
    If RsCC.State = 1 Then RsCC.Close
    Set RsCC = Nothing
    
    If CnTmp.State = 1 Then CnTmp.Close
    CnTmp.Open StrCn
    'Grid.Dataset.Close
    Grid.Dataset.Active = False
    Grid.Dataset.ADODataset.ConnectionString = CnTmp
    Grid.Dataset.ADODataset.CommandText = "select * from TMPORDENDECOMPRA order by item"
    Grid.Dataset.Active = True
    dxDBGrid1.KeyField = "item"
    Grid.Dataset.Close
    Grid.Dataset.Open
    If Grid.Dataset.RecordCount = 0 Then
        AdicionaItemGrid
    End If
    
End Sub

Private Sub listarGrilla()
    abrirCnTemporal
    
    With Grid.Dataset
        .Active = False
        .ADODataset.ConnectionString = cnDBTemp  'CnTmp
        .ADODataset.CommandText = "select * from TMPORDENDECOMPRA order by item"
        .Active = True
        
        Grid.KeyField = "ITEM"
        
        .Close
        .Open
    End With
End Sub

Private Sub CONFIGURA_GRID()
    
    With dxDBGrid1.Options
'        .Set (egoEditing)
'        .Set (egoTabs)
'        .Set (egoTabThrough)
'        .Set (egoCanDelete)
'        .Set (egoCanAppend)
'        .Set (egoCanInsert)
'        .Set (egoImmediateEditor)
'        '.Set (egoShowIndicator)
'        .Set (egoCanNavigation)
'        .Set (egoHorzThrough)
'        .Set (egoVertThrough)
'        .Set (egoAutoWidth)
'        .Set (egoEnterShowEditor)
'        .Set (egoEnterThrough)
'        .Set (egoShowButtonAlways)
'
'        .Set (egoColumnSizing)
'        .Set (egoColumnMoving)
'        .Set (egoTabThrough)
'        .Set (egoConfirmDelete)
'        .Set (egoCanNavigation)
'        .Set (egoCancelOnExit)
'        .Set (egoLoadAllRecords)
'        .Set (egoShowHourGlass)
'        .Set (egoUseBookmarks)
'        .Set (egoUseLocate)
'        .Set (egoAutoCalcPreviewLines)
'        .Set (egoBandSizing)
'        .Set (egoBandMoving)
'        .Set (egoDragScroll)
'        .Set (egoExpandOnDblClick)
'        .Set (egoShowFooter)
'        .Set (egoShowGrid)
'        .Set (egoShowButtons)
'        .Set (egoNameCaseInsensitive)
'        .Set (egoShowHeader)
'        .Set (egoShowPreviewGrid)
'        .Set (egoShowBorder)
'        .Set (egoDynamicLoad)


        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
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
         .Set (egoShowGrid)
         .Set (egoShowButtons)
         .Set (egoNameCaseInsensitive)
         .Set (egoShowHeader)
         .Set (egoShowPreviewGrid)
         .Set (egoShowBorder)
         .Set (egoDynamicLoad)

    End With
    dxDBGrid1.Columns(1).Visible = False
    dxDBGrid1.Columns.ColumnByFieldName("f3precos").Caption = "P. Costo"
    dxDBGrid1.Columns.ColumnByFieldName("f3precos").DecimalPlaces = 3
    If wf1visualiza_dctos = "*" Then
        dxDBGrid1.Columns.ColumnByFieldName("f3pordct").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("f3totdct").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("f5valvta").Visible = False
    End If
    
End Sub


Private Sub configuraGrilla()
    With Grid.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        '.Set (egoAutoWidth)
        
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
         .Set (egoShowGrid)
         .Set (egoShowButtons)
         .Set (egoNameCaseInsensitive)
         .Set (egoShowHeader)
         .Set (egoShowPreviewGrid)
         .Set (egoShowBorder)
         .Set (egoDynamicLoad)
    End With
End Sub


Private Function Nueva_orden() As String
Dim sql     As String
Dim Orden   As String
Dim strCodCosto As String

    strCodCosto = ObtenerCampo("centros", "cconcar", "f3costo", TxtCodCosto.Text, "T", cnn_dbbancos)
    If wTipoOC = "1" Then
        TOC = "OC"
        sql = "SELECT top 1 RIGHT(F4NUMORD,10) AS MAYOR FROM IF4ORDEN WHERE F4LOCAL='OC' AND LEFT(F4NUMORD,2) = 'OC' order by RIGHT(F4NUMORD,10) desc" 'and right(f4numord,3) = '" & StrCodCosto & "'
    Else
        TOC = "OS"
        sql = "SELECT top 1 RIGHT(F4NUMORD,10) AS MAYOR FROM IF4ORDEN WHERE F4LOCAL='OS'  AND LEFT(F4NUMORD,2) = 'OS' order by RIGHT(F4NUMORD,10) desc" 'and right(f4numord,3) = '" & StrCodCosto & "'
    End If
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not (rst.EOF) Then
        'Orden = rst.Fields("f4numord") + 1
        'Orden = wanno & "-" & Format(Mid(rst.Fields("f4numord"), 6, 5) + 1, "00000") & "/0"
        wmes = Month(txt_fecha.value)
        Orden = Format(Val(rst.Fields("MAYOR") & "") + 1, "000000000000") ' & "/" & StrCodCosto
    Else
        'Orden = wanno & "-" & Format(1, "00000") & "/0"
        wmes = Month(txt_fecha.value)
        Orden = "000000000001" '& StrCodCosto
    End If
    If wTipoOC = "1" Then
        Nueva_orden = "OC" & right(Orden, 10)
    Else
        Nueva_orden = "OS" & right(Orden, 10)
    End If
    
End Function

Sub GrabarOC()
    On Error GoTo CapturaError
    
    Dim codi                As String
    Dim wcantidad           As Double
    Dim wcc                 As String
    Dim wproducto           As String
    Dim sql                 As String
    Dim ocompra             As Double
    Dim Cant                As Double
    Dim rsdetaoc            As New ADODB.Recordset
    Dim ncant_ant           As Double
    'Dim amovs_cab(0 To 31)  As a_grabacion
    Dim amovs_cab(0 To 37)  As a_grabacion
    Dim ctipo               As String
    Dim StrMsg              As String
    Dim RsPre               As ADODB.Recordset
    Dim RsCom               As ADODB.Recordset
    Dim dblCanPre As Double, dblTotPre As Double
    Dim dblCanCom As Double, dblTotCom As Double
    Dim soli_acum As String
    
    flag = 0
    
    If Trim(Txt_NumOC.Text) <> "" Then
        jc = 1
    Else
        jc = 0
    End If
    
    If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
        Grid.Dataset.Post
    End If
    
    Calcula_PvtaTotalGrid
    
    sql = "select sum(f3total) as ztotal from TMPORDENDECOMPRA "
    
    Set rst = Af.OpenSQLForwardOnly(sql, StrCn)
    
    If rst.RecordCount > 0 Then
        If Val(rst!ztotal & "") < 0 Then
            MsgBox "Debe Ingresar y/o Seleccionar Productos a Comprar", vbInformation, "Sistema de Logística"
            Grid.SetFocus
            Exit Sub
        End If
        If rst.State = 1 Then rst.Close
        
    End If
    'If jc = 0 Then
    'Verifica con Presupuesto
        
'        For I = 1 To Grid.Dataset.RecordCount
'            Grid.Dataset.RecNo = I
            'csql = "SELECT Sum(REGISMOV.F3CANTIDAD) AS SumaDeCANTIDAD, Sum(REGISMOV.F3IMPORTE) AS SumaDeIMPORTE "
            'csql = csql & "From REGISMOV "
            'csql = csql & "WHERE REGISMOV.F5CODPRO='" & Grid.Columns.ColumnByFieldName("f3codpro").Value & "' "
            'If TxtCodCosto.Text = "998" Then
            '    csql = csql & "AND REGISMOV.F3CENCOS='" & Grid.Columns.ColumnByFieldName("f5codcosto").Value & "'"
            'Else
            '    csql = csql & "AND REGISMOV.F3CENCOS='" & TxtCodCosto.Text & "'"
            'End If
'            csql = "SELECT Sum((REGISMOV.F3CANTIDAD)) AS SUMADECANTIDAD, "
'            If Cmbmone.ListIndex = 0 Then
'                csql = csql & "Sum(IIf(REGISDOC.F4MONEDA='D',REGISMOV.F3IMPORTE*REGISDOC.F4TIPCAM,REGISMOV.F3IMPORTE)) AS SUMADEIMPORTE "
'            Else
'                csql = csql & "Sum(IIf(REGISDOC.F4MONEDA='S',REGISMOV.F3IMPORTE/REGISDOC.F4TIPCAM,REGISMOV.F3IMPORTE)) AS SUMADEIMPORTE "
'            End If
'            csql = csql & "FROM REGISDOC INNER JOIN REGISMOV ON (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) AND (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) "
'            csql = csql & "WHERE REGISMOV.F5CODPRO='" & Grid.Columns.ColumnByFieldName("f3codpro").Value & "' "
'            If TxtCodCosto.Text = "998" Then
'                csql = csql & "AND REGISMOV.F3CENCOS='" & Grid.Columns.ColumnByFieldName("f5codcosto").Value & "'"
'            Else
'                csql = csql & "AND REGISMOV.F3CENCOS='" & TxtCodCosto.Text & "'"
'            End If
'
'            Set RsCom = Af.OpenSQLForwardOnly(csql, cconex_dbbancos) 'suma lo comprado
'            dblCanCom = 0: dblTotCom = 0
'            If RsCom.RecordCount > 0 Then
'                dblCanCom = Val(RsCom!sumadecantidad & ""): dblTotCom = Val(RsCom!sumadeimporte & "")
'            End If
'            csql = "select * from presup_gastos where F5CODPRO='" & Grid.Columns.ColumnByFieldName("f3codpro").Value & "' "
'            If TxtCodCosto.Text = "998" Then
'                csql = csql & "AND F3COSTO='" & Grid.Columns.ColumnByFieldName("f5codcosto").Value & "'"
'            Else
'                csql = csql & "AND F3COSTO='" & TxtCodCosto.Text & "'"
'            End If
'            csql = "SELECT Sum(Presup_Gastos.CANTIDAD) AS SumaDeCANTIDAD, "
'            If Cmbmone.ListIndex = 0 Then
'                csql = csql & "Sum(IIf(Presup_Gastos.MONEDA='D',Presup_Gastos.MONTO*Presup_Gastos.TIPCAM,Presup_Gastos.MONTO)) AS SumadeMONTO "
'            Else
'                csql = csql & "Sum(IIf(Presup_Gastos.MONEDA='S',Presup_Gastos.MONTO/Presup_Gastos.TIPCAM,Presup_Gastos.MONTO)) AS SumadeMONTO "
'            End If
'            csql = csql & "From Presup_Gastos "
'            csql = csql & "WHERE Presup_Gastos.F5CODPRO='" & Grid.Columns.ColumnByFieldName("f3codpro").Value & "' "
'            If TxtCodCosto.Text = "998" Then
'                csql = csql & "AND Presup_Gastos.F3COSTO='" & Grid.Columns.ColumnByFieldName("f5codcosto").Value & "'"
'            Else
'                csql = csql & "AND Presup_Gastos.F3COSTO='" & TxtCodCosto.Text & "'"
'            End If
            
'            Set RsPre = Af.OpenSQLForwardOnly(csql, cconex_dbbancos) 'muestra lo presupuestado
'            dblCanPre = 0: dblTotPre = 0
'            If RsPre.RecordCount > 0 Then
'                dblCanPre = Val(RsPre!sumadecantidad & ""): dblTotPre = Val(RsPre!sumadeMONTO & "")
'            End If
            'comparando
'            If dblTotCom + Val(Grid.Columns.ColumnByFieldName("f3TOTAL").Value & "") > dblTotPre Then
'                MsgBox "Con esta Orden de Compra excedería el TOTAL presupuestado en " & Format(dblTotPre - (dblTotCom + Val(Grid.Columns.ColumnByFieldName("f3total").Value & "")), "###,###,##0.00"), vbExclamation, "Restricción del Sistema"
'                Exit Sub
'            End If
'            If dblCanCom + Val(Grid.Columns.ColumnByFieldName("f3canpro").Value & "") > dblCanPre Then
'                MsgBox "Con esta Orden de Compra excedería la CANTIDAD presupuestada en " & dblCanPre - (dblCanCom + Val(Grid.Columns.ColumnByFieldName("f3canpro").Value & "")), vbExclamation, "Restricción del Sistema"
'                Exit Sub
'            End If
            

'        Next
    'End If
        
        'suma la orden de compra
        
    '************************
    If Txt_Prove = "" Then MsgBox "Ingrese Código de Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
    If pnlnomprv = "" Then MsgBox "Ingrese Nombre de Proveedor", 48, "Sistema de Logística": Txt_Prove.SetFocus: Exit Sub
    If txtcodsoli = "" Then MsgBox "Ingrese Código de solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
    If pnlnomsoli = "" Then MsgBox "Ingrese Nombre de solicitante", 48, "Sistema de Logística": txtcodsoli.SetFocus: Exit Sub
    If txtcodforma = "" Then MsgBox "Ingrese código de forma de pago", 48, "Sistema de Logística": txtcodforma.SetFocus: Exit Sub
    If Cmbmone.ListIndex < 0 Then MsgBox "Seleccione moneda", 48, "Sistema de Logística": Cmbmone.SetFocus: Exit Sub
    If Val(txt_tc.Text) = 0 Then MsgBox "Ingrese Tipo de Cambio", 48, "Sistema de Logística": txt_tc.SetFocus: Exit Sub
    If abofechaentrega = Empty Then MsgBox "Ingrese Fecha de Entrega", 48, "Sistema de Logística": txt_tc.SetFocus: Exit Sub
    
    'Nueva Versión
    If Loc = 1 Then
        Select Case jc
            Case 0
            Txt_NumOC.Text = Nueva_orden
            Txt_TOC.Text = TOC
        End Select
    End If
    
    
    If Loc = 1 Then
        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
        If SwRenovar = True Then
            'wNumOC = Mid(Txt_NumOC.Text, 1, 11) & Val(Mid(Txt_NumOC.Text, 12, 1)) + 1
            wNumOc = Nueva_orden
            rsOrdenCab.Open "SELECT F4ESTNUL,F4FALTA,F4ESTVAL from if4orden where f4numord='" & wNumOc & "' AND F4local = '" & Txt_TOC.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        Else
            rsOrdenCab.Open "SELECT F4ESTNUL,F4FALTA,F4ESTVAL from if4orden where f4numord='" & Txt_NumOC & "' AND F4local = '" & Txt_TOC.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        End If
        If Not (rsOrdenCab.EOF) Then
            ctipo = "M"
            TOC = Txt_TOC.Text
        Else
            ctipo = "A"
            flag = 1
            sw_nuevo_documento = False
        End If
        If SwRenovar = True Then
            amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = wNumOc: amovs_cab(0).Tipo = "T"
        Else
            amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = Txt_NumOC.Text: amovs_cab(0).Tipo = "T"
        End If
        If ctipo = "A" Then
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = "N": amovs_cab(1).Tipo = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = "1": amovs_cab(2).Tipo = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = 0: amovs_cab(3).Tipo = "T"
            amovs_cab(4).campo = "F4FECGRA": amovs_cab(4).valor = Format(Now, "dd/MM/yyyy hh:mm:ss"): amovs_cab(4).Tipo = "F"
            amovs_cab(5).campo = "F4USEGRA": amovs_cab(5).valor = wusuario: amovs_cab(5).Tipo = "T"
        Else
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = rsOrdenCab.Fields("F4ESTNUL"): amovs_cab(1).Tipo = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = rsOrdenCab.Fields("F4FALTA"): amovs_cab(2).Tipo = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = rsOrdenCab.Fields("F4ESTVAL"): amovs_cab(3).Tipo = "T"
            amovs_cab(4).campo = "F4FECMOD": amovs_cab(4).valor = Format(Now, "dd/MM/yyyy hh:mm:ss"): amovs_cab(4).Tipo = "F"
            amovs_cab(5).campo = "F4USEMOD": amovs_cab(5).valor = wusuario: amovs_cab(5).Tipo = "T"
        End If
        
        amovs_cab(6).campo = "F4CODSOL": amovs_cab(6).valor = txtcodsoli.Text: amovs_cab(6).Tipo = "T"
        amovs_cab(7).campo = "F4FECEMI": amovs_cab(7).valor = Format(txt_fecha.value, "DD/MM/YYYY"): amovs_cab(7).Tipo = "F"
        amovs_cab(8).campo = "F4CODPRV": amovs_cab(8).valor = Txt_Prove: amovs_cab(8).Tipo = "T"
        amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = txt_tc.Text: amovs_cab(9).Tipo = "N"
        amovs_cab(10).campo = "F4FORPAG": amovs_cab(10).valor = txtcodforma.Text: amovs_cab(10).Tipo = "T"
        amovs_cab(11).campo = "F4REFERE": amovs_cab(11).valor = Txt_Referencia.Text: amovs_cab(11).Tipo = "T"
        amovs_cab(12).campo = "F4OBSERVA": amovs_cab(12).valor = txtobserva.Text: amovs_cab(12).Tipo = "T"
        amovs_cab(13).campo = "F4CODSOLICITUD": amovs_cab(13).valor = Trim(Txt_NumSolComp.Text): amovs_cab(13).Tipo = "T"
        amovs_cab(14).campo = "F4TIPMON": amovs_cab(14).valor = IIf(Cmbmone.ListIndex = 0, "S", "D"): amovs_cab(14).Tipo = "T"
        amovs_cab(15).campo = "F4IGV": amovs_cab(15).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3igv").SummaryFooterValue, "0.00")): amovs_cab(15).Tipo = "N"
        amovs_cab(16).campo = "F4MONINA": amovs_cab(16).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3monina").SummaryFooterValue, "0.00")): amovs_cab(16).Tipo = "N"
        amovs_cab(17).campo = "F4BASIMP": amovs_cab(17).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3baseimp").SummaryFooterValue, "0.00")): amovs_cab(17).Tipo = "N"
        amovs_cab(18).campo = "F4MONTO": amovs_cab(18).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3total").SummaryFooterValue, "0.00")): amovs_cab(18).Tipo = "N"
        amovs_cab(19).campo = "F4LOCAL": amovs_cab(19).valor = Txt_TOC.Text: amovs_cab(19).Tipo = "T"
        amovs_cab(20).campo = "F4EMPRESA": amovs_cab(20).valor = txtempresa.Text: amovs_cab(20).Tipo = "T"
        amovs_cab(21).campo = "F4NUMCOTIZA": amovs_cab(21).valor = txtCotizacion.Text: amovs_cab(21).Tipo = "T"
        amovs_cab(22).campo = "F4LUGAR_ENTREGA": amovs_cab(22).valor = txtlugar_entrega.Text: amovs_cab(22).Tipo = "T"
        amovs_cab(23).campo = "F4CONTACTO": amovs_cab(23).valor = txtcontacto.Text: amovs_cab(23).Tipo = "T"
        amovs_cab(24).campo = "F4FECENT": amovs_cab(24).valor = Format(abofechaentrega.value, "DD/MM/YYYY"): amovs_cab(24).Tipo = "F"
        amovs_cab(25).campo = "F4RND": amovs_cab(25).valor = 0: amovs_cab(25).Tipo = "N"
        amovs_cab(26).campo = "F4CENTRO": amovs_cab(26).valor = (TxtCodCosto.Text & ""): amovs_cab(26).Tipo = "T"
        amovs_cab(27).campo = "F4TIPDOC": amovs_cab(27).valor = right((CmbTipDoc.Text & ""), 2): amovs_cab(27).Tipo = "T"
        amovs_cab(28).campo = "F4REGULARIZA": amovs_cab(28).valor = IIf(ChK_regularizacion.Checked = False, 0, 1): amovs_cab(28).Tipo = "T"
        amovs_cab(29).campo = "F4DIAPAGO": amovs_cab(29).valor = txtFechaPago.Text: amovs_cab(29).Tipo = "T"
        amovs_cab(30).campo = "F4PAGOPARCIAL": amovs_cab(30).valor = IIf(Chk_pagoparcial.Checked = False, 0, 1): amovs_cab(30).Tipo = "T"
        amovs_cab(31).campo = "F4NOMPROV": amovs_cab(31).valor = pnlnomprv.Caption: amovs_cab(31).Tipo = "T"
        
        Rem SK ADD:
        amovs_cab(32).campo = "F4COLOCADA": amovs_cab(32).valor = IIf(CBool(chkOrdenEnviada.value), -1, 0): amovs_cab(32).Tipo = "N"
        amovs_cab(33).campo = "F4COLOCADAUSER": amovs_cab(33).valor = IIf(CBool(chkOrdenEnviada.value), Trim(txtEnviadoPor.Text), vbNullString): amovs_cab(33).Tipo = "T"
        amovs_cab(34).campo = "F4COLOCADAFECHA": amovs_cab(34).valor = IIf(CBool(chkOrdenEnviada.value), Format(dtpFechaEnvio.value, "Short Date"), vbNullString): amovs_cab(34).Tipo = "F"
        amovs_cab(35).campo = "F4ATENDIDA": amovs_cab(35).valor = IIf(CBool(chkOrdenRecepcionada.value), -1, 0): amovs_cab(35).Tipo = "N"
        amovs_cab(36).campo = "F4ATENDIDAUSER": amovs_cab(36).valor = IIf(CBool(chkOrdenRecepcionada.value), Trim(txtRecepcionadoPor.Text), vbNullString): amovs_cab(36).Tipo = "T"
        amovs_cab(37).campo = "F4ATENDIDAFECHA": amovs_cab(37).valor = IIf(CBool(chkOrdenRecepcionada.value), Format(dtpFechaRecepcion.value, "Short Date"), vbNullString): amovs_cab(37).Tipo = "F"
        
        rsOrdenCab.Close
        
        
        'GRABA_REGISTRO_logistica amovs_cab(), "IF4ORDEN", ctipo, 31, cnn_dbbancos, "F4NUMORD = '" & Txt_NumOC.Text & "' AND F4local = '" & TOC & "'"
        GRABA_REGISTRO_logistica amovs_cab(), "IF4ORDEN", ctipo, 37, cnn_dbbancos, "F4NUMORD = '" & Txt_NumOC.Text & "' AND F4local = '" & TOC & "'"
        
        
    End If
    
    '---------- GRABANDO EL DETALLE DE LA ORDEN DE COMPRA ----------------------'
    If ctipoadm_bd = "M" Then
        If SwRenovar = True Then
            sql = ("delete from if3orden where f4numord= '" & wNumOc & "'  AND F4local = '" & TOC & "'")
            cnn_dbbancos.Execute sql
            ''AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        Else
            sql = ("delete from if3orden where f4numord= '" & Txt_NumOC.Text & "'  AND F4local = '" & TOC & "'")
            cnn_dbbancos.Execute sql
            ''AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        End If
    Else
        If SwRenovar = True Then
            sql = ("delete * from if3orden where f4numord= '" & wNumOc & "'  AND F4local = '" & TOC & "'")
             cnn_dbbancos.Execute sql
            ''AlmacenaQuery_sql sql, cnn_dbbancos
           Actualiza_Log sql, cnn_dbbancos.ConnectionString
        Else
            sql = ("delete * from if3orden where f4numord= '" & Txt_NumOC.Text & "'  AND F4local = '" & TOC & "'")
            cnn_dbbancos.Execute sql
            ''AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        End If
    End If
    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
    'rsOrdenDet.Open "select * from if3orden", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    'If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
    'rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
    soli_acum = ""
    sql = "SELECT * FROM " & cnomtabla & " order by item"
    Set rsdetaoc = Af.OpenSQLForwardOnly(sql, StrCn)
    If rsdetaoc.RecordCount > 0 Then
        With rsdetaoc
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(.Fields("f3codpro")) Then
                
                    codi = .Fields("f3codpro")
                    wcantidad = .Fields("f3canpro")
                    wproducto = Trim$(codi)
                    
                    sql = "select f3presu,f3consumido,f3ocompra from centroproductos where " _
                    & "f3costo='" & wcc & "' and f5codpro='" & wproducto & "'"
    
                    If rstaux.State = adStateOpen Then rst.Close
                    rstaux.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not (rstaux.EOF) Then
                        If jc = 0 Then  'Nuevo
                            ocompra = Val(rstaux.Fields("f3ocompra").value)
                            rstaux.Fields("f3ocompra").value = ocompra + wcantidad
                        Else             'Modifica
                            rstaux.Fields("f3ocompra").value = wcantidad
                        End If
                        rstaux.Update
                    End If
                    rstaux.Close
                    
                    ' sql = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD" _
                        & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL" _
                        & ",F3FENTREGA,F5NOMPRO, F4LOCAL,F3CENCOS,item,f4tipo) VALUES " _
                        & "('" & (Txt_NumOC.Text) & "','" & .Fields("f3codpro") & "','" & .Fields("f5codfab") & "'," _
                        & .Fields("f3canpro") & ",'" & .Fields("f5marca") & "','" & .Fields("f3medida") & "'," _
                        & .Fields("f3canpro") & "," & IIf(IsNull(.Fields("f3preuni")), "0", .Fields("f3preuni")) & "," & IIf(IsNull(.Fields("f3precos")), "0", .Fields("f3precos")) & "," _
                        & IIf(IsNull(.Fields("f3pordct")), "0", .Fields("f3pordct")) & "," & IIf(IsNull(.Fields("f3totdct")), "0", .Fields("f3totdct")) & "," _
                        & IIf(IsNull(.Fields("f5valvta")), "0", .Fields("f5valvta")) & ",'" & IIf(IsNull(.Fields("f5afecto")), " ", .Fields("f5afecto")) & "'," _
                        & IIf(IsNull(.Fields("f3igv")), "0", .Fields("f3igv")) & "," & .Fields("f3total") & ",'" & .Fields("f3fentrega") & "','" _
                        & .Fields("f5nompro") & "','0','" & .Fields("F3CENCOS") & "'," & w & "," & CboTipo.ListIndex & ")"
                    
                    
                    If SwRenovar = True Then
                        sql = "INSERT INTO IF3ORDEN("
                        sql = sql & "F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD, "
                        sql = sql & "F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL, "
                        sql = sql & "F3FENTREGA,item,F5NOMPRO, F4LOCAL,F3CENCOS,F3OBSERVA,cod_solicitud, CODCOLOR, "
                        sql = sql & "F3PORCDEMASIA, F3CANPRO2, F3PRENETO, F5NOMPRO_ING) "
                        sql = sql & "VALUES("
                        sql = sql & "'" & wNumOc & "','" & .Fields("f3codpro") & "','" & .Fields("f3codfab") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & ",'','" & .Fields("f3codmedida") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & "," & .Fields("f3COnigv") & "," & .Fields("f3SInigv") & ","
                        sql = sql & Val(.Fields("f3pordesc") & "") & "," & Val(.Fields("f3valdesc") & "") & ","
                        sql = sql & Val(.Fields("f3baseimp") & "") + Val(.Fields("f3monina") & "") & ", "
                        sql = sql & "'" & IIf((.Fields("f5afecto" & "")) = False, " ", "*") & "',"
                        sql = sql & Val(.Fields("f3igv") & "") & "," & Val(.Fields("f3total") & "") & ",null," & Val(.Fields("item") & "") & ",'"
                        sql = sql & .Fields("f5nompro") & "','1','"
                        
                        If TxtCodCosto.Text = "998" Then
                            sql = sql & .Fields("f3cencos") & "','"
                        Else
                            sql = sql & TxtCodCosto.Text & "','"
                        End If
                        
                        sql = sql & left(.Fields("f3OBSERVA"), 255) & "','" & .Fields("cod_solicitud") & "', '" & Trim(.Fields("CODCOLOR").value & "") & "', "
                        sql = sql & Val(.Fields("F3PORCDEMASIA").value & "") & ", " & Val(.Fields("F3CANPROMAX").value & "") & ", "
                        sql = sql & Val(.Fields("F3NETO").value & "") & ", '" & .Fields("f5nompro_ing") & "')"
                    Else
                        sql = "INSERT INTO IF3ORDEN("
                        sql = sql & "F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD, "
                        sql = sql & "F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL, "
                        sql = sql & "F3FENTREGA,item,F5NOMPRO, F4LOCAL,F3CENCOS,F3OBSERVA,cod_solicitud, CODCOLOR, "
                        sql = sql & "F3PORCDEMASIA, F3CANPRO2, F3PRENETO, F5NOMPRO_ING) "
                        sql = sql & "VALUES("
                        sql = sql & "'" & Txt_NumOC.Text & "', '" & .Fields("f3codpro") & "','" & .Fields("f3codfab") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & ",'','" & .Fields("f3codmedida") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & "," & Val(.Fields("f3COnigv") & "") & "," & Val(.Fields("f3SInigv") & "") & ","
                        sql = sql & Val(.Fields("f3pordesc") & "") & "," & Val(.Fields("f3valdesc") & "") & ","
                        sql = sql & Val(.Fields("f3baseimp") & "") + Val(.Fields("f3monina") & "") & ", "
                        sql = sql & "'" & IIf((.Fields("f5afecto" & "")) = False, " ", "*") & "',"
                        sql = sql & Val(.Fields("f3igv") & "") & "," & Val(.Fields("f3total") & "") & ",null," & Val(.Fields("item") & "") & ",'"
                        sql = sql & .Fields("f5nompro") & "','" & TOC & "','"
                        
                        If TxtCodCosto.Text = "998" Then
                            sql = sql & .Fields("f3cencos") & "','"
                        Else
                            sql = sql & TxtCodCosto.Text & "','"
                        End If
                        
                        sql = sql & left(.Fields("f3OBSERVA"), 255) & "','" & .Fields("cod_solicitud") & "', '" & Trim(.Fields("CODCOLOR").value & "") & "', "
                        sql = sql & Val(.Fields("F3PORCDEMASIA").value & "") & ", " & Val(.Fields("F3CANPROMAX").value & "") & ", "
                        sql = sql & Val(.Fields("F3NETO").value & "") & ", '" & .Fields("f5nompro_ing") & "')"
                    End If
                    
                    cnn_dbbancos.Execute sql
                    ''AlmacenaQuery_sql sql, cnn_dbbancos
                    Actualiza_Log sql, cnn_dbbancos.ConnectionString

                    If Not IsNull(.Fields("cod_solicitud")) And right(soli_acum, 12) <> .Fields("cod_solicitud") Then
                        If Len(Trim(soli_acum)) > 1 Then
                            If InStr(1, soli_acum, .Fields("cod_solicitud")) = 0 Then
                                soli_acum = soli_acum & ", " & .Fields("cod_solicitud")
                            End If
                            
                            sql = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC & "' where cod_solicitud='" & .Fields("cod_solicitud") & "'"
                            cnn_dbbancos.Execute sql
                            ''AlmacenaQuery_sql sql, cnn_dbbancos
                            Actualiza_Log sql, cnn_dbbancos.ConnectionString
                        Else
                            soli_acum = .Fields("cod_solicitud")
                            sql = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC & "' where cod_solicitud='" & .Fields("cod_solicitud") & "'"
                            cnn_dbbancos.Execute sql
                            ''AlmacenaQuery_sql sql, cnn_dbbancos
                            Actualiza_Log sql, cnn_dbbancos.ConnectionString
                        End If
                        
                        ActualizarNumOrd (.Fields("cod_solicitud"))
                        
                        If rst.State = adStateOpen Then rst.Close
                        '************NUEVO CAMBIO ********************************
                        rst.Open "SELECT TB_DETSOLICITUD.COD_SOLICITUD, Sum([tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '')) AS saldo " & _
                                 " FROM TB_DETSOLICITUD LEFT JOIN IF3ORDEN ON (TB_DETSOLICITUD.cod_solicitud = IF3ORDEN.COD_SOLICITUD) AND (TB_DETSOLICITUD.cod_producto = IF3ORDEN.F3CODPRO) Where TB_DETSOLICITUD.COD_SOLICITUD = '" & .Fields("cod_solicitud") & "' " & _
                                 " GROUP BY TB_DETSOLICITUD.COD_SOLICITUD ", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If rst.Fields("saldo") > 0 Then
                            sql = "update tb_cabsolicitud set cs_estado='3' where cod_solicitud='" & .Fields("cod_solicitud") & "'"
                        Else
                            sql = "update tb_cabsolicitud set cs_estado='4' where cod_solicitud='" & .Fields("cod_solicitud") & "'"
                        End If
                        cnn_dbbancos.Execute sql
                        ''AlmacenaQuery_sql sql, cnn_dbbancos
                        Actualiza_Log sql, cnn_dbbancos.ConnectionString
        
                        '**********************************************************
                        rst.Close
                    
                        If rst.State = adStateOpen Then rst.Close
                    End If
                    
'                    Rem SK ADD: Actualizar Historial de Proveedor (EF2PROD_PROV)
'                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "EF2PROD_PROV", "F2CODPRV", Trim(Txt_Prove.Text), "T", "AND F5CODPRO = '" & Trim(.Fields("f3codpro") & "") & "'") = vbNullString Then
'                        CadSql = vbNullString
'                        CadSql = CadSql & "INSERT INTO EF2PROD_PROV(F2CODPRV, F2NOMPRV, F5CODPRO, F5NOMPRO, F5VALVTA, F5CODFAB, "
'                        CadSql = CadSql & "F7CODMED, F2FECHA, F2MONEDA) "
'                        CadSql = CadSql & "VALUES ('" & Trim(Txt_Prove.Text) & "', '" & Trim(pnlnomprv.Caption) & "', "
'                        CadSql = CadSql & "'" & Trim(.Fields("f3codpro") & "") & "', '" & Trim(.Fields("f5nompro") & "") & "', "
'                        CadSql = CadSql & Val(.Fields("f3sinigv") & "") & ", '" & Trim(.Fields("f3codfab") & "") & "', "
'                        CadSql = CadSql & "'" & Trim(.Fields("f3codmedida") & "") & "', CVDATE('" & Format(txt_fecha.value, "Short Date") & "'), "
'                        CadSql = CadSql & "'" & left(Cmbmone.Text, 1) & "')"
'
'                        'cnn_dbbancos.Execute CadSql
'                        ''AlmacenaQuery_sql CadSql, cnn_dbbancos
'                        'Actualiza_Log CadSql, cnn_dbbancos.ConnectionString
'                    Else
'                        CadSql = vbNullString
'                        CadSql = CadSql & "UPDATE EF2PROD_PROV "
'                        CadSql = CadSql & "SET "
'                        CadSql = CadSql & "F2NOMPRV = '" & Trim(pnlnomprv.Caption) & "', "
'                        CadSql = CadSql & "F5NOMPRO = '" & Trim(.Fields("f5nompro") & "") & "', "
'                        CadSql = CadSql & "F5VALVTA = " & Val(.Fields("f3sinigv") & "") & ", "
'                        CadSql = CadSql & "F5CODFAB = '" & Trim(.Fields("f3codfab") & "") & "', "
'                        CadSql = CadSql & "F7CODMED = '" & Trim(.Fields("f3codmedida") & "") & "', "
'                        CadSql = CadSql & "F2FECHA = CVDATE('" & Format(txt_fecha.value, "Short Date") & "'), "
'                        CadSql = CadSql & "F2MONEDA = '" & left(Cmbmone.Text, 1) & "' "
'                        CadSql = CadSql & "WHERE "
'                        CadSql = CadSql & "F2CODPRV = '" & Trim(Txt_Prove.Text) & "' AND "
'                        CadSql = CadSql & "F5CODPRO = '" & Trim(.Fields("f3codpro") & "") & "' AND "
'                        CadSql = CadSql & "CVDATE(F2FECHA) <= CVDATE('" & Format(txt_fecha.value, "Short Date") & "')"
'                    End If
'
'                    cnn_dbbancos.Execute CadSql
'                    ''AlmacenaQuery_sql CadSql, cnn_dbbancos
'                    Actualiza_Log CadSql, cnn_dbbancos.ConnectionString
                    
                    If Val(rsdetaoc.Fields("F3PORDESC") & "") > 0 Then
                        With objAyudaProvDscto
                            .CodigoProveedor = Trim(Txt_Prove.Text)
                            .CodigoProducto = Trim(rsdetaoc.Fields("f3codpro") & "")
                            .DescripcionProducto = Trim(rsdetaoc.Fields("f5nompro") & "")
                            .CodigoUM = Trim(rsdetaoc.Fields("f3codmedida") & "")
                            .Cantidad = Val(rsdetaoc.Fields("F3CANPRO") & "")
                            .Porcentaje = Val(rsdetaoc.Fields("F3PORDESC") & "")
                            .Fecha = Format(txt_fecha.value, "DD/MM/YYYY")
    
                            Call .guardarProvDscto
                        End With
                    End If
                End If
                
                .MoveNext
            Loop
            
        End With
    End If
    
    rsdetaoc.Close
    
    Call VERIFIC_PPRV
    
    If soli_acum <> "" Then
        
        
        sql = "update IF4ORDEN set F4CODSOLICITUD='" & soli_acum & "' where F4NUMORD='" & _
        Txt_NumOC & "'"
        
        cnn_dbbancos.Execute sql
        ''AlmacenaQuery_sql sql, cnn_dbbancos
        Actualiza_Log sql, cnn_dbbancos.ConnectionString
    End If
    
    If Txt_NumSolComp.Text <> "" Then
        
    
        'If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
        'rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", cnn_form, adOpenDynamic, adLockOptimistic
        csql = "SELECT * FROM " & cnomtabla & ""
        Set rsdetaoc = Af.OpenSQLForwardOnly(csql, StrCn)
        If Not rsdetaoc.EOF Then
            With rsdetaoc
                .MoveFirst
                Do While Not .EOF
                    codprod = .Fields("f3codpro") & ""
                    Rem NSE If Val("" & .Fields("f3precos")) > 0 Then
                    If .Fields("check") = True Then
                        Cant = Val("" & .Fields("f3canpro"))
                        ncant_ant = Val("" & .Fields("cant_ant"))
                        
                        sql = "update tb_detsolicitud set candis= candis+" & ncant_ant & "-" & _
                        Cant & " where cod_solicitud='" & _
                        Txt_NumSolComp.Text & "' and cod_producto='" & codprod & "'"
                        cnn_dbbancos.Execute sql
                        ''AlmacenaQuery_sql sql, cnn_dbbancos
                        Actualiza_Log sql, cnn_dbbancos.ConnectionString
                        
                    End If
                    .MoveNext
                Loop
                
''                If RsT.State = adStateOpen Then RsT.Close
''                '************NUEVO CAMBIO ********************************
''                RsT.Open "SELECT TB_DETSOLICITUD.COD_SOLICITUD, Sum([tb_detsolicitud].[ds_cantidad]-Val([if3orden].[f3canpro] & '')) AS saldo " & _
''                         " FROM TB_DETSOLICITUD LEFT JOIN IF3ORDEN ON (TB_DETSOLICITUD.cod_solicitud = IF3ORDEN.COD_SOLICITUD) AND (TB_DETSOLICITUD.cod_producto = IF3ORDEN.F3CODPRO) Where TB_DETSOLICITUD.COD_SOLICITUD = '" & Txt_NumSolComp & "' " & _
''                         " GROUP BY TB_DETSOLICITUD.COD_SOLICITUD ", cnn_dbbancos, adOpenDynamic, adLockOptimistic
''                If RsT.Fields("saldo") > 0 Then
''                    sql = "update tb_cabsolicitud set cs_estado='3' where cod_solicitud='" & Txt_NumSolComp & "'"
''                Else
''                    sql = "update tb_cabsolicitud set cs_estado='4' where cod_solicitud='" & Txt_NumSolComp & "'"
''                End If
''                cnn_dbbancos.Execute sql
''                ''AlmacenaQuery_sql sql, cnn_dbbancos
''                Actualiza_Log sql, cnn_dbbancos.ConnectionString
''
''                '**********************************************************
''                RsT.Close
''
''                If RsT.State = adStateOpen Then RsT.Close
                wgraba = 1
            End With
        End If
        rsdetaoc.Close
    End If
    If SwRenovar = True Then
        MsgBox "Orden de Compra Renovada" & Chr(13) & Txt_NumOC.Text & " --> " & wNumOc, vbInformation, "Sistema de Logistica"
    Else
'        atbmenu.Tools.ITEM("ID_Aprobacion").Visible = True
 '       atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = True
        atbmenu.Tools.ITEM("ID_Grabar").Visible = True
        atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
        atbmenu.Tools.ITEM("ID_Anular").Visible = True
        atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
        
        MsgBox "Orden de Compra Actualizada", vbInformation, "Sistema de Logistica"
    End If
    
    swGrabacion = False
    
    Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Logística"
    'Resume
    Resume Next
End Sub

Private Sub VERIFIC_PPRV()
    Dim CODPROV     As String
    Dim NOMPROV     As String
    Dim NomProd     As String
    Dim rsdetaoc    As New ADODB.Recordset
    Dim sql         As String
    Dim cmoneda     As String
    Dim dfecha      As Date
    Dim ccodfab     As String
    Dim ccodmed     As String
    Dim nprecos     As Double
    
    If rsdetaoc.State = adStateOpen Then rsdetaoc.Close
    'rsdetaoc.Open "SELECT * FROM " & cnomtabla & "", CnTmp, adOpenDynamic, adLockOptimistic
    Set rsdetaoc = Af.OpenSQLForwardOnly("SELECT * FROM " & cnomtabla, StrCn)
    If Not rsdetaoc.EOF Then
        With rsdetaoc
            .MoveFirst
            Do While Not .EOF
                Rem SK ADD: Actualizar Historial de Proveedor (EF2PROD_PROV)
                If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F5CODPRO", "EF2PROD_PROV", "F2CODPRV", Trim(Txt_Prove.Text), "T", "AND F5CODPRO = '" & Trim(.Fields("f3codpro") & "") & "'") = vbNullString Then
                    CadSql = vbNullString
                    CadSql = CadSql & "INSERT INTO EF2PROD_PROV(F2CODPRV, F2NOMPRV, F5CODPRO, F5NOMPRO, F5VALVTA, F5CODFAB, "
                    CadSql = CadSql & "F7CODMED, F2FECHA, F2MONEDA) "
                    CadSql = CadSql & "VALUES ('" & Trim(Txt_Prove.Text) & "', '" & Trim(pnlnomprv.Caption) & "', "
                    CadSql = CadSql & "'" & Trim(.Fields("f3codpro") & "") & "', '" & Trim(.Fields("f5nompro") & "") & "', "
                    CadSql = CadSql & Val(.Fields("f3sinigv") & "") & ", '" & Trim(.Fields("f3codfab") & "") & "', "
                    CadSql = CadSql & "'" & Trim(.Fields("f3codmedida") & "") & "', CVDATE('" & Format(txt_fecha.value, "Short Date") & "'), "
                    CadSql = CadSql & "'" & left(Cmbmone.Text, 1) & "')"
                Else
                    CadSql = vbNullString
                    CadSql = CadSql & "UPDATE EF2PROD_PROV "
                    CadSql = CadSql & "SET "
                    CadSql = CadSql & "F2NOMPRV = '" & Trim(pnlnomprv.Caption) & "', "
                    CadSql = CadSql & "F5NOMPRO = '" & Trim(.Fields("f5nompro") & "") & "', "
                    CadSql = CadSql & "F5VALVTA = " & Val(.Fields("f3sinigv") & "") & ", "
                    CadSql = CadSql & "F5CODFAB = '" & Trim(.Fields("f3codfab") & "") & "', "
                    CadSql = CadSql & "F7CODMED = '" & Trim(.Fields("f3codmedida") & "") & "', "
                    CadSql = CadSql & "F2FECHA = CVDATE('" & Format(txt_fecha.value, "Short Date") & "'), "
                    CadSql = CadSql & "F2MONEDA = '" & left(Cmbmone.Text, 1) & "' "
                    CadSql = CadSql & "WHERE "
                    CadSql = CadSql & "F2CODPRV = '" & Trim(Txt_Prove.Text) & "' AND "
                    CadSql = CadSql & "F5CODPRO = '" & Trim(.Fields("f3codpro") & "") & "' AND "
                    CadSql = CadSql & "CVDATE(F2FECHA) <= CVDATE('" & Format(txt_fecha.value, "Short Date") & "')"
                End If

                cnn_dbbancos.Execute CadSql
                ''AlmacenaQuery_sql CadSql, cnn_dbbancos
                Actualiza_Log CadSql, cnn_dbbancos.ConnectionString


'                CODPROV = Txt_Prove.Text
'                NOMPROV = pnlnomprv.Caption
'                codprod = .Fields("f3codpro") & ""
'                NomProd = .Fields("f5nompro") & ""
'                cmoneda = IIf(Cmbmone.ListIndex = 0, "S", "D")
'                dfecha = Format(txt_fecha.value, "DD/MM/YYYY")
'                'nprecos = Val("" & .Fields("F3PRECOS"))
'                If rsproductos.State = adStateOpen Then rsproductos.Close
'                rsproductos.Open "SELECT F5CODFAB,F7codmed FROM IF5PLA WHERE F5CODPRO='" & codprod & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                If Not rsproductos.EOF Then
'                    ccodfab = left("" & rsproductos.Fields("F5CODFAB"), 15)
'                    ccodmed = "" & rsproductos.Fields("F7codmed")
'                End If
'                rsproductos.Close
'
'                If rst.State = adStateOpen Then rst.Close
'                rst.Open "SELECT * FROM EF2PROD_PROV WHERE F5CODPRO='" & codprod & "' AND " _
'                & "F2CODPRV='" & CODPROV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                If rst.RecordCount = 0 Then
''                    rst.AddNew
''                    rst!F2CODPRV = CodProv
''                    rst!F2NOMPRV = NOMPROV
''                    rst!f5codpro = codprod
''                    rst!f5nompro = NomProd
''                    rst!f5valvta = nprecos
''                    rst.Fields("F2MONEDA") = cmoneda
''                    rst.Fields("F2FECHA") = dfecha
''                    rst!f5codfab = ccodfab
''                    rst!f7codmed = ccodmed
''                    rst.Fields("F2COND_PAGO") = txtcodforma.Text
''                    rst.Fields("F2FORPAG") = txtcodforma.Text
''                    rst.Update
'                    If ctipoadm_bd = "M" Then
'                        sql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F2MONEDA,F2FECHA,F5CODFAB,F7CODMED,F2COND_PAGO,F2FORPAG) VALUES " _
'                            & "('" & CODPROV & "','" & NOMPROV & "','" & codprod & "'," & nprecos & ",'" & cmoneda & "','" & dfecha & "','" _
'                            & ccodfab & "','" & ccodmed & "','" & txtcodforma.Text & "','" & txtcodforma.Text & "')"
'                    Else
'                        sql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F2MONEDA,F2FECHA,F5CODFAB,F7CODMED,F2COND_PAGO,F2FORPAG) VALUES " _
'                            & "('" & CODPROV & "','" & NOMPROV & "','" & codprod & "'," & nprecos & ",'" & cmoneda & "',CVDATE('" & dfecha & "'),'" _
'                            & ccodfab & "','" & ccodmed & "','" & txtcodforma.Text & "','" & txtcodforma.Text & "')"
'                    End If
'                Else
'                    If ctipoadm_bd = "M" Then
'                        sql = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA='" & dfecha & "' WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CODPROV & "'"
'                    Else
'                        sql = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA=CVDATE('" & dfecha & "') WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CODPROV & "'"
'                    End If
'                    cnn_dbbancos.Execute (sql)
'                    'AlmacenaQuery_sql sql, cnn_dbbancos
'                    Actualiza_Log sql, cnn_dbbancos.ConnectionString
'                End If
'                rst.Close
                .MoveNext
            Loop
        End With
    End If
    rsdetaoc.Close
    
End Sub

Private Sub Txt_Prove_DblClick()
    Txt_Prove_KeyDown vbKeyF2, 0
End Sub

Private Sub Txt_Prove_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            With ayuda_proveedores_ocl
                sw_ayuda = True
                sw_ocompra = False
                
                wrucprov = vbNullString
                
                .Show 1
                
                sw_ayuda = False
                
                If Trim(wrucprov) <> vbNullString Then
                    Txt_Prove.Text = wrucprov
                    pnlnomprv.Caption = wnomprov
                    pnldireprv.Caption = wdirprov
                    txtcontacto.Text = wcontacto

                    CmbTipDoc.ListIndex = ModUtilitario.seleccionarItem(CmbTipDoc, wdcto, "DER", 2)

                    txtcodforma.Text = wfpagoprov
                    pnlnomforma.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", wfpagoprov, "T")
                    
                    Select Case Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TIPMON", "EF2PROVEEDORES", "F2NEWRUC", wrucprov, "T"))
                        Case "A"
                            Cmbmone.ListIndex = ModUtilitario.seleccionarItem(Cmbmone, ModUtilitario.sGetINI(App.Path & strNombreFicheroConfigCPgeneral, "ConfigCP", "MonedaPredeterminada", "l"), "IZQ", 1)
                        Case Else
                            Cmbmone.ListIndex = ModUtilitario.seleccionarItem(Cmbmone, ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2TIPMON", "EF2PROVEEDORES", "F2NEWRUC", wrucprov, "T"), "IZQ", 1)
                    End Select
                    
                    ModUtilitario.pulsarTecla vbKeyTab
                End If
            End With
    End Select
    
'    Dim q As Integer
'    If KeyCode = 113 Then
'        sw_ayuda = True
'        sw_ocompra = False
'        'hlp_proveedores.Show 1
'        ayuda_proveedores_ocl.Show 1
'        sw_ayuda = False
'        Txt_Prove.Text = wrucprov
'        pnlnomprv.Caption = wnomprov
'        pnldireprv.Caption = wdirprov
'        txtcontacto.Text = wcontacto
'
'        Rem SK ADD:
'        If wdcto <> vbNullString Then 'Solo seleccionar el Tipo de Documento predeterminado, si este esta configurado para el Proveedor seleccionado
'            For q = 0 To CmbTipDoc.ListCount - 1
'                CmbTipDoc.ListIndex = q
'                If right(CmbTipDoc.Text, 2) = wdcto Then 'wdcto
'                    CmbTipDoc.ListIndex = q
'                    Exit For
'                End If
'            Next
'        End If
'
'        If Len(Trim(wfpagoprov)) > 0 Then
'            txtcodforma.Text = wfpagoprov
'            If rst.State = adStateOpen Then rst.Close
'            rst.Open "SELECT * from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not (rst.EOF) Then
'                pnlnomforma.Caption = Trim("" & rst.Fields("F2DESPAG"))
'            End If
'            rst.Close
'        End If
'        Txt_Prove_KeyPress 13
'    End If
End Sub

Private Sub Txt_Prove_KeyPress(KeyAscii As Integer)
'    On Error Resume Next
'    If KeyAscii = 13 Then
'        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
'    End If
    KeyAscii = ModUtilitario.validarCajaNumerica(KeyAscii)
End Sub

Private Sub Txt_Referencia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

Private Sub Txt_TC_GotFocus()
    
    txt_tc.SelStart = 0
    txt_tc.SelLength = Len(txt_tc.Text)
    
End Sub

Private Sub Txt_TC_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        If txt_tc = "" Then
            MsgBox "Ingrese tipo de cambio", 48, "Sistema de Logística"
            txt_tc.Text = 0#
            txt_tc.SetFocus
            Exit Sub
        End If
        txt_tc = Format(txt_tc, "#0.000")
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

Private Sub Txt_TC_LostFocus()
 
    If Not IsNumeric(txt_tc) Then
        MsgBox "Dato mal ingresado ...Verifique!", vbInformation, "Sistema de Logistica"
        txt_tc.SetFocus
    End If
'
End Sub

'Private Sub txtcodcosto_DblClick()
'
'    txtcodcosto_KeyDown 113, 0
'
'End Sub
'
'Private Sub txtcodcosto_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 113 Then
'        wcodcosto = ""
'        sw_ayuda = True
'        'hlp_centros.Show 1
'        Ayuda_CENTROS.Show 1
'        sw_ayuda = False
'        If Len(Trim(wcodcosto)) > 0 Then
'            txtcodcosto = wcodcosto
'            pnlnomcosto = wdescosto
'            txtcodcosto_KeyPress 13
'        End If
'    End If
'
'End Sub

Private Sub txtcodforma_DblClick()
    
    txtcodforma_KeyDown 113, 0

End Sub

Private Sub txtcodforma_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

'Private Sub txtcodsoli_Change()
'
'    If Not inicio Then swGrabacion = True
'    If Len(Trim(txtcodsoli.Text)) > 0 Then
'        If rst.State = adStateOpen Then rst.Close
'        rst.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(txtcodsoli.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'        If Not rst.EOF Then
'            pnlnomsoli.Caption = "" & rst.Fields("f2nomuser")
'        Else
'            pnlnomsoli.Caption = ""
'            MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
''            txtcodsoli.SetFocus
'        End If
'        rst.Close
'    End If
'End Sub

Private Sub txtcodsoli_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
    
End Sub

'Private Sub imprimir()
'
'    LLENA_TEMPCAB
'    acr_ocompra.Show 1
'
'End Sub

Private Sub eliminar()
Dim gcodigo     As String
Dim gcant       As Double
    
    If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
    rsOrdenCab.Open "SELECT * from if4orden where f4numord='" & Txt_NumOC & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsOrdenCab.EOF Then
        If MsgBox("¿Desea Anular La Orden de Compra?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Logística") = 6 Then
            
            sql = "Update if4ORDEN set f4estado=5,f4estnul='S',f4monto=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            sql = "Update if3ORDEN set f3canpro=0,f3igv=0,f3preuni=0,f5valvta=0,f3precos=0,f3total=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            With dxDBGrid1
                .Dataset.First
                If Not (.Dataset.EOF) Then
                    .Dataset.First
                    If .Dataset.RecordCount > 0 Then
                        Do While Not (.Dataset.EOF)
                            gcodigo = .Dataset.FieldValues("f3codpro")
                            gcant = .Dataset.FieldValues("f3canpro")
                            If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
                            rsOrdenDet.Open "select * from tb_detsolicitud where " _
                            & "cod_solicitud='" & Txt_NumSolComp.Text & "' and cod_producto='" & _
                            gcodigo & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                            
                            If Not (rsOrdenDet.EOF) Then
                                rsOrdenDet.Fields("candis") = rsOrdenDet.Fields("candis") + Val(gcant)
                                rsOrdenDet.Update
                            End If
                            
                            .Dataset.Next
                        Loop
                    End If
'                    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
'                    rsOrdenDet.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'                    If rsOrdenDet(0).Value = 0 Then
'
'                        sql = "update tb_cabsolicitud set cs_estado='1' where cod_solicitud='" & Txt_NumSolComp & "'"
'                        cnn_dbbancos.Execute sql
'                        'AlmacenaQuery_sql sql, cnn_dbbancos
'                    Else
'                        sql = "update tb_cabsolicitud set cs_estado='2' where cod_solicitud='" & Txt_NumSolComp & "'"
'                        cnn_dbbancos.Execute sql
'                        'AlmacenaQuery_sql sql, cnn_dbbancos
'                    End If
'                    rsOrdenDet.Close
                    MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " ha sido Anulada", vbInformation, App.Title
'                    Call Visi
'                    Call Limpia_Orden
'                    sw_nuevo_documento = False
'                    AdicionaItemGrid
'
'                    sw_nuevo_documento = True
'                    Call limpiarCajas
                    txt_fecha.SetFocus
                End If
            End With
        End If
    End If
    rsOrdenCab.Close
    
End Sub

Private Sub eliminar_sin_preguntar()
Dim gcodigo     As String
Dim gcant       As Double
    
    If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
    rsOrdenCab.Open "SELECT * from if4orden where f4numord='" & Txt_NumOC & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsOrdenCab.EOF Then
        'If MsgBox("¿Desea Anular La Orden de Compra?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Logística") = 6 Then
            
            sql = "Update if4ORDEN set f4estado=5,f4estnul='S',f4monto=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            sql = "Update if3ORDEN set f3canpro=0,f3igv=0,f3preuni=0,f5valvta=0,f3precos=0,f3total=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            sql = "update tb_cabsolicitud set cs_estado='2', cs_orden = '', numorden = '' where cod_solicitud='" & Txt_NumSolComp & "'"
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
'            With dxDBGrid1
'                .Dataset.First
'                If Not (.Dataset.EOF) Then
'                    .Dataset.First
'                    If .Dataset.RecordCount > 0 Then
'                        Do While Not (.Dataset.EOF)
'                            gcodigo = .Dataset.FieldValues("f3codpro")
'                            gcant = .Dataset.FieldValues("f3canpro")
'                            If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
'                            rsOrdenDet.Open "select * from tb_detsolicitud where " _
'                            & "cod_solicitud='" & Txt_NumSolComp.Text & "' and cod_producto='" & _
'                            gcodigo & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'
'                            If Not (rsOrdenDet.EOF) Then
'                                rsOrdenDet.Fields("candis") = rsOrdenDet.Fields("candis") + Val(gcant)
'                                rsOrdenDet.Update
'                            End If
'
'                            .Dataset.Next
'                        Loop
'                    End If
''                    If rsOrdenDet.State = adStateOpen Then rsOrdenDet.Close
''                    rsOrdenDet.Open "select sum(candis) as cant from tb_detsolicitud where cod_solicitud='" & Txt_NumSolComp & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
''                    If rsOrdenDet(0).Value = 0 Then
''
''                        sql = "update tb_cabsolicitud set cs_estado='1' where cod_solicitud='" & Txt_NumSolComp & "'"
''                        cnn_dbbancos.Execute sql
''                        'AlmacenaQuery_sql sql, cnn_dbbancos
''                    Else
''                        sql = "update tb_cabsolicitud set cs_estado='2' where cod_solicitud='" & Txt_NumSolComp & "'"
''                        cnn_dbbancos.Execute sql
''                        'AlmacenaQuery_sql sql, cnn_dbbancos
''                    End If
''                    rsOrdenDet.Close
'                   ' MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " ha sido Anulada", vbInformation, App.Title
''                    Call Visi
''                    Call Limpia_Orden
''                    sw_nuevo_documento = False
''                    AdicionaItemGrid
''
''                    sw_nuevo_documento = True
''                    Call limpiarCajas
'                    txt_fecha.SetFocus
'                End If
'            End With
        'End If
    End If
    rsOrdenCab.Close
    filtrox = 1
    'Txt_NumOC_KeyPress 13
End Sub
Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.FocusedIndex = 1
        End If
        If Action = daPost Then
            calcula
        End If
    End If
    
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim rsproduc    As New ADODB.Recordset
    If Not inicio Then swGrabacion = True
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3codpro", "f5codfab":
            If pnlnomprv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
            
            If rsproduc.State = adStateOpen Then rsproduc.Close
            sql = "SELECT B.F5CODPRO,B.F5TEXTO,B.F5NOMPRO,B.F5AFECTO,B.F5CODFAB,B.F5VALVTA,B.F7CODMED,B.F5MARCA FROM EF2PROD_PROV AS A,IF5PLA AS B WHERE A.F2CODPRV='" & wrucprov & "' AND B.F5CODFAB='" & dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").value & "' ORDER BY B.F5CODPRO"
            rsproduc.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
            If Not rsproduc.EOF Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").value = rsproduc.Fields("F5CODPRO") & ""
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").value = rsproduc.Fields("F5CODFAB") & ""
                If Len(Trim(rsproduc.Fields("F5TEXTO")) & "") > 0 Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5nompro").value = rsproduc.Fields("F5TEXTO") & ""
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5nompro").value = rsproduc.Fields("F5NOMPRO") & ""
                End If
                dxDBGrid1.Columns.ColumnByFieldName("ds_unidmed").value = rsproduc.Fields("F7CODMED") & ""
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & rsproduc.Fields("F5MARCA") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").value = rsproduc.Fields("F5MARCA") & ""
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").value = rsproduc.Fields("F5AFECTO") & ""
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Val(rsproduc.Fields("F5VALVTA") & "")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format$(abofechaentrega.value, "DD/MM/YYYY"))
                dxDBGrid1.Dataset.FieldValues("check") = True
                dxDBGrid1.Dataset.Post
            End If
            rsproduc.Close
            Set rsproduc = Nothing
            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("check").ColIndex - 1
        Case "f3preuni"
                Dim Cantidad As Double
                Dim totdcto As Double
                Dim ValVta As Double
                Dim IGV  As Double
                With dxDBGrid1
                    Cantidad = Val(Format(dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO").value, "0.00"))
                    If Cantidad > 0 Then
                        .Dataset.Edit
                                totdcto = 0
                                
                                ValVta = Val(Format(Cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3PREuni").value, "0.0000")) - totdcto, "0.00"))
                                
                                .Columns.ColumnByFieldName("F5VALVTA").value = Format$(ValVta, "###,##0.00")
                                IGV = 0
                                .Columns.ColumnByFieldName("F3precos").value = Val(Format("" & .Columns.ColumnByFieldName("F3PREuni").value, "0.0000"))
                                If .Columns.ColumnByFieldName("F5AFECTO").value = "*" Then     'Afecto
                                    
                                    .Columns.ColumnByFieldName("F3precos").value = Val(Format("" & .Columns.ColumnByFieldName("F3PREuni").value, "0.0000")) / (1 + (wwigv / 100))
                                    IGV = (.Columns.ColumnByFieldName("F3precos").value * (wwigv / 100)) * Cantidad
                                End If
                                .Columns.ColumnByFieldName("F3IGV").value = Format$(IGV, "#,##0.00")
                                .Columns.ColumnByFieldName("F3TOTAL").value = Format$(ValVta, "###,##0.00")
                            .Dataset.Post
                        End If
                    End With
        Case "f3canpro", "f3precos", "f3pordct", "f5afecto":
            dxDBGrid1.Dataset.Edit
            Calcula_PvtaTot
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = True
            sw_nuevo_item = False
            calcula
        End Select
        If dxDBGrid1.Columns.FocusedColumn.ObjectName = "check" Then
            dxDBGrid1.Dataset.Edit
            'dxDBGrid1.Columns.ColumnByFieldName("check").Value = True
            Calcula_PvtaTot
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = True
            sw_nuevo_item = False
            calcula
        End If
End Sub
Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
     If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").value & "")) = 0 Then
                    Allow = False
                Else
                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").ColIndex
                End If
            End If
        End If
        If Action = daDelete Then
            sw_detalle = True
            dxDBGrid1.Dataset.Refresh
        End If
    End If
   
'    If sw_nuevo_item = False Then
'        If Action = daInsert Then
'            If dxDBGrid1.Dataset.RecordCount > 0 Then
'                If Len(Trim(dxDBGrid1.Columns(1).Value & "")) = 0 Then
'                    Allow = False
'                End If
'            End If
'        End If
'        If Action = daDelete Then
'            dxDBGrid1.Dataset.Delete
'        End If
'    End If

End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
Dim i               As Integer
    
    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_documento = False Then
        DELETEREC_LOG cnomtabla, cnn_form
        DELETEREC_LOG cnomtabla, CnTmp
        dxDBGrid1.Dataset.Refresh
    End If
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.ADODataset.CommandText = "select * from TMPORDENDECOMPRA"

    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    With dxDBGrid1.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
        For i = 1 To 1
            If sw_nuevo_temp = False Then
                If sw_nuevo_documento = True Then
                    .Edit
                Else
                    .Append
                End If
                sw_nuevo_temp = True
            Else
                .Append
            End If
            .FieldValues("item") = i
            .FieldValues("f3codpro") = ""
            .FieldValues("f5nompro") = ""
            .FieldValues("f3medida") = ""
            .FieldValues("f5marca") = ""
            .FieldValues("f3canpro") = Null
            .FieldValues("f3precos") = Null
            .FieldValues("f3pordct") = Null
            .FieldValues("f3totdct") = Null
            .FieldValues("f5valvta") = Null
            .FieldValues("f5afecto") = ""
            .FieldValues("f3igv") = Format(0, "###,##0.00")
            .FieldValues("f3preuni") = Format(0, "###,##0.00")
            .FieldValues("f3total") = Format(0, "###,##0.00")
            .FieldValues("f5codfab") = ""
            .FieldValues("f3fentrega") = Format$(abofechaentrega.value, "dd/mm/yyyy")
            .FieldValues("check") = False
            .FieldValues("cant_ant") = 0#
    
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

End Sub

Private Sub AdicionaItemGrid()
Dim sw_nuevo_temp   As Boolean
Dim i               As Integer
    
    Grid.Dataset.Active = False
    Grid.Dataset.Close
    Grid.Dataset.ADODataset.ConnectionString = StrCn
    Grid.Dataset.ADODataset.CommandText = "select * from TMPORDENDECOMPRA"
    Grid.Dataset.Active = True
    Grid.Dataset.Open
    
    If sw_nuevo_documento = False Then
        
        If Grid.Dataset.RecordCount > 0 Then
            'DELETEREC_LOG cnomtabla, CnTmp
            For i = Grid.Dataset.RecordCount To 1 Step -1
                Grid.Dataset.RecNo = i
                Grid.Dataset.Delete
            Next
        End If
        'Grid.Dataset.Edit
        'Grid.Dataset.Post
        Grid.Dataset.Refresh
    End If
    
    With Grid.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
        For i = 1 To 1
            If sw_nuevo_temp = False Then
                If sw_nuevo_documento = True Then
                    .Edit
                Else
                    .Append
                End If
                sw_nuevo_temp = True
            Else
                .Append
            End If
            .FieldValues("item") = i
            .FieldValues("f3codpro") = ""
            .FieldValues("f3codfab") = ""
            .FieldValues("f5nompro") = ""
            .FieldValues("f3codmedida") = ""
            .FieldValues("f3canpro") = Null
            .FieldValues("f3sinigv") = Null
            .FieldValues("f3conigv") = Null
            .FieldValues("f3monina") = Null
            .FieldValues("f3baseimp") = Null
            .FieldValues("f5afecto") = False
            .FieldValues("f3igv") = Format(0, "###,##0.00")
            '.FieldValues("f3preuni") = Format(0, "###,##0.00")
            .FieldValues("f3total") = Format(0, "###,##0.00")
            
            '.FieldValues("f3fentrega") = Format$(abofechaentrega.Value, "dd/mm/yyyy")
            '.FieldValues("check") = False
            '.FieldValues("cant_ant") = 0#
    
        Next
        .Post
        
        '
        sw_nuevo_item = False
    End With
    Grid.Dataset.Close
    Grid.Dataset.Open
    Grid.Dataset.Refresh
End Sub


Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim sql         As String
    
    Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3codpro", "f5codfab":
        
            If pnlnomprv.Caption = "" Then
                MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
                Txt_Prove.SetFocus
                Exit Sub
            End If
        
            wcodproducto = ""
            wrucprov = Trim(Txt_Prove.Text)
            wnomprov = Trim(pnlnomprv.Caption)
            
            Con_Ayu = 3
            
            dxDBGrid1.Dataset.Close
            
            ayuda_productos.Show 1
            
            dxDBGrid1.Dataset.Open
            
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").value = wcodfab
                
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").value = wmarca
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").value = wafecto

                Rem EMB dxDBGrid1.Dataset.FieldValues("f5valvta") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.value, "dd/mm/yyyy"))
                dxDBGrid1.Dataset.FieldValues("f3canpro") = Format(0, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("check") = True
                dxDBGrid1.Columns.FocusedIndex = 6
            End If
  End Select
 
 Select Case dxDBGrid1.Columns.FocusedColumn.FieldName

        Case "f3cencos":
             ' If dxDBGrid1.Columns.FocusedColumn.ColumnByFieldName = "f3cencos" Then
            'MsgBox "CenCos"
                           
            wcodcosto = "": wdescosto = "": wunicosto = "":
            
            Ayuda_Centros.Show 1
                      
            If Len(Trim(wcodcosto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3cencos").value = wcodcosto
                dxDBGrid1.Dataset.Post
               
            End If
            'End If
            '****************
            
            dxDBGrid1.Columns.FocusedIndex = 4
    End Select

 
 
    If dxDBGrid1.Columns.FocusedColumn.ObjectName = "COLUMNELIMINAR" Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Count = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItemGrid
                sw_detalle = False
                'atbmenu.Tools("IDGrabar").Enabled = False
            Else
                dxDBGrid1.Dataset.Delete
            End If
            'CalculaTotal
            mostrarTotales
            
            sw_nuevo_item = False
        End If
    End If


End Sub

Private Sub txtplazo_entrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub

Private Sub TxtRnd_GotFocus()
   TxtRnd.SelStart = 0
   TxtRnd.SelLength = Len(TxtRnd.Text)

End Sub

Private Sub TxtRnd_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57, 8: KeyAscii = KeyAscii
Case 45
    CtaPuntos = 0
    For i = 1 To Len(TxtRnd.Text)
        If Mid(TxtRnd.Text, i, 1) = "-" Then
            CtaPuntos = CtaPuntos + 1
        End If
    Next
    If CtaPuntos >= 1 Then
        KeyAscii = 0
    End If
Case 46
    CtaPuntos = 0
    For i = 1 To Len(TxtRnd.Text)
        If Mid(TxtRnd.Text, i, 1) = "." Then
            CtaPuntos = CtaPuntos + 1
        End If
    Next
    If CtaPuntos >= 1 Then
        KeyAscii = 0
    End If
Case 13
    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub TxtRnd_LostFocus()
calcula
End Sub

Private Sub txtusuario_Change()
    
    If Not inicio Then swGrabacion = True

End Sub

Private Sub MODIFICAR_OC()
    flagwin = True
    Wnuevo = False
'    Txt_TOC.Text = strTipoOrden
'    Txt_NumOC.Text = strNumeroOrden
'
'    With rsOrdenCab
'        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
'
'        sql = "SELECT * from if4orden where f4numord='" & GOC & "' AND f4LOCAL = '" & TOC & "'"
'
'        rsOrdenCab.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'
'        If Not (.EOF) Then
''            MostrarDatosOC
'            MostrarDatosOC_Grid
'            ExisteOrdenCompra = True
'        Else
'            ExisteOrdenCompra = False
'        End If
'
'        .Close
'    End With

    limpiarCajas
    
'    If sw_nuevo_documento = True Then
'        sw_nuevo_documento = False
'
'        AdicionaItemGrid
'
'        sw_nuevo_documento = True
'    Else
'        inicio = True
'
'        MODIFICAR_OC
'
'        sw_nuevo_documento = False
'        inicio = False
'    End If
    
    With objAyudaOrden
        .inicializarEntidades
        
        .TipoOrden = strTipoOrden
        .NumeroOrden = strNumeroOrden
                
        If .obtenerOrden Then
            inicio = True
        
            MostrarDatosOC_Grid
            ExisteOrdenCompra = True
            
            sw_nuevo_documento = False
            inicio = False
        Else
            sw_nuevo_documento = False

            AdicionaItemGrid
            ExisteOrdenCompra = False
            
            sw_nuevo_documento = True
        End If
    End With
End Sub

'Private Sub txtuupp_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 113 Then
'        wcodlocalidad = "": wdeslocalidad = ""
''        hlp_uupp.Show 1
'        If Len(Trim(wcodlocalidad)) > 0 Then
'            txtuupp.Text = Trim(wcodlocalidad)
'            txtdesuupp.Caption = Trim(wdeslocalidad)
'            txtuupp_KeyPress 13
'        End If
'    End If
'
'End Sub

Private Sub txtuupp_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If

End Sub


Private Sub GRABA_GRID(prucprov As String)
Dim rsprodprov      As New ADODB.Recordset
Dim rstempdet       As New ADODB.Recordset
Dim csql            As String
Dim nitem           As Integer
Dim ccodprod        As String
Dim nprecos         As Double
Dim totdcto         As Double
Dim ValVta          As Double
Dim IGV             As Double
Dim preciounit      As Double
Dim TOTAL           As Double
    
    If rstempdet.State = adStateOpen Then rstempdet.Close
    rstempdet.Open "SELECT * FROM " & cnomtabla & " ORDER BY F3CODPRO", CnTmp, adOpenDynamic, adLockBatchOptimistic
    If Not rstempdet.EOF Then
        rstempdet.MoveFirst
        Do While Not rstempdet.EOF
            nitem = Val(rstempdet.Fields("ITEM") & "")
            ccodprod = Trim(rstempdet.Fields("F3CODPRO") & "")
            If rsprodprov.State = adStateOpen Then rsprodprov.Close
            rsprodprov.Open "SELECT F5VALVTA FROM EF2PROD_PROV WHERE F5CODPRO='" & ccodprod & "' AND F2CODPRV='" & prucprov & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rsprodprov.EOF Then
                nprecos = Val(rsprodprov.Fields("F5VALVTA") & "")
                If Trim(rstempdet.Fields("F5AFECTO") & "") = "*" Then
                    totdcto = nprecos * Val(rstempdet.Fields("F3PORDCT") & "") / 100
                    ValVta = Val(rstempdet.Fields("F3CANPRO") & "") * nprecos - totdcto
                    IGV = ValVta * (wwigv / 100)
                    preciounit = nprecos + (nprecos * (wwigv / 100))
                    TOTAL = ValVta + IGV
                Else
                    totdcto = nprecos * Val(rstempdet.Fields("F3PORDCT") & "") / 100
                    ValVta = Val(rstempdet.Fields("F3CANPRO") & "") * nprecos - totdcto
                    IGV = 0#
                    preciounit = nprecos
                    TOTAL = ValVta + IGV
                End If
                csql = "UPDATE " & cnomtabla & " SET F3PRECOS=" & nprecos & ",F3TOTDCT=" & totdcto & _
                       ",F5VALVTA=" & ValVta & ",F3IGV=" & IGV & ",F3PREUNI=" & preciounit & ",F3TOTAL=" & TOTAL & _
                       " WHERE ITEM=" & nitem & " AND F3CODPRO='" & ccodprod & "'"
                'cnn_form.Execute (csql)
                cnDBTemp.Execute csql
                
                'AlmacenaQuery_sql sql, cnDBTemp
                End If
            rsprodprov.Close
            rstempdet.MoveNext
        Loop
    End If
    rstempdet.Close
    dxDBGrid1.Dataset.Refresh

End Sub

Private Sub TRASLADA_CTASXPAGAR(pnumero As String)
Dim ncorre_d            As Double
Dim amovs_cab(0 To 18)  As a_grabacion
Dim rsif4orden            As New ADODB.Recordset
Dim rsbf5pla            As New ADODB.Recordset
Dim RsProveedor         As New ADODB.Recordset
Dim ntotal              As Double
Dim ntc                 As Double
Dim cdetal              As String
Dim dfechamov           As Date
Dim ccodprov            As String
Dim cnomprov            As String
Dim cruc                As String
Dim cnro_comp           As String
Dim Moneda              As String
Dim csql                As String
Dim RSPAG_DCTO          As New ADODB.Recordset

    If cnn_ctaspag.State = adStateOpen Then cnn_ctaspag.Close
    cconex_ctaspag = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos.MDB" & ";Persist Security Info=False"
    cnn_ctaspag.Open cconex_ctaspag
    
    ntotal = 0#: ntc = 0#: cdetal = ""
    If rsif4orden.State = adStateOpen Then rsif4orden.Close
    rsif4orden.Open "SELECT F4MONTO,F4FECEMI,F4TIPCAM,F4OBSERVA FROM IF4ORDEN WHERE F4NUMORD=" & pnumero & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsif4orden.EOF Then
        ntotal = Val("" & rsif4orden.Fields("F4MONTO"))
        ntc = Val("" & rsif4orden.Fields("F4TIPCAM"))
        cdetal = left(Trim("" & rsif4orden.Fields("F4OBSERVA")), 100)
        dfechamov = Format(rsif4orden.Fields("F4FECEMI"), "DD/MM/YYYY")
    End If
    rsif4orden.Close
    
    cruc = Txt_Prove.Text
    If RsProveedor.State = adStateOpen Then RsProveedor.Close
    csql = "SELECT F2NOMPROV,F2CODPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & cruc & "'"
    RsProveedor.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProveedor.EOF Then
        cnomprov = RsProveedor.Fields("F2NOMPROV") & ""
        ccodprov = RsProveedor.Fields("F2CODPROV") & ""
    End If
    RsProveedor.Close
    
    If RSPAG_DCTO.State = adStateOpen Then RSPAG_DCTO.Close
    RSPAG_DCTO.Open "SELECT CORRELA FROM PAG_DCTO ORDER BY CORRELA DESC", cnn_ctaspag, adOpenDynamic, adLockOptimistic
    If Not RSPAG_DCTO.EOF Then
        ncorre_d = RSPAG_DCTO.Fields("CORRELA") + 1
    Else
        ncorre_d = 1
    End If
    RSPAG_DCTO.Close
    
    cnro_comp = "O/c" & Format(pnumero, "0000000")
    Moneda = IIf(Cmbmone.ListIndex = 0, "S", "D")
    
    amovs_cab(0).campo = "VIA_INGR": amovs_cab(0).valor = "1": amovs_cab(0).Tipo = "T"
    amovs_cab(1).campo = "CORRELA": amovs_cab(1).valor = ncorre_d: amovs_cab(1).Tipo = "N"
    amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = cnro_comp: amovs_cab(2).Tipo = "T"
    amovs_cab(3).campo = "FCH_COMP": amovs_cab(3).valor = dfechamov: amovs_cab(3).Tipo = "F"
    amovs_cab(4).campo = "PROVEEDORO": amovs_cab(4).valor = ccodprov: amovs_cab(4).Tipo = "T"
    amovs_cab(5).campo = "RUC": amovs_cab(5).valor = cruc: amovs_cab(5).Tipo = "T"
    amovs_cab(6).campo = "MONEDAO": amovs_cab(6).valor = Moneda: amovs_cab(6).Tipo = "T"
    amovs_cab(7).campo = "TOTALO": amovs_cab(7).valor = ntotal: amovs_cab(7).Tipo = "N"
    amovs_cab(8).campo = "TCAMBIOO": amovs_cab(8).valor = ntc: amovs_cab(8).Tipo = "N"
    amovs_cab(9).campo = "PROVEEDOR": amovs_cab(9).valor = ccodprov: amovs_cab(9).Tipo = "T"
    amovs_cab(10).campo = "MONEDA": amovs_cab(10).valor = Moneda: amovs_cab(10).Tipo = "T"
    amovs_cab(11).campo = "TCAMBIO": amovs_cab(11).valor = ntc: amovs_cab(11).Tipo = "N"
    amovs_cab(12).campo = "TOTAL": amovs_cab(12).valor = ntotal: amovs_cab(12).Tipo = "N"
    amovs_cab(13).campo = "SALDO": amovs_cab(13).valor = ntotal: amovs_cab(13).Tipo = "N"
    amovs_cab(14).campo = "DEB_HAB": amovs_cab(14).valor = "H": amovs_cab(14).Tipo = "T"
    amovs_cab(15).campo = "REFERENCIA": amovs_cab(15).valor = cdetal: amovs_cab(15).Tipo = "T"
    amovs_cab(16).campo = "NOMPROV": amovs_cab(16).valor = cnomprov: amovs_cab(16).Tipo = "T"
    amovs_cab(17).campo = "CONCEPTO": amovs_cab(17).valor = cdetal: amovs_cab(17).Tipo = "T"
    amovs_cab(18).campo = "FCH_VCTO": amovs_cab(18).valor = dfechamov: amovs_cab(18).Tipo = "F"
    
    GRABA_REGISTRO_logistica amovs_cab(), "PAG_DCTO", "A", 18, cnn_ctaspag, ""
    
    cnn_ctaspag.Close
    sql = ("UPDATE IF4ORDEN SET F4CORRELA=" & ncorre_d & " WHERE F4NUMORD=" & pnumero & "")
    cnn_dbbancos.Execute sql
    'AlmacenaQuery_sql sql, cnn_dbbancos
    Actualiza_Log sql, cnn_dbbancos.ConnectionString
End Sub

Private Sub EnviaMail(pNumOC As String)
    Dim Cab As New ADODB.Recordset
    Dim Det As New ADODB.Recordset
    Dim RsCC As New ADODB.Recordset
    Dim nError As Long
    Dim csql As String
    Dim d As Integer
    Dim wSep As String
    Dim StrDetLine As String
    
    csql = "select * from ef2proveedores"
    
    Set RsCC = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    '0001.-informa pago a proveedor
    csql = "SELECT EF2USERS.MAIL "
    csql = csql & "FROM EF2USERS INNER JOIN (EF2TAREAS INNER JOIN EF2TAREAUSERS ON EF2TAREAS.F2CODTAREA = EF2TAREAUSERS.F2CODTAREA) "
    csql = csql & "ON EF2USERS.F2CODUSER = EF2TAREAUSERS.F2CODUSER "
    csql = csql & "WHERE (((EF2TAREAS.F2CODTAREA)='0005'))"
    
    Set Rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    editTo = ""
    
    If Rs.RecordCount = 0 Then Exit Sub
    'llena variables
    InternetMail.Organization = wnomcia
    InternetMail.UserName = wMailReplyCia: editFrom = wMailReplyCia
    InternetMail.Password = wPassReplyCia
    InternetMail.RelayServer = "smtp.gmail.com"
    InternetMail.RelayPort = 465
    InternetMail.Secure = True
    InternetMail.Options = 5
    
    
    d = 0
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            If d = 0 Then
                editTo = Rs!Mail
            Else
                editTo = editTo & ";" & Rs!Mail
            End If
            
            Rs.MoveNext
            d = d + 1
        Loop
    End If
    'carga cabecera
    csql = "SELECT * from IF4ORDEN WHERE f4local='" & TOC & "' AND F4NUMORD='" & pNumOC & "'"
    Set Cab = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    'carga detalle
    'csql = "SELECT * FROM IF3ORDEN WHERE f4local='"& TOC & "' AND F4NUMORD='" & pNumOC & "' order by Item"
    
    csql = "SELECT IF3ORDEN.ITEM, IF3ORDEN.F3CANPRO, IF3ORDEN.F5NOMPRO, EF7MEDIDAS.F7SIGMED, IF3ORDEN.F3TOTAL "
    csql = csql & "FROM IF3ORDEN LEFT JOIN EF7MEDIDAS ON IF3ORDEN.UNIDAD = EF7MEDIDAS.F7CODMED "
    csql = csql & "WHERE (((IF3ORDEN.F4LOCAL)='1') AND ((IF3ORDEN.F4NUMORD)='" & pNumOC & "')) "
    csql = csql & "ORDER BY IF3ORDEN.ITEM"
    
    Set Det = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    If Cab.RecordCount > 0 And Det.RecordCount > 0 Then
        Cab.MoveFirst
        Det.MoveFirst
        editSubject = "SDA-" & Cab.Fields("F4NUMORD")
        RsCC.Filter = adFilterNone
        RsCC.Filter = "f2newruc='" & Cab!F4CODPRV & "'"
        If RsCC.RecordCount > 0 Then
            editMessageText = "Proveedor: " & RsCC!F2NOMPROV & vbCrLf
        End If
        editMessageText = editMessageText & "Forma de Pago: " & ObtenerCampo("ef2forpag", "f2despag", "f2forpag", Cab!F4FORPAG, "T", cnn_dbbancos) & vbCrLf
        editMessageText = editMessageText & "Solicitante: " & Cab!F4CODSOL & vbCrLf
        editMessageText = editMessageText & "Importe: " & Format(Cab!F4MONTO, "###,###,###,##0.00") & vbCrLf
        editMessageText = editMessageText & "Moneda: " & IIf(Cab!F4TIPMON = "S", "Nuevos Soles", "Dólares Americanos") & vbCrLf
        editMessageText = editMessageText & vbCrLf
        editMessageText = editMessageText & "Detalle" & vbCrLf
        editMessageText = editMessageText & "-------"
        editMessageText = editMessageText & vbCrLf
        d = 0
        Do While Not Det.EOF
            wSep = "................................................."
            StrDetLine = Format(Det!ITEM & "", "000") & ".- (" & Val(Det!F3CANPRO & "") & " " & left(Det!F7SIGMED & Space(3), 3) & ") "
            StrDetLine = StrDetLine & left(Det!F5NOMPRO & wSep, 50) & right(wSep & Format(Val(Det!F3TOTAL & ""), "###,###,##0.00"), 12)
            If d = 0 Then
                editMessageText = editMessageText & StrDetLine
            Else
                editMessageText = editMessageText & vbCrLf
                editMessageText = editMessageText & StrDetLine
            End If
            Det.MoveNext
            d = d + 1
        Loop
        nError = CreateMessage()
        If nError Then
            
            MsgBox "Unable to create a new message" & vbCrLf & _
                   InternetMail.LastErrorString, vbExclamation
            Exit Sub
        End If
        
        If InternetMail.Recipients = 0 Then
            MsgBox "There are no recipients for this message", _
                   vbInformation
            Exit Sub
        End If
        
        ' Begin the process of delivering the message

        nError = InternetMail.SendMessage()
        
        If nError Then
            
            MsgBox "Unable to send message" & vbCrLf & _
                   InternetMail.LastErrorString, vbExclamation
                   
            Exit Sub
        Else
            
            csql = "UPDATE IF4ORDEN SET F4ESTNUL='P' WHERE f4local='" & TOC & "' AND F4NUMORD='" & Txt_NumOC.Text & "'"
            cnn_dbbancos.Execute csql
            MsgBox "Su solicitud de aprobación (SDA-" & Txt_NumOC.Text & "), fue enviada.", vbInformation, wnomcia
            atbmenu.Tools.ITEM("ID_Grabar").Visible = False
            atbmenu.Tools.ITEM("ID_Anular").Visible = False
            atbmenu.Tools.ITEM("ID_Eliminar").Visible = False
            'atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = False
            atbmenu.Tools.ITEM("ID_Aprobacion").Visible = False
        End If
    End If
End Sub

Private Function CreateMessage() As Long
    Dim strFontName As String
    Dim strFontSize As String
    Dim strMessageHTML As String
    Dim nCharacterSet As Long
    Dim nEncodingType As Long
    Dim nindex As Long
    Dim nError As Long
        
    CreateMessage = 0
    
    ' If the user has entered any HTML text, then make sure that it is
    ' properly formed HTML and specify a font and font size to use;
    ' this is hard-coded, but an application would obviously want to
    ' make something like font selection customizable
    strMessageHTML = ""
    
    ' Determine what character set was selected by the user
    '
    nCharacterSet = 2
    
    nEncodingType = 1
    
    ' Use the ComposeMessage method to do all of the hard work of
    ' creating the actual message
    '
    nError = InternetMail.ComposeMessage(editFrom, _
                                          editTo, _
                                          editCc, _
                                          editBcc, _
                                          editSubject, _
                                          editMessageText, _
                                          strMessageHTML, _
                                          nCharacterSet, _
                                          nEncodingType)
    
    If nError Then
        CreateMessage = nError
        Exit Function
    End If
    
    '
    ' Attach each file that was selected by the user
    '
    
    
    '
    ' Set the Priority property to the message priority that
    ' was selected by the user
    '
    
    
    

    '
    ' Set the Organization property to the name of the current user's
    ' organization or company; this is specified in the options dialog
    '
    
    
End Function


Private Sub VerificaAtencionDeRequerimiento(ByVal Numero_de_Requerimiento As String)
Dim RsV As New ADODB.Recordset
csql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.ITEM, TB_CABSOLICITUD.cs_fecha, TB_CABSOLICITUD.cs_codsolicitante, "
csql = csql & "EF2USERS.F2NOMUSER, TB_CABSOLICITUD.cs_observaciones,TB_CABSOLICITUD.vbjefecc, TB_DETSOLICITUD.F5CODCOSTO, CENTROS.F3DESCRIP, "
csql = csql & "TB_DETSOLICITUD.cod_producto, TB_DETSOLICITUD.ds_descripcion, TB_DETSOLICITUD.ds_unidmed, EF7MEDIDAS.F7SIGMED, "
csql = csql & "TB_DETSOLICITUD.ds_cantidad, IIf(IsNull(ORDENES.CANT_ORDEN),0,ORDENES.CANT_ORDEN) AS TOT_ORDEN, "
csql = csql & "TB_DETSOLICITUD.ds_cantidad-IIf(IsNull(ORDENES.CANT_ORDEN),0,ORDENES.CANT_ORDEN) AS SALDO "
csql = csql & "FROM (TB_CABSOLICITUD LEFT JOIN EF2USERS ON TB_CABSOLICITUD.cs_codsolicitante = EF2USERS.F2CODUSER) "
csql = csql & "INNER JOIN (((TB_DETSOLICITUD LEFT JOIN ["
csql = csql & "SELECT IF3ORDEN.COD_SOLICITUD, IF3ORDEN.F3CENCOS, IF3ORDEN.F3CODPRO, Sum(IF3ORDEN.F3CANPRO) AS CANT_ORDEN "
csql = csql & "From IF3ORDEN "
csql = csql & "GROUP BY IF3ORDEN.COD_SOLICITUD, IF3ORDEN.F3CENCOS, IF3ORDEN.F3CODPRO "
csql = csql & "ORDER BY IF3ORDEN.COD_SOLICITUD, IF3ORDEN.F3CENCOS, IF3ORDEN.F3CODPRO"
csql = csql & "]. AS ORDENES "
csql = csql & "ON (TB_DETSOLICITUD.cod_producto = ORDENES.F3CODPRO) AND (TB_DETSOLICITUD.F5CODCOSTO = ORDENES.F3CENCOS) "
csql = csql & "AND (TB_DETSOLICITUD.cod_solicitud = ORDENES.COD_SOLICITUD)) "
csql = csql & "LEFT JOIN CENTROS ON TB_DETSOLICITUD.F5CODCOSTO = CENTROS.F3COSTO) LEFT JOIN EF7MEDIDAS ON "
csql = csql & "TB_DETSOLICITUD.ds_unidmed = EF7MEDIDAS.F7CODMED) ON TB_CABSOLICITUD.cod_solicitud = TB_DETSOLICITUD.cod_solicitud "
csql = csql & "Where ((([TB_DETSOLICITUD].[ds_cantidad] - IIf(IsNull([ORDENES].[CANT_ORDEN]), 0, [ORDENES].[CANT_ORDEN])) > 0) "
csql = csql & ") "
csql = csql & "AND TB_DETSOLICITUD.cod_solicitud='" & Numero_de_Requerimiento & "'"

Set RsV = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
If RsV.RecordCount = 0 Then
    csql = "update TB_CABSOLICITUD set cs_estado='3' where cod_solicitud='" & Numero_de_Requerimiento & "'"
    cnn_dbbancos.Execute csql
Else
    If RsV.RecordCount > 0 Then
        RsV.MoveFirst
        Do While Not RsV.EOF
            If RsV!VBJEFECC = True Then
                csql = "update TB_CABSOLICITUD set cs_estado='2' where cod_solicitud='" & Numero_de_Requerimiento & "'"
            Else
                    csql = "update TB_CABSOLICITUD set cs_estado='1' where cod_solicitud='" & Numero_de_Requerimiento & "'"
            End If
            cnn_dbbancos.Execute csql
            RsV.MoveNext
        Loop
    End If
End If

End Sub







'-------------------------------------------------------------------------------------------------------------
Private Sub listarEstadoEnImageCombo()
    With imgCmbEstado
        .ComboItems.Clear
        
        .ComboItems.Add , , "Orden en Edición", "Estado 1", "Estado 1"
        .ComboItems.Add , , "Orden Aprobada", "Estado 2", "Estado 2"
        .ComboItems.Add , , "Orden Enviada", "Estado 3", "Estado 3"
        .ComboItems.Add , , "Orden Recepcionada", "Estado 4", "Estado 4"
        .ComboItems.Add , , "Atención Parcial", "Estado 5", "Estado 5"
        .ComboItems.Add , , "Atención Total", "Estado 6", "Estado 6"
        .ComboItems.Add , , "Orden Cerrada", "Estado 7", "Estado 7"
        .ComboItems.Add , , "Orden Anulada", "Estado 8", "Estado 8"
        
        .Enabled = False
        .BackColor = DF
    End With
End Sub

Private Sub listarTipoComprobanteEnCombo()
    Dim rstTipoComprobante As New ADODB.Recordset
    
    If rstTipoComprobante.State = 1 Then rstTipoComprobante.Close
    
    rstTipoComprobante.Open "SELECT * FROM DOCUMENTOS WHERE F2TIPO IN ('P', 'A') ORDER BY F2DESDOC", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    CmbTipDoc.Clear
    
    If Not rstTipoComprobante.EOF Then
        rstTipoComprobante.MoveFirst
        
        Do While Not rstTipoComprobante.EOF
            CmbTipDoc.AddItem Trim(rstTipoComprobante!F2DESDOC & "") & Space(100) & Trim(rstTipoComprobante!F2CODDOC & "")
            
            rstTipoComprobante.MoveNext
        Loop
            If CmbTipDoc.ListCount > 0 Then
                CmbTipDoc.ListIndex = 0
            End If
    End If
End Sub

Private Sub listarGrillaOrden()
    With Grid.Dataset
        abrirCnTemporal
        
        .Active = False
        .Refresh
        
        abrirCnTemporal
        
        .ADODataset.ConnectionString = cnDBTemp
        .ADODataset.CommandText = "SELECT * FROM TMPORDENDECOMPRA ORDER BY ITEM"
        .Active = False
        .Active = True
        .Close
        .Open
    End With
End Sub

Private Sub adicionarItemOrden()
    With Grid.Dataset
        .Close
        
        abrirCnTemporal
        
        cnDBTemp.Execute "DELETE * FROM TMPORDENDECOMPRA"

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

'        .FieldValues("ITEM") = .RecordCount + 1
'        .FieldValues("F3CODPRO") = vbNullString
'
'        .Post

        .Close
        .Open
    End With
End Sub

Private Sub limpiarCajasV2()
    Rem SK ADD:
    Select Case strTipoOrden
        Case "OC"
            Me.Caption = "Orden de Compra"
            Label1(12).Caption = "Nº Orden Compra"
        Case "OS"
            Me.Caption = "Orden de Servicio"
            Label1(12).Caption = "Nº Orden Servicio"
    End Select
    
    Txt_TOC.Text = vbNullString
    Txt_NumOC.Text = vbNullString
    imgCmbEstado.ComboItems(1).Selected = True
    
    lblAnulada.Visible = False
    
    txt_fecha.value = Format(Date, "Short Date")
    chkSinProveedorEsp.value = vbUnchecked
    Txt_Prove.Text = vbNullString
        pnlnomprv.Caption = vbNullString
        pnldireprv.Caption = vbNullString
    txtcontacto.Text = vbNullString
    CmbTipDoc.ListIndex = -1
    abofechaentrega.value = Format(Date, "Short Date")
    txtcodsoli.Text = wusuario
    pnlnomsoli.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", wusuario, "T")
    txtcodforma.Text = vbNullString
        pnlnomforma.Caption = vbNullString
    TxtCodCosto.Text = vbNullString
        PnlNomCosto.Caption = vbNullString
    txtlugar_entrega.Text = wdireccion
    
    Cmbmone.ListIndex = ModUtilitario.seleccionarItem(Cmbmone, "S", "IZQ", 1)
    txt_tc.Text = Format(Val(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CAMBIO", "CAMBIOS", "FECHA", txt_fecha.value, "F")), "#.000")
    
    
    Txt_Referencia.Text = vbNullString
    ChK_regularizacion.Checked = False
    Chk_pagoparcial.Checked = False
    
    txtCotizacion.Text = vbNullString
    
    chkOrdenEnviada.value = vbUnchecked
        txtEnviadoPor.Text = vbNullString
        dtpFechaEnvio.value = Format(Date, "Short Date")
        
    chkOrdenEnviada_Click
    
    chkOrdenRecepcionada.value = vbUnchecked
        txtRecepcionadoPor.Text = vbNullString
        dtpFechaRecepcion.value = Format(Date, "Short Date")
        
    chkOrdenRecepcionada_Click
    
    txtempresa.Text = UCase(wnomcia)
    txtobserva.Text = vbNullString
    
    txtbase.Text = "0.00"
    txtmonto.Text = "0.00"
    txtigv.Text = "0.00"
    TxtRnd.Text = "0.00"
    txttotal.Text = "0.00"
    
    chkDetraccionAplicar.value = vbUnchecked
    txtDetraccionPorc.Text = "0.00"
    
    'Activar columnas de Gasto y Cta. Contable
    dxCheckBox1.Checked = IIf(strTipoOrden = "OC", False, True)
        dxCheckBox1_Click
    'Activar 6 decimales en precio unitario
    dxCheckBox2.Checked = False
        dxCheckBox2_Click
    'Visualizar B. Imponible por Item
    dxCheckBox3.Checked = False
        dxCheckBox3_Click
    'Visualizar Porcentaje de Demasia por Item
    dxCheckBox4.Checked = False
    dxCheckBox4.Enabled = IIf(strTipoOrden = "OC", True, False)
        dxCheckBox4_Click
    'Visualizar Descuentos por Item
    dxCheckBox5.Checked = False
        dxCheckBox5_Click
    'Visualizar Observaciones por Item
    dxCheckBox6.Checked = False
        dxCheckBox6_Click
    'Visualizar Descripcion Interna
    dxCheckBox7.Checked = False
        dxCheckBox7_Click
    'Visualizar Cliente de Requerimiento
    dxCheckBox8.Checked = False
    dxCheckBox8.Enabled = IIf(strTipoOrden = "OC", True, False)
        dxCheckBox8_Click
    'Cierre de Orden por Item
    dxCheckBox9.Checked = False
    dxCheckBox9.Enabled = False
        dxCheckBox9_Click
    
    atbmenu.Tools.ITEM("ID_Nuevo").Enabled = True
    atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
    atbmenu.Tools.ITEM("ID_Anular").Enabled = False
    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
    atbmenu.Tools.ITEM("ID_Email").Enabled = False
    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
    atbmenu.Tools.ITEM("Reposicion").Enabled = True
    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
    
    Grid.Columns.ColumnByFieldName("F3DESMEDIDA").Visible = IIf(strTipoOrden = "OC", True, False)
   'Grid.Columns.ColumnByFieldName("F3CANPRO").Visible = IIf(strTipoOrden = "OC", True, False)
    Grid.Columns.ColumnByFieldName("DESCOLOR").Visible = IIf(strTipoOrden = "OC", True, False)
End Sub

Public Sub consultarOrden()
    Set objOrden = New ClsOrden
    
    limpiarCajasV2
    
    Grid.Dataset.Close
    
    With objOrden
        .inicializarEntidades
        
        .TipoOrden = strTipoOrden
        .NumeroOrden = strNumeroOrden
        
        If .obtenerOrden Then
            Txt_TOC.Text = .TipoOrden
            Txt_NumOC.Text = .NumeroOrden
            
            If Val(.Estado) >= 1 And Val(.Estado) <= 8 Then
                imgCmbEstado.ComboItems(Val(.Estado)).Selected = True
            End If
            
            lblAnulada.Visible = .DocAnulado
            
             If Txt_TOC = "OS" Then
                lblAutorizado.Visible = True
                txtAutorizado.Visible = True
                txtAutorizado.Text = Trim("" & .Responsable & "")
            End If
            
            txt_fecha.value = Format(.FechaEmision, "Short Date")
            chkSinProveedorEsp.value = IIf(.SinProveedorEspecifico, vbChecked, vbUnchecked)
                chkSinProveedorEsp.Enabled = False
                Txt_Prove.Enabled = Not .SinProveedorEspecifico
                txtcontacto.Enabled = Not .SinProveedorEspecifico
                
            Txt_Prove.Text = .RucProveedor
                pnlnomprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMPROV", "EF2PROVEEDORES", "F2NEWRUC", .RucProveedor, "T")
                pnldireprv.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DIRPROV", "EF2PROVEEDORES", "F2NEWRUC", .RucProveedor, "T")
                txtcontacto.Text = .ContactoProveedor
                
            CmbTipDoc.ListIndex = ModUtilitario.seleccionarItem(CmbTipDoc, .CodTipoComprobante, "DER", 2)
            abofechaentrega.value = Format(.FechaEntrega, "Short Date")
            
            txtcodsoli.Text = .CodigoSolicitante
                pnlnomsoli.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2NOMUSER", "EF2USERS", "F2CODUSER", .CodigoSolicitante, "T")
                
            txtcodforma.Text = .CodFormaPago
                pnlnomforma.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", .CodFormaPago, "T")
            TxtCodCosto.Text = .CentroCosto
                PnlNomCosto.Caption = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", .CentroCosto, "T")
            txtlugar_entrega.Text = .LugarEntrega
            
            Cmbmone.ListIndex = ModUtilitario.seleccionarItem(Cmbmone, .CodMoneda, "IZQ", 1)
            txt_tc.Text = Format(Val(.TipoCambio), "#.000")
            
            Txt_Referencia.Text = .referencia
            ChK_regularizacion.Checked = .OrdenRegularizada
            Chk_pagoparcial.Checked = .PagoParcial
            txtCotizacion.Text = .NumeroCotizacion
            
            chkOrdenEnviada.value = IIf(CBool(.Colocada), vbChecked, vbUnchecked)
                txtEnviadoPor.Text = .ColocadaUsuario
                dtpFechaEnvio.value = IIf(.ColocadaFecha <> vbNullString, Format(.ColocadaFecha, "Short Date"), Date)
                
            chkOrdenEnviada_Click
            
            chkOrdenRecepcionada.value = IIf(CBool(.Atendida), vbChecked, vbUnchecked)
                txtRecepcionadoPor.Text = .AtendidaUsuario
                dtpFechaRecepcion.value = IIf(.AtendidaFecha <> vbNullString, Format(.AtendidaFecha, "Short Date"), Date)
            
            chkOrdenRecepcionada_Click
            
            txtempresa.Text = .Empresa
            txtobserva.Text = .Observacion
            
            listarGrillaOrden
            
            If Grid.Dataset.RecordCount = 0 Then
                Grid.Dataset.Close
                
                adicionarItemOrden
            End If
            
            Select Case Val(.Estado)
                Case 1 'Orden en Edición
                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = True
                    atbmenu.Tools.ITEM("ID_Anular").Enabled = True
                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
                    atbmenu.Tools.ITEM("ID_Email").Enabled = False
                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
                Case 2, 3 'Orden Aprobada / Orden Enviada / 'Orden Recepcionada
                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Anular").Enabled = True
                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
                    atbmenu.Tools.ITEM("ID_Email").Enabled = True
                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
                Case 4 'Orden Recepcionada
'                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = True
'                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
'                    atbmenu.Tools.ITEM("ID_Anular").Enabled = True
'                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
'                    atbmenu.Tools.ITEM("ID_Email").Enabled = True
'                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = True
'                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
                Case 5 'Atención Parcial
                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Anular").Enabled = False
                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
                    atbmenu.Tools.ITEM("ID_Email").Enabled = False
                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = True
                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = True
                    dxCheckBox9.Enabled = True
                Case 6 'Atención Total
                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Anular").Enabled = False
                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
                    atbmenu.Tools.ITEM("ID_Email").Enabled = False
                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
                Case 7 'Orden Cerrada
                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Anular").Enabled = False
                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = True
                    atbmenu.Tools.ITEM("ID_Email").Enabled = False
                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
                Case 8 'Orden Anulada
                    atbmenu.Tools.ITEM("ID_Grabar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Eliminar").Enabled = False
                    atbmenu.Tools.ITEM("ID_Anular").Enabled = False
                    atbmenu.Tools.ITEM("ID_Imprimir").Enabled = False
                    atbmenu.Tools.ITEM("ID_Email").Enabled = False
                    atbmenu.Tools.ITEM("ID_Cerrar").Enabled = False
                    atbmenu.Tools.ITEM("ID_ModificarOrden").Enabled = False
            End Select
            
            atbmenu.Tools.ITEM("Reposicion").Enabled = False
            
'            atbmenu.Tools.ITEM("ID_Grabar").enabled = Not .DocAnulado
'            atbmenu.Tools.ITEM("ID_Anular").enabled = Not .DocAnulado
'            atbmenu.Tools.ITEM("ID_Eliminar").enabled = Not .DocAnulado
'            atbmenu.Tools.ITEM("ID_Imprimir").enabled = Not .DocAnulado
'            atbmenu.Tools.ITEM("ID_Email").enabled = Not .DocAnulado
            
            txtbase.Text = Format(.SUBTOTAL, "#,0.00")
            txtmonto.Text = Format(.TotalInafecto, "#,0.00")
            txtigv.Text = Format(.TotalImpuesto, "#,0.00")
            TxtRnd.Text = Format(.TotalRedondeo, "#,0.00")
            txttotal.Text = Format(.TotalFacturado, "#,0.00")
        Else
            Txt_TOC.Text = strTipoOrden
            
            listarGrillaOrden
            
            adicionarItemOrden
        End If
    End With
    
    Set objOrden = Nothing
End Sub

Private Sub validarCajas()
    If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
        Grid.Dataset.Post
    End If
    
    If Trim(Txt_TOC.Text) = vbNullString Then
        MsgBox "Tipo de Orden indefinido, comuniquese con su administrador de sistemas.", vbInformation + vbOKOnly, App.ProductName

        Txt_TOC.SetFocus
        
        Exit Sub
    End If
    
    If Trim(Txt_Prove.Text) = vbNullString Then
        MsgBox "El Campo Proveedor es obligatorio.", vbInformation + vbOKOnly, App.ProductName

        Txt_Prove.SetFocus
        
        Exit Sub
    End If
    
    If CmbTipDoc.ListIndex = -1 Then
        MsgBox "El Campo Tipo de Documento es obligatorio.", vbInformation + vbOKOnly, App.ProductName

        CmbTipDoc.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtcodsoli.Text) = vbNullString Or Trim(pnlnomsoli.Caption) = vbNullString Then
        MsgBox "El Campo Solicitante es obligatorio.", vbInformation + vbOKOnly, App.ProductName
        
        txtcodsoli.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtcodforma.Text) = vbNullString Or Trim(pnlnomforma.Caption) = vbNullString Then
        MsgBox "El Campo Forma de Pago es obligatorio.", vbInformation + vbOKOnly, App.ProductName
        
        txtcodforma.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtlugar_entrega.Text) = vbNullString Then
        If MsgBox("No ha ingresado el Lugar de Entrega, ¿Desea continuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
            txtlugar_entrega.SetFocus
            
            Exit Sub
        End If
    End If
    
    If Cmbmone.ListIndex = -1 Then
        MsgBox "El Campo Moneda es obligatorio.", vbInformation + vbOKOnly, App.ProductName
        
        Cmbmone.SetFocus
        
        Exit Sub
    End If
    
    If Val(txt_tc.Text) <= 0 Then
        MsgBox "Tipo de Cambio invalido, verifique.", vbInformation + vbOKOnly, App.ProductName
        
        txt_tc.SetFocus
        
        Exit Sub
    End If
    
    If fraSeguimiento.Enabled Then
        If CBool(chkOrdenEnviada.value) Then
            If Trim(txtEnviadoPor.Text) = vbNullString Then
                MsgBox "El Campo 'Enviada Por' es obligatorio.", vbInformation + vbOKOnly, App.ProductName
                
                txtEnviadoPor.SetFocus
                
                Exit Sub
            End If
            
            If CDate(dtpFechaEnvio.value) < CDate(txt_fecha.value) Then
                MsgBox "Fecha de Envio no puede ser menor a la Fecha de O/C, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                dtpFechaEnvio.SetFocus
                
                Exit Sub
            End If
        End If
        
        If CBool(chkOrdenRecepcionada.value) Then
            If Trim(txtRecepcionadoPor.Text) = vbNullString Then
                MsgBox "El Campo 'Recepcionada Por' es obligatorio.", vbInformation + vbOKOnly, App.ProductName
                
                txtRecepcionadoPor.SetFocus
                
                Exit Sub
            End If
            
            If CDate(dtpFechaRecepcion.value) < CDate(txt_fecha.value) Then
                MsgBox "Fecha de Envio no puede ser menor a la Fecha de la Orden, verifique.", vbInformation + vbOKOnly, App.ProductName
                
                dtpFechaRecepcion.SetFocus
                
                Exit Sub
            End If
        End If
    End If
    
    If Val(ModUtilitario.ObtenerCampoV2(cnDBTemp, "COUNT(F3SINIGV) AS CANTIDAD", "TMPORDENDECOMPRA", "F3SINIGV", "0", "N", "AND TRIM(F3CODPRO & '') <> '' GROUP BY F3SINIGV")) > 0 Then
        If MsgBox("Se han detectado Items sin Precio ingresado, ¿Desea continuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
            Grid.SetFocus
            
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Desea guardar el registro?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        guardarOrden
    End If
End Sub

Private Sub guardarOrden()
    On Error GoTo errGuardarOrden
    
    Dim rstTemporalGuardarOrden As New ADODB.Recordset
    
    Set objOrden = New ClsOrden
    
    Grid.Dataset.Close
    
    abrirCnTemporal
    
    With objOrden
        .inicializarEntidades
        
        .TipoOrden = Trim(Txt_TOC.Text)
        .NumeroOrden = Trim(Txt_NumOC.Text)
        
        .FechaEmision = Format(txt_fecha.value, "Short Date")
        .SinProveedorEspecifico = CBool(chkSinProveedorEsp.value)
        .NomProveedor = Replace(Trim(pnlnomprv.Caption), "'", "' & Chr(39) & '", 1)
        .RucProveedor = Trim(Txt_Prove.Text)
        .CodProveedor = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODPROV", "EF2PROVEEDORES", "F2NEWRUC", .RucProveedor, "T")
        .ContactoProveedor = Replace(Trim(txtcontacto.Text), "'", "' & Chr(39) & '", 1)
        
        .CodTipoComprobante = right(CmbTipDoc.Text, 2)
        .OrdenRegularizada = CBool(ChK_regularizacion.Checked)

        .FechaEntrega = Format(abofechaentrega.value, "Short Date")
        .CodigoSolicitante = Trim(txtcodsoli.Text)
        .CodFormaPago = Trim(txtcodforma.Text)
        .CentroCosto = Trim(TxtCodCosto.Text)
        .LugarEntrega = Trim(txtlugar_entrega.Text)
        .PagoParcial = CBool(Chk_pagoparcial.Checked)

        .CodMoneda = left(Cmbmone.Text, 1)
        .TipoCambio = Format(Val(txt_tc.Text), "#.000")
        .NumeroCotizacion = Trim(txtCotizacion.Text)

        .Colocada = CBool(chkOrdenEnviada.value)
            .ColocadaUsuario = Trim(txtEnviadoPor.Text)
            .ColocadaFecha = IIf(.Colocada, Format(dtpFechaEnvio.value, "Short Date"), vbNullString)

        .Atendida = CBool(chkOrdenRecepcionada.value)
            .AtendidaUsuario = Trim(txtRecepcionadoPor.Text)
            .AtendidaFecha = IIf(.Atendida, Format(dtpFechaRecepcion.value, "Short Date"), vbNullString)
        
        .Empresa = Trim(txtempresa.Text)
        .Observacion = Trim(txtobserva.Text)

        .SUBTOTAL = Val(Format(Grid.Columns.ColumnByFieldName("F3BASEIMP").SummaryFooterValue, "0.00"))
        .TotalInafecto = Val(Format(Grid.Columns.ColumnByFieldName("F3MONINA").SummaryFooterValue, "0.00"))
        .TotalImpuesto = Val(Format(Grid.Columns.ColumnByFieldName("F3IGV").SummaryFooterValue, "0.00"))
        .TotalFacturado = Val(Format(Grid.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "0.00"))
        
        .FechaReg = Format(Date, "Short Date")
        .UsuarioReg = wusuario
        .FechaMod = Format(Date, "Short Date")
        .UsuarioMod = wusuario
        
        .Estado = imgCmbEstado.SelectedItem.Index
        
        .Responsable = txtAutorizado.Text
        
        If imgCmbEstado.SelectedItem.Index = 3 And .Atendida Then
            .Estado = 4
        End If
        
        If .guardarOrden Then
            Actualiza_Log .SQLSelectAlter, StrConexDbBancos

            .SQLSelectAlter = "DELETE FROM IF3ORDEN WHERE F4LOCAL = '" & .TipoOrden & "' AND F4NUMORD = '" & .NumeroOrden & "'"
            
            cnn_dbbancos.Execute .SQLSelectAlter

            Actualiza_Log .SQLSelectAlter, StrConexDbBancos

            If rstTemporalGuardarOrden.State = 1 Then rstTemporalGuardarOrden.Close

            rstTemporalGuardarOrden.Open "SELECT * FROM TMPORDENDECOMPRA WHERE TRIM(F3CODPRO & '') <> '' ORDER BY ITEM", cnDBTemp, adOpenForwardOnly, adLockReadOnly 'AND VAL(F3CANPRO & '') > 0
            
            If Not rstTemporalGuardarOrden.EOF Then
                rstTemporalGuardarOrden.MoveFirst
                
                Do While Not rstTemporalGuardarOrden.EOF
                    .inicializarEntidadesDetalle

                    .ITEM = Val(rstTemporalGuardarOrden!ITEM & "")
                    .Requerimiento = Trim(rstTemporalGuardarOrden!COD_SOLICITUD & "")
                    .CodigoProducto = Trim(rstTemporalGuardarOrden!F3CODPRO & "")
                    .CodigoFabricante = Trim(rstTemporalGuardarOrden!F3CODFAB & "")
                    .NombreProducto = Trim(rstTemporalGuardarOrden!F5NOMPRO & "")
                    .NombreProductoInterno = Trim(rstTemporalGuardarOrden!F5NOMPRO_ING & "")
                    .CodigoUM = Trim(rstTemporalGuardarOrden!F3CODMEDIDA & "")
                    .Cantidad = Val(rstTemporalGuardarOrden!F3CANPRO & "")
                    .CantidadMaxima = Val(rstTemporalGuardarOrden!F3CANPROMAX & "")
                    .CantidadFaltante = Val(rstTemporalGuardarOrden!CANT_ANT & "")
                    
                    .PorcentajeDemasia = Val(rstTemporalGuardarOrden!F3PORCDEMASIA & "")
                    
                    .PrecioSinImpuesto = Val(rstTemporalGuardarOrden!F3SINIGV & "")
                    .PrecioConImpuesto = Val(rstTemporalGuardarOrden!F3CONIGV & "")
                    .PrecioNetoSinImpuesto = Val(rstTemporalGuardarOrden!F3NETO & "")
                    
                    .PorcentajeDscto = Val(rstTemporalGuardarOrden!F3PORDESC & "")
                    .TotalDscto = Val(rstTemporalGuardarOrden!F3VALDESC & "")
                    
                    .Afecto = CBool(rstTemporalGuardarOrden!F5AFECTO)
                    
                    .BasePorItem = Val(rstTemporalGuardarOrden!F3BASEIMP & "")
                    .ImpuestoPorItem = Val(rstTemporalGuardarOrden!F3IGV & "")
                    .TotalPorItem = Val(rstTemporalGuardarOrden!F3TOTAL & "")
                    
                    .CodigoColor = Trim(rstTemporalGuardarOrden!CODCOLOR & "")
                    .ObservacionPorItem = Trim(rstTemporalGuardarOrden!F3OBSERVA & "")
                    
                    .ItemAjustado = CBool(rstTemporalGuardarOrden!CERRAR)
                    
                    .CodigoGasto = Trim(rstTemporalGuardarOrden!F3GASTO & "")
                    .CuentaContable = Trim(rstTemporalGuardarOrden!F3CUENTA & "")
                    
                    .guardarOrdenDetalleOneByOne
                    
                    Actualiza_Log .SQLSelectAlter, StrConexDbBancos

                    rstTemporalGuardarOrden.MoveNext
                Loop
            End If
            
            strTipoOrden = .TipoOrden
            strNumeroOrden = .NumeroOrden
            
            
            consultarOrden
            
            MsgBox "Registro guardado.", vbInformation + vbOKOnly, App.ProductName
        Else
            listarGrillaOrden
        End If
    End With
    
    Set objOrden = Nothing
    
    Exit Sub
errGuardarOrden:
    MsgBox "No.: " & Err.Number & vbNewLine & "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Actualiza_Log "Registro no guardado correctamente: No.: " & Err.Number & " / Descripción: " & Err.Description, StrConexDbBancos
    
    Err.Clear
End Sub

Private Sub eliminarOrden()
    Set objOrden = New ClsOrden
    
    With objOrden
        .TipoOrden = Trim(Txt_TOC.Text)
        .NumeroOrden = Trim(Txt_NumOC.Text)
        
        .obtenerConfigOrden
        
        If Not .verificarExistencia Then
            MsgBox "Orden no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        'Verificar si la Orden de Compra no tiene Movimiento.
        If .obtenerMovimientoOrdenEnAlmacen Then
            MsgBox "Orden no puede ser eliminada, cuenta con" & vbNewLine & _
                    "movimiento en Almacen, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        'Verificar si la Orden, ya fue atendida por el Proveedor.
        If .Atendida Then
            If MsgBox("La Orden esta marcada como Atendida por " & .AtendidaUsuario & vbNewLine & _
                        "con fecha " & .AtendidaFecha & "." & vbNewLine & _
                        "¿Desea continuar con la acción?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If MsgBox("¿Desea eliminar la Orden con No. " & .NumeroOrden & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            If .eliminarOrden Then
                Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                
                strTipoOrden = .TipoOrden
                strNumeroOrden = .NumeroOrden
                
                consultarOrden
                
                MsgBox "Registro eliminado.", vbInformation + vbOKOnly, App.ProductName
            End If
        End If
    End With

    Set objOrden = Nothing
End Sub

Private Sub anularOrden()
    Set objOrden = New ClsOrden
    
    With objOrden
        .TipoOrden = Trim(Txt_TOC.Text)
        .NumeroOrden = Trim(Txt_NumOC.Text)
        
        .obtenerConfigOrden
        
        If Not .verificarExistencia Then
            MsgBox "Orden no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        'Verificar si la Orden de Compra no tiene Movimiento.
        If .obtenerMovimientoOrdenEnAlmacen Then
            MsgBox "Orden no puede ser anulada, cuenta con" & vbNewLine & _
                    "movimiento en Almacen, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        'Verificar si la Orden, ya fue atendida por el Proveedor.
        If .Atendida Then
            If MsgBox("La Orden esta marcada como Atendida por " & .AtendidaUsuario & vbNewLine & _
                        "con fecha " & .AtendidaFecha & "." & vbNewLine & _
                        "¿Desea continuar con la acción?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If MsgBox("¿Desea anular la Orden con No. " & .NumeroOrden & "?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            .UsuarioMod = wusuario
            .FechaMod = Format(Date, "Short Date")
            
            If .anularOrden Then
                strTipoOrden = .TipoOrden
                strNumeroOrden = .NumeroOrden
                
                consultarOrden
                
                MsgBox "Orden con Nº " & .NumeroOrden & " anulada.", vbInformation + vbOKOnly, App.ProductName
            End If
        End If
    End With

    Set objOrden = Nothing
End Sub

Private Sub enviarViaMailOrden()
'''    If Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "LCASE(F2EMAIL)", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text & ""), "T")) = vbNullString Or _
'''        Trim(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "LCASE(F2EMAIL)", "EF2PROVEEDORES", "F2NEWRUC", Trim(Txt_Prove.Text & ""), "T")) = "@" Then
'''
'''        SqlCad = vbNullString
'''        SqlCad = Trim(InputBox("Proveedor sin e-mail configurado," & vbNewLine & _
'''                        "ingrese la cuenta de correo:", "Envió de E-mail", "example@dominio.com"))
'''
'''        If SqlCad = vbNullString Or SqlCad = "example@dominio.com" Then
'''            MsgBox "Imposible enviar e-mail, Proveedor sin cuenta de correo configurada.", vbInformation + vbOKOnly, App.ProductName
'''
'''            Exit Sub
'''        End If
'''
'''        SqlCad = "UPDATE EF2PROVEEDORES SET F2EMAIL = '" & SqlCad & "' WHERE F2NEWRUC = '" & Trim(Txt_Prove.Text & "") & "'"
'''
'''        cnn_dbbancos.Execute SqlCad
'''
'''        Actualiza_Log SqlCad, StrConexDbBancos
'''    End If
    
    If MsgBox("Guarde los cambios antes de Enviar." & vbNewLine & vbNewLine & _
                "¿Desea enviar la orden vía e-mail?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                
        'Imprime_Orden_ParaEnvioMail 1
        imprimeOrdenV2 strTipoOrden, strNumeroOrden, True
    End If
End Sub

Private Sub cerrarOrden()
    Set objOrden = New ClsOrden
    
    With objOrden
        .TipoOrden = Trim(Txt_TOC.Text)
        .NumeroOrden = Trim(Txt_NumOC.Text)
        
        .obtenerConfigOrden
        
        If Not .verificarExistencia Then
            MsgBox "Orden no existente.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
        'Verificar si la Orden de Compra tiene Movimiento.
        If Not .obtenerMovimientoOrdenEnAlmacen Then
            MsgBox "Orden no puede ser cerrada, no cuenta con" & vbNewLine & _
                    "movimiento en Almacen, verifique.", vbInformation + vbOKOnly, App.ProductName
            
            Exit Sub
        End If
        
'        'Verificar si la Orden, ya fue atendida por el Proveedor.
'        If .Atendida Then
'            If MsgBox("La Orden esta marcada como Atendida por " & .AtendidaUsuario & vbNewLine & _
'                        "con fecha " & .AtendidaFecha & "." & vbNewLine & _
'                        "¿Desea continuar con la acción?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'
'                Exit Sub
'            End If
'        End If
        
'        If MsgBox("Confirme el Tipo de Cierre a efectuar:" & vbNewLine & vbNewLine & _
'                    "SI: Cierre Total." & vbNewLine & _
'                    "NO: Cierre Parcial." & vbNewLine & vbNewLine & _
'                    "ATENCIÓN: Usar el Cierre Parcial de Orden solo en caso se desestime la entrega, en coordinación con el proveedor, de algunos Items y necesariamente otros queden Pendiente de Entrega; caso contrario use el Cierre Total de la Orden.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
'
'
'        Else
            If MsgBox("¿Desea cerrar la Orden con No. " & .NumeroOrden & "?" & vbNewLine & vbNewLine & _
                        "ATENCIÓN: " & vbNewLine & _
                        "Verifique que los Ingresos por Compra se encuentren actualizados antes de proceder con el cierre de la Orden.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                        
                .UsuarioMod = wusuario
                .FechaMod = Format(Date, "Short Date")
                
                If .cerrarOrden Then
                    strTipoOrden = .TipoOrden
                    strNumeroOrden = .NumeroOrden
                    
                    consultarOrden
                    
                    MsgBox "Orden con Nº " & .NumeroOrden & " cerrada.", vbInformation + vbOKOnly, App.ProductName
                End If
            End If
'        End If
    End With

    Set objOrden = Nothing
End Sub

Private Sub actualizarSaldosPorEntregarDeProductos()
    Dim rstAtencion As New ADODB.Recordset
    
    Grid.Dataset.Close
    
    abrirCnTemporal
    
    With objAyudaOrden
        .TipoOrden = strTipoOrden
        .NumeroOrden = strNumeroOrden
    End With
    
    If rstAtencion.State = 1 Then rstAtencion.Close
    
    rstAtencion.Open objAyudaOrden.devuelveSQLAtencionOrden(True), cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstAtencion.EOF Then
        rstAtencion.MoveFirst
        
        Do While Not rstAtencion.EOF
            CadSql = vbNullString
            CadSql = CadSql & "UPDATE TMPORDENDECOMPRA "
            CadSql = CadSql & "SET "
            CadSql = CadSql & "PORENTREGAR = " & Val(rstAtencion!SALDO & "") & " "
            CadSql = CadSql & "WHERE "
            CadSql = CadSql & "COD_SOLICITUD = '" & Trim(rstAtencion!COD_SOLICITUD & "") & "' AND "
            CadSql = CadSql & "F3CODPRO = '" & Trim(rstAtencion!F3CODPRO & "") & "'"
            
            cnDBTemp.Execute CadSql
            
            rstAtencion.MoveNext
        Loop
    End If
    
    CadSql = vbNullString
    
    If rstAtencion.State = 1 Then rstAtencion.Close
    
    Set rstAtencion = Nothing
    
    Me.MousePointer = vbDefault
End Sub




