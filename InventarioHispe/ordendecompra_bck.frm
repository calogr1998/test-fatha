VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{791923BA-56CB-4A36-9EA3-1B4ED74622AA}#1.0#0"; "csimxctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ordendecompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   8925
   ClientLeft      =   1005
   ClientTop       =   1155
   ClientWidth     =   18825
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   18825
   Begin VB.Frame Frame4 
      Height          =   1530
      Left            =   0
      TabIndex        =   41
      Top             =   7380
      Width           =   18600
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11985
         MaxLength       =   100
         TabIndex        =   73
         Text            =   "0.00"
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txtempresa 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   13
         Top             =   300
         Width           =   6270
      End
      Begin VB.TextBox txtobserva 
         Height          =   675
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   8775
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10320
         TabIndex        =   74
         Top             =   300
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   42
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame FrameOC 
      Height          =   1425
      Left            =   60
      TabIndex        =   18
      Top             =   480
      Width           =   11955
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
         Left            =   9120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   240
         Width           =   405
      End
      Begin VB.ComboBox CmbTipDoc 
         Height          =   330
         Left            =   9060
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1020
         Width           =   2835
      End
      Begin VB.TextBox Txt_Prove 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   1185
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
         Left            =   9540
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   2325
      End
      Begin VB.TextBox Txt_Referencia 
         Height          =   315
         Left            =   3480
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   -120
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.TextBox txtcontacto 
         Height          =   315
         Left            =   9060
         TabIndex        =   1
         Top             =   660
         Width           =   2820
      End
      Begin VB.TextBox txtusuario 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6420
         TabIndex        =   15
         Top             =   -120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txt_NumSolComp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9060
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1500
         Visible         =   0   'False
         Width           =   2535
      End
      Begin Threed.SSPanel pnldireprv 
         Height          =   270
         Left            =   1560
         TabIndex        =   19
         Top             =   990
         Width           =   5970
         _Version        =   65536
         _ExtentX        =   10530
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
         Left            =   2760
         TabIndex        =   57
         Top             =   600
         Width           =   4770
         _Version        =   65536
         _ExtentX        =   8414
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
         Left            =   1560
         TabIndex        =   69
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90112001
         CurrentDate     =   40611
      End
      Begin CONTROLSLibCtl.dxCheckBox ChK_regularizacion 
         Height          =   270
         Left            =   5760
         TabIndex        =   66
         Top             =   240
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
         TextLayout      =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   7620
         TabIndex        =   63
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
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
         Left            =   7380
         TabIndex        =   54
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   7620
         TabIndex        =   44
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   6300
         TabIndex        =   23
         Top             =   -60
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Requerimiento"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   7620
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   990
         Width           =   675
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
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   2100
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   18825
      _Version        =   65536
      _ExtentX        =   33205
      _ExtentY        =   741
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
      BevelOuter      =   0
      Alignment       =   6
      Begin InternetMailCtl.InternetMail InternetMail 
         Left            =   11040
         Top             =   0
         _cx             =   741
         _cy             =   741
         Enabled         =   -1  'True
      End
      Begin ActiveToolBars.SSActiveToolBars atbmenu 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   9
         Tools           =   "ordendecompra.frx":000C
         ToolBars        =   "ordendecompra.frx":71C9
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   1908
      Left            =   120
      TabIndex        =   25
      Top             =   1860
      Width           =   15075
      Begin VB.TextBox txtFechaPago 
         Height          =   315
         Left            =   10620
         TabIndex        =   10
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txtlugar_entrega 
         Height          =   315
         Left            =   7260
         MaxLength       =   100
         TabIndex        =   8
         Top             =   960
         Width           =   7635
      End
      Begin VB.TextBox TxtCodCosto 
         Height          =   312
         Left            =   1440
         TabIndex        =   9
         Top             =   1380
         Width           =   1020
      End
      Begin VB.TextBox txtcodsoli 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txtcodforma 
         Height          =   312
         Left            =   1440
         TabIndex        =   7
         Top             =   1020
         Width           =   1020
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   315
         Left            =   13380
         TabIndex        =   11
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txt_tc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   14040
         TabIndex        =   4
         Text            =   "2.7"
         Top             =   600
         Width           =   795
      End
      Begin VB.ComboBox Cmbmone 
         Height          =   330
         ItemData        =   "ordendecompra.frx":73CE
         Left            =   10620
         List            =   "ordendecompra.frx":73D8
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1416
      End
      Begin Threed.SSPanel pnlnomsoli 
         Height          =   300
         Left            =   2520
         TabIndex        =   26
         Top             =   660
         Width           =   6615
         _Version        =   65536
         _ExtentX        =   11668
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
         TabIndex        =   27
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
         TabIndex        =   58
         Top             =   1380
         Width           =   6615
         _Version        =   65536
         _ExtentX        =   11668
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
         TabIndex        =   70
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90112001
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker aBoHoraEntrega 
         Height          =   315
         Left            =   3120
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   90112002
         CurrentDate     =   40611
      End
      Begin CONTROLSLibCtl.dxCheckBox Chk_pagoparcial 
         Height          =   270
         Left            =   9000
         TabIndex        =   65
         Top             =   240
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
         TextLayout      =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Pago"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   9360
         TabIndex        =   64
         Top             =   1380
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Entrega"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5940
         TabIndex        =   62
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   120
         TabIndex        =   59
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
         TabIndex        =   45
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Cotización"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   12240
         TabIndex        =   32
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00000000&
         Height          =   216
         Left            =   132
         TabIndex        =   31
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda "
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   19
         Left            =   9360
         TabIndex        =   29
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cambio"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   17
         Left            =   12540
         TabIndex        =   28
         Top             =   660
         Width           =   1200
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid Grid 
      Height          =   3525
      Left            =   0
      OleObjectBlob   =   "ordendecompra.frx":73EC
      TabIndex        =   12
      Top             =   3780
      Width           =   18495
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   1365
      Left            =   0
      OleObjectBlob   =   "ordendecompra.frx":FEC4
      TabIndex        =   60
      Top             =   6060
      Visible         =   0   'False
      Width           =   11715
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   4800
      TabIndex        =   33
      Top             =   7380
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txtigv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   480
         Width           =   1290
      End
      Begin VB.TextBox txtmonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   480
         Width           =   1320
      End
      Begin VB.TextBox txtbase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtRnd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblmoneda 
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
         Left            =   5940
         TabIndex        =   53
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblmoneda 
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
         TabIndex        =   52
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblmoneda 
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
         Left            =   3360
         TabIndex        =   34
         Top             =   240
         Width           =   240
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
         Left            =   4260
         TabIndex        =   46
         Top             =   240
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
         TabIndex        =   40
         Top             =   240
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
         TabIndex        =   39
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label11 
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
         TabIndex        =   38
         Top             =   240
         Width           =   450
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
         Left            =   5280
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblmoneda 
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
         TabIndex        =   36
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblmoneda 
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
         TabIndex        =   35
         Top             =   240
         Width           =   240
      End
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox2 
      Height          =   270
      Left            =   15480
      TabIndex        =   68
      Top             =   3360
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
      Left            =   15480
      TabIndex        =   67
      Top             =   3000
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
      Left            =   7380
      TabIndex        =   17
      Top             =   60
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
'variables mail
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
Dim fecha                   As String
Dim existe As Boolean
Dim SwRenovar               As Boolean
Dim wNumOc                  As String
Dim CtaPuntos               As Integer, i As Integer
Private cImgInfo As cImageInfo

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
    With Acr_OrdenCompra
        Set cImgInfo = New cImageInfo
   ' MsgBox "Acr_OrdenC_Otros"
        If Cmbmone.ListIndex = 0 Then
            .LblTotF.Caption = "Total " & "S/."
        Else
            .LblTotF.Caption = "Total " & "US$"

        End If
        .flddirec1.Text = wf1direc1
        .FldTelf.Text = "Teléfono: " & wtelefono & " // Fax: " & wfax
        .LblCentroCosto.Text = PnlNomCosto.Caption
        '.flddirec2.Text = wf1direc2
        nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
        .fldruc.Text = "R.U.C. " & wrucempresa
        If FileExist(App.Path & "\" & wrucempresa & ".jpg") = True Then
            .fldempresa.Visible = False
            .ImageLogo.Visible = True
            .ImageLogo.Picture = LoadPicture(App.Path & "\" & wrucempresa & ".jpg")
            With cImgInfo
                .ReadImageInfo App.Path & "\" & wrucempresa & ".jpg"
                
                Acr_OrdenCompra.ImageLogo.Height = 850
                Acr_OrdenCompra.ImageLogo.Width = 850 * .Width / .Height
            End With
            .ImageLogo.top = 0
            .ImageLogo.left = 0
        Else
            .fldempresa.Visible = True
            .ImageLogo.Visible = False
            .fldempresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
        End If
        
        '.IGV.Caption = wigv
        GOC = Txt_NumOC.Text
        If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
        sql = "SELECT A.F4NUMORD,A.F4NUMCOTIZA,A.F4ESTNUL, A.F4CODSOLICITUD,A.F4TIPDOC, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO, A.F4REGULARIZA,A.F4DIAPAGO,  B.F2TELPROV,  B.F2FAXPROV, " & _
              "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4RND,A.F4OBSERVA, " & _
              "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4ESTADO,A.F4FECENT,A.F4OBSERVA,A.F4CODPRV,A.F4TIPMON,A.F4REFERE,A.F4TIPCAM,A.F4FECGRA,A.F4USEGRA,A.F4FECMOD,A.F4USEMOD " & _
              " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & GOC & _
              "' AND A.F4LOCAL='" & TOC & "' ORDER BY A.F4NUMORD DESC"
    
        RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RSCONSULTA.EOF Then
            
            '.F4NUMORD.Text = ("" & rsconsulta.Fields("F4NUMORD"))
            If left(GOC, 2) = "OC" Then
                .LblTitle.Caption = "ORDEN DE COMPRA"
                .Label120.Caption = "-Recepción de facturas en Oficina Central los días Lunes y Miércoles; entre las 08:15 – 13:00 y 14:00 – 18:00."
                .Label121.Caption = "-Se debe adjuntar Guía de Remisión Original (Destinatario y Sunat) con fecha, sello, nombre y firma de recepción de almacén en obra y copia de O/C y certificado de calidad y/o garantía de los productos entregados."
                .Label122.Caption = "-Las guías que componen la factura deben estar asociadas a solo una O/C."
                .Label123.Caption = "-En caso de estar afecto a detracción, colocar el sello legible con el número de cuenta respectivo."
                .Label124.Caption = "-La fecha de vencimiento de las facturas, se consideran a partir de la fecha de recepción en Oficina Central."
                .Label125.Caption = "-El pago de facturas se realizara el día Viernes de la semana correspondiente a la fecha de vencimiento."
                .Label126.Caption = "-Se recibirán facturas emitidas en el mes anterior hasta el 2° día hábil del mes en curso."
                .Label127.Caption = "-La empresa se reserva el derecho de devolver la mercadería que no cumpla con las especificaciones solicitadas."
            Else
                .LblTitle.Caption = "ORDEN DE SERVICIO"
                .Label120.Caption = "-Recepción de facturas de CONTRATISTAS en Oficina Central el día Lunes entre las 08:15 – 13:00 y 14:00 – 18:00."
                .Label121.Caption = "-Los Contratistas deberán adjuntar la Valorización emitida por obra con fecha, sello, nombre y firma de autorización para programación pago y copia de la O/S. En caso de ser primer adelanto, adjuntar CONTRATO debidamente firmado"
                .Label122.Caption = "-En caso de estar afecto a detracción, colocar el sello legible con el número de cuenta respectivo."
                .Label123.Caption = "-Recepción de facturas de Proveedores en Oficina Central los días Lunes y Miércoles; entre las 08:15 – 13:00 y 14:00 – 18:00. "
                .Label124.Caption = "-Los PROVEEDORES deberán adjuntar copia de la O/S y la Guía de Remisión (Destinatario y Sunat) con sello, nombre y firma de almacén por la conformidad en la recepción de los productos en obra."
                .Label125.Caption = "-La fecha de  vencimiento de las facturas se consideran a partir de la fecha de recepción en Oficina Central."
                .Label126.Caption = "-El pago de facturas se realizará el día Viernes de la semana correspondiente a la fecha de vencimiento."
                .Label127.Caption = "-Se recibirán facturas emitidas en el mes anterior hasta el 2° día hábil del mes en curso."

            End If
            .LblNroOC.Caption = "N° " & RSCONSULTA.Fields("F4NUMORD")
            .fldsolicitud.Text = "" & RSCONSULTA.Fields("F4CODSOLICITUD")
            If Len(Trim(RSCONSULTA.Fields("F4USEGRA") & "")) > 0 Then
                .LblCrea.Caption = "Creación Usuario: " & RSCONSULTA.Fields("F4USEGRA") & " (" & RSCONSULTA.Fields("F4fecGRA") & ")"
            Else
                .LblCrea.Caption = ""
            End If
            If Len(Trim(RSCONSULTA.Fields("F4USEmod") & "")) > 0 Then
                .LblModifica.Caption = "Último Usuario: " & RSCONSULTA.Fields("F4USEmod") & " (" & RSCONSULTA.Fields("F4fecmod") & ")"
            Else
                .LblModifica.Caption = ""
            End If
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
            .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
            .FldFchEntrega = "" & Format(RSCONSULTA.Fields("F4FECent"), "dd/mm/yyyy")
            .FldTipCam.Text = Format(Val("" & RSCONSULTA.Fields("F4tipcam")), "0.000")
            '.FLDDIADEPAGO.Text = "" & RSCONSULTA.Fields("F4DIAPAGO")
            .fldsolicitante.Text = ObtenerCampo("ef2users", "f2nomuser", "f2coduser", RSCONSULTA.Fields("F4CODSOL"), "T", cnn_dbbancos)
'            .F4IGV.Text = Format("" & RSCONSULTA.Fields("F4IGV"), "###,###,##0.00")
            .FldTipDoc.Text = UCase("" & ObtenerCampo("DOCUMENTOS", "F2DESDOC", "F2CODDOC", RSCONSULTA!f4tipdoc & "", "T", cnn_dbbancos))
            If RSCONSULTA!f4tipdoc & "" = "02" Then
               .LblImp.Caption = "Reten."
            Else
                .LblImp.Caption = "I.G.V."
            End If
'            .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "###,###,###,##0.00")
'            '.Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            .FldObservaAll.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
'            .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "###,###,###,##0.00")
            '.F4PLAZO.Text = rsconsulta.Fields("F4PLAZO_ENTREGA") & ""
            Rem NSE .F3FECEN.Text = DateAdd("d", Val(rsconsulta.Fields("F4PLAZO_ENTREGA") & ""), .F4FECEMI.Text)
            '.F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
            '.F4NOTA.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
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
            'If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
            'RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
'            sql = "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'"
'            Set RsCTR_COM = Af.OpenSQLForwardOnly(sql, cconex_ctrcom)
'            If Not RsCTR_COM.EOF Then
'                .LblFirma1.Caption = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
'                .LblCargo1.Caption = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
'                .LblFirma2.Caption = "" & RsCTR_COM.Fields("F1EMITIDO_OCI")
'                .LblCargo2.Caption = "" & RsCTR_COM.Fields("F1OBSGEN_OCi")
'                .LblFirma3.Caption = ""
'                .LblCargo3.Caption = ""
'                If Trim(PnlNomCosto.Caption) <> "BRITANIA" Then
'                .LblFirma3.Caption = "---------------------------------"
'                .LblCargo3.Caption = "VoBo Gerencia Logística"
'                End If
'                .Refresh
'            End If
'            RsCTR_COM.Close
'            If Len(Trim(.LblFirma1.Caption)) > 0 And Len(Trim(.LblFirma2.Caption)) > 0 And Len(Trim(.LblFirma3.Caption)) > 0 Then
'                .LblFirma1.Visible = True
'                .LblFirma2.Visible = True
'                .LblFirma3.Visible = True
'                .LblCargo3.Visible = True
'                .LblFirma1.Width = nAnchoHoja / 3
'                .LblCargo1.Width = nAnchoHoja / 3
'                .LblFirma1.Left = 0
'                .LblCargo1.Left = 0
'                .LblFirma2.Width = nAnchoHoja / 3
'                .LblCargo2.Width = nAnchoHoja / 3
'                .LblFirma2.Left = .LblFirma1.Width
'                .LblCargo2.Left = .LblCargo1.Width
'                .LblFirma3.Width = nAnchoHoja / 3
'                .LblCargo3.Width = nAnchoHoja / 3
'                .LblFirma3.Left = (.LblFirma1.Width) * 2
'                .LblCargo3.Left = (.LblCargo1.Width) * 2
'
'            ElseIf Len(Trim(.LblFirma1.Caption)) > 0 And Len(Trim(.LblFirma2.Caption)) > 0 Then
'                .LblFirma1.Visible = True
'                .LblFirma2.Visible = True
'                .LblFirma1.Width = nAnchoHoja / 2
'                .LblCargo1.Width = nAnchoHoja / 2
'                .LblFirma1.Left = 0
'                .LblCargo1.Left = 0
'                .LblFirma2.Width = nAnchoHoja / 2
'                .LblCargo2.Width = nAnchoHoja / 2
'                .LblFirma2.Left = .LblFirma1.Width
'                .LblCargo2.Left = .LblCargo1.Width
'            ElseIf Len(Trim(.LblFirma1.Caption)) > 0 And Len(Trim(.LblFirma2.Caption)) = 0 Then
'                .LblFirma1.Visible = True
'                .LblFirma2.Visible = False
'                .LblFirma1.Width = nAnchoHoja
'                .LblCargo1.Width = nAnchoHoja
'                .LblFirma1.Left = 0
'                .LblCargo1.Left = 0
'
'            End If
            
            .F4COTIZACION.Text = "" & RSCONSULTA.Fields("F4NUMCOTIZA")
            .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
            '.F4EMITIR.Text = "" & wnomcia
            '.f4emitir2.Text = "" & wdireccion
'            .f4emitir3.Text = "Ph: " & wtelefono & "  Fax: " & wfax
            
            If Rs.State = 1 Then Rs.Close
            
            Rs.Open "Select F2NOMUSER From EF2USERS where F2CODUSER = '" & wusuario & "'", cnn_dbbancos, _
            adOpenKeyset, adLockOptimistic
            
            '.LBLFIRMA.Caption = rs(0)  ' Trim("" & pnlnomsoli.Caption)
            '.lblcargo.Caption = traerCampo("EF2USERS", "F2CARGO", "F2CODUSER", wusuario & "")
            '.lblempresa.Caption = wnomcia
        End If
        .DataControl1.ConnectionString = cnn_dbbancos
        
        CadSql = "SELECT IF3ORDEN.UNIDAD AS F3MEDIDA, CENTROS.F3CODCLI, CENTROS.PO, IF3ORDEN.* "
        CadSql = CadSql & "FROM (IF3ORDEN LEFT JOIN CENTROS ON IF3ORDEN.F3CENCOS = CENTROS.F3COSTO) "
        CadSql = CadSql & "LEFT JOIN EF7MEDIDAS ON IF3ORDEN.UNIDAD = EF7MEDIDAS.F7CODMED "
        CadSql = CadSql & "WHERE (((IF3ORDEN.F4NUMORD)='" & GOC & "') AND ((IF3ORDEN.F4LOCAL)='" & TOC & "')) "
        CadSql = CadSql & "order by val(IF3ORDEN.item)"

        .DataControl1.Source = CadSql
        '.F3FECEN.Text = dxDBGrid1.Columns.ColumnByFieldName("F3ENTREGA").Value
        'AGRUPACION POR CLIENTE
        '********************************************
'        sql = "SELECT CE.F3CODCLI,D.ITEM, D.F5NOMPRO, D.F3CANPRO, D.F3PRECOS, D.F3TOTAL FROM IF3ORDEN AS D, CENTROS AS CE"
'        sql = sql & " " & "Where D.F4local = '" & TOC & "'  AND CE.F3ABREV=D.F3CENCOS And D.F4NUMORD = '" & GOC & "' ORDER BY CE.F3CODCLI"
'
'
'       If Rs.State = 1 Then Rs.Close
'        Rs.Open sql, cnn_dbbancos, 3, 1
'
'        Set .DataControl1.Recordset = Rs
'        .GroupHeader1.DataField = "f3codcli"
'        .DataControl1.ConnectionString = cnn_dbbancos
'        .DataControl1.Source = sql
        
        .Caption = "ORDEN DE COMPRA NACIONAL"
        RSCONSULTA.Close
        .Show 1
        
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
    " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = " & GOC & _
    " AND A.F4ESTNUL<>'S'ORDER BY A.F4NUMORD DESC;"

    RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RSCONSULTA.EOF Then
        '.F4NUMORD.Text = Format("" & rsconsulta.Fields("F4NUMORD"), "0000000")
        .LblTitle.Caption = "ORDEN DE COMPRA Nro. " & RSCONSULTA.Fields("F4NUMORD")
        .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").Value
        .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
        .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
        .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
        .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
        .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
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
    .DataControl1.Source = "select * from tmpOrdendeCompra"

    .Show vbModal
End With
End If
Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Logistica"
    Resume Next
End Sub
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
            .lblmoneda4.Caption = "S/."
            .lblmoneda2.Caption = "S/."
            .lblmoneda1.Caption = "S/."
            .lblmoneda3.Caption = "S/."
        Else
            .lblmoneda4.Caption = "US$"
            .lblmoneda2.Caption = "US$"
            .lblmoneda1.Caption = "US$"
            .lblmoneda3.Caption = "US$"
        End If
        .flddirec1.Text = wf1direc1
        .flddirec2.Text = wf1direc2
        .fldruc.Text = wrucempresa
        .fldempresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
        
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
            .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
            .F4IGV.Text = Format("" & RSCONSULTA.Fields("F4IGV"), "0.00")
            .F4RND.Text = Format("" & RSCONSULTA.Fields("F4RND"), "0.00")
            .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "0.00")
            .Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            '.F4OBSERVA.Text = "" & rsconsulta.Fields("F4OBSERVA")
            .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "0.00")
            '.F4PLAZO.Text = rsconsulta.Fields("F4PLAZO_ENTREGA") & ""
            Rem NSE .F3FECEN.Text = DateAdd("d", Val(rsconsulta.Fields("F4PLAZO_ENTREGA") & ""), .F4FECEMI.Text)
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
        .DataControl1.Source = "SELECT * FROM tmpOrdendeCompra"
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
        .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").Value
        .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
        .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
        .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
        .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
        .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
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
    .DataControl1.Source = "select * from tmpOrdendeCompra"

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
            .lblmoneda4.Caption = "S/."
            .lblmoneda2.Caption = "S/."
            .lblmoneda1.Caption = "S/."
            .lblmoneda3.Caption = "S/."
        Else
            .lblmoneda4.Caption = "US$"
            .lblmoneda2.Caption = "US$"
            .lblmoneda1.Caption = "US$"
            .lblmoneda3.Caption = "US$"
        End If
        .flddirec1.Text = wf1direc1
        .flddirec2.Text = wf1direc2
        .fldruc.Text = wrucempresa
        .fldempresa.Text = wnomcia '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
        
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
            .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
            .F4IGV.Text = Format("" & RSCONSULTA.Fields("F4IGV"), "0.00")
            .F4RND.Text = Format("" & RSCONSULTA.Fields("F4RND"), "0.00")
            .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "0.00")
            .Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            '.F4OBSERVA.Text = "" & rsconsulta.Fields("F4OBSERVA")
            .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "0.00")
            '.F4PLAZO.Text = rsconsulta.Fields("F4PLAZO_ENTREGA") & ""
            Rem NSE .F3FECEN.Text = DateAdd("d", Val(rsconsulta.Fields("F4PLAZO_ENTREGA") & ""), .F4FECEMI.Text)
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
        .DataControl1.Source = "SELECT * FROM tmpOrdendeCompra"
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
        .F2NOMPROV.Text = "" & RSCONSULTA.Fields("f2nomprov").Value
        .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
        .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
        .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
        .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
        .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
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
    .DataControl1.Source = "select * from tmpOrdendeCompra"

    .Show vbModal
End With
End If
End Sub
Private Sub EMAIL()

End Sub

Private Sub Calcula_PvtaTot()
Dim cantidad    As Double
Dim totdcto     As Double
Dim ValVta      As Double
Dim IGV         As Double
Dim preciounit  As Double
Dim TOTAL       As Double
Dim costo       As Double

    With Grid
        cantidad = Val(Format(.Columns.ColumnByFieldName("F3CANPRO").Value, "0.00"))
        If cantidad > 0 Then
            'If cmbtipopera.ListIndex = 0 Then
                If .Columns.ColumnByFieldName("F5AFECTO").Value = "*" Then     'Afecto
                    totdcto = (Val(Format("" & .Columns.ColumnByFieldName("F3canpro").Value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PREcos").Value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").Value, "0.00"))) / 100
                    .Columns.ColumnByFieldName("F3TOTDCT").Value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) - totdcto, "0.00"))
                    .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                    IGV = ValVta * (wwigv / 100)
                    .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "#,##0.00")
                    preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) + (Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) * (wwigv / 100))
                    .Columns.ColumnByFieldName("F3PREUNI").Value = Format$(preciounit, "###,##0.00")
                    TOTAL = ValVta + IGV
                    .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
                Else  'Inafecto
                    IGV = 0
                    .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "0.00")
                    totdcto = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").Value, "0.00")) / 100
                    .Columns.ColumnByFieldName("F3TOTDCT").Value = Format$(totdcto, "####,##0.00")
                    ValVta = Val(Format(cantidad * Val(Format(.Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) - totdcto, "0.00"))
                    .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                    preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000"))
                    .Columns.ColumnByFieldName("F3PREUNI").Value = Format$(preciounit, "###,##0.00")
                    TOTAL = ValVta + IGV
                    .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
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
Dim cantidad    As Double
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
        If .Columns.ColumnByFieldName("F3PREUNI").Value = 0 Then
            .Dataset.Edit
            cantidad = Val(Format(.Columns.ColumnByFieldName("F3CANPRO").Value, "0.00"))
            If cantidad > 0 Then
                'If cmbtipopera.ListIndex = 0 Then
                    If .Columns.ColumnByFieldName("F5AFECTO").Value = True Then     'Afecto
                        totdcto = (Val(Format("" & .Columns.ColumnByFieldName("F3canpro").Value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PREcos").Value, "0.00")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").Value, "0.00"))) / 100
                        .Columns.ColumnByFieldName("F3TOTDCT").Value = Format$(totdcto, "####,##0.00")
                        ValVta = Val(Format(cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3sinigv").Value, "0.0000")) - totdcto, "0.00"))
                        .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                        IGV = ValVta * (wwigv / 100)
                        .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "#,##0.00")
                        preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3sinigv").Value, "0.0000")) + (Val(Format("" & .Columns.ColumnByFieldName("F3sinigv").Value, "0.0000")) * (wwigv / 100))
                        .Columns.ColumnByFieldName("F3PREUNI").Value = Format$(preciounit, "###,##0.00")
                        TOTAL = ValVta + IGV
                        .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
                    Else  'Inafecto
                        IGV = 0
                        .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "0.00")
                        totdcto = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) * Val(Format("" & .Columns.ColumnByFieldName("F3PORDCT").Value, "0.00")) / 100
                        .Columns.ColumnByFieldName("F3TOTDCT").Value = Format$(totdcto, "####,##0.00")
                        ValVta = Val(Format(cantidad * Val(Format(.Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000")) - totdcto, "0.00"))
                        .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                        preciounit = Val(Format("" & .Columns.ColumnByFieldName("F3PRECOS").Value, "0.0000"))
                        .Columns.ColumnByFieldName("F3PREUNI").Value = Format$(preciounit, "###,##0.00")
                        TOTAL = ValVta + IGV
                        .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(TOTAL, "###,##0.00")
                    End If
            End If
            .Dataset.Post
        End If
        .Dataset.Next
        Next K
    End With
End Sub


Sub MostrarDatos()
Dim sw_nuevo_temp   As Boolean
Dim xnombre         As String
Dim i               As Integer
Dim entrega         As Date
Dim J               As Integer
    
    If TOC = "OC" Then
        Me.Caption = "Orden de Compra"
    Else
        Me.Caption = "Orden de Servicio"
    End If
    
    csql = "select * from tb_cabsolicitud where cod_solicitud='" & num_solcomp & "' and cs_documento='" & TOC & "'"
    TOC = dxDBGrid1.Columns.ColumnByFieldName("cs_documento").Value
    Set rssolcab = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    With rssolcab
        If Not .EOF And Not .Bof Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2NEWRUC,F2NOMPROV,F2DIRPROV from EF2PROVEEDORES where F2newruc='" & !cs_proveedor & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                Txt_Prove.Text = "" & rst!F2NEWRUC
                pnlnomprv.Caption = rst!F2NOMPROV
                pnldireprv.Caption = IIf(IsNull(rst!F2DIRPROV), " ", rst!F2DIRPROV)
            Else
                pnlnomprv.Caption = "Ruc es menor a 11 digitos"
                pnldireprv.Caption = "No tiene "
            End If
            rst.Close

            Txt_NumSolComp = !cod_solicitud & ""
            
            xnombre = !cs_codsolicitante
            txtobserva.Text = Trim("" & !cs_observaciones)
            If rstaux.State = adStateOpen Then rstaux.Close
            rstaux.Open "SELECT f2nomuser FROM ef2users WHERE f2coduser='" & Trim(xnombre) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rstaux.EOF Then
                txtcodsoli.Text = xnombre
                pnlnomsoli.Caption = "" & rstaux.Fields("f2nomuser")
            Else
                pnlnomsoli.Caption = ""
            End If
            rstaux.Close
            If !cs_moneda = "S" Then
                Cmbmone.ListIndex = 0
            Else
                Cmbmone.ListIndex = 1
            End If
            TxtCodCosto.Text = "" & !cs_codcosto
            txtcodcosto_KeyPress 13
            txtlugar_entrega.Text = left(Trim("" & !cs_LugEntr), 100)
            'txt_tc.Text = Format(Val(.Fields("F4TIPCAM") & ""), "0.000")
        End If
        rssolcab.Close
    End With
     
    '*** detalle de solicitud de compra
    'Versión Nueva
    With Grid
        
        csql = "SELECT TB_DETSOLICITUD.cod_solicitud, TB_DETSOLICITUD.observa,TB_DETSOLICITUD.ITEM, TB_CABSOLICITUD.cs_fecha, TB_CABSOLICITUD.cs_codsolicitante, "
        csql = csql & "EF2USERS.F2NOMUSER, TB_CABSOLICITUD.cs_observaciones, TB_DETSOLICITUD.F5CODCOSTO, CENTROS.F3DESCRIP, "
        csql = csql & "TB_DETSOLICITUD.cod_producto,TB_DETSOLICITUD.precio, TB_DETSOLICITUD.ds_descripcion, TB_DETSOLICITUD.ds_unidmed, EF7MEDIDAS.F7SIGMED, "
        csql = csql & "TB_DETSOLICITUD.ds_cantidad, IIf(IsNull(ORDENES.CANT_ORDEN),0,ORDENES.CANT_ORDEN) AS TOT_ORDEN, "
        csql = csql & "TB_DETSOLICITUD.ds_cantidad-IIf(IsNull(ORDENES.CANT_ORDEN),0,ORDENES.CANT_ORDEN) AS SALDO, TB_DETSOLICITUD.f5SINigv, TB_DETSOLICITUD.f5CONigv, TB_DETSOLICITUD.ruc_proveedor, TB_DETSOLICITUD.F5AFECTO "
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
        csql = csql & "And ((TB_CABSOLICITUD.cs_estado) >= '2')) "
        csql = csql & "AND TB_DETSOLICITUD.cod_solicitud='" & num_solcomp & "'"
        If item_solcomp > 0 Then
            csql = csql & "and TB_DETSOLICITUD.item=" & item_solcomp & " "
        End If
        csql = csql & "ORDER BY TB_DETSOLICITUD.ITEM"
        Set rsSolDet = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
        If Not (rsSolDet.EOF) Then
            
            sw_nuevo_temp = False
            sw_nuevo_item = False
            rsSolDet.MoveFirst
            J = 0
            Do While Not (rsSolDet.EOF)
                If rsSolDet!cod_solicitud = Trim(num_solcomp) Then
                J = J + 1
                    If J = 1 Then
                        Grid.Dataset.Edit
                    Else
                        Grid.Dataset.Append
                    End If
                    .Dataset.FieldValues("f3canpro") = rsSolDet!Saldo
                    .Dataset.FieldValues("f3redondeo") = rsSolDet!Saldo
                    .Dataset.FieldValues("f3codpro") = rsSolDet!COD_PRODUCTO & ""
                    .Dataset.FieldValues("f5nompro") = rsSolDet!ds_descripcion & ""
                    .Dataset.FieldValues("f3codmedida") = rsSolDet!ds_unidmed & ""
                    .Dataset.FieldValues("f3desmedida") = rsSolDet!ds_unidmed 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rsSolDet!ds_unidmed & "", "T", cnn_dbbancos)
                    .Dataset.FieldValues("f3sinigv") = Val(Format(rsSolDet!f5SINigv, "0.0000"))
                    .Dataset.FieldValues("f5afecto") = rsSolDet!F5AFECTO
                    .Dataset.FieldValues("f3conigv") = Val(Format(rsSolDet!f5CONigv, "0.0000"))
                    .Dataset.FieldValues("f3valdesc") = 0
                    .Dataset.FieldValues("f3pordesc") = 0
                    .Dataset.FieldValues("f3total") = 0
                    .Dataset.FieldValues("cod_solicitud") = num_solcomp
                    .Dataset.FieldValues("f3observa") = rsSolDet!OBSERVA & ""
                    .Dataset.FieldValues("f3gasto") = ""
                    .Dataset.FieldValues("f5codcta") = ""
'                    entrega = IIf(IsNull(rsSolDet!cs_fentrega), Format$(Date, "dd/mm/yyyy"), Format$(rsSolDet!cs_fentrega, "dd/mm/yyyy"))
                    '.Dataset.FieldValues("f3fentrega") = entrega
                    '.Dataset.FieldValues("check") = True
                    '.Dataset.FieldValues("cant_ant") = 0#
                    '.Dataset.FieldValues("f5codfab") = rsSolDet!f5codfab & ""
                    .Dataset.Post
                End If
            rsSolDet.MoveNext
            Loop
            
            '.Dataset.EnableControls
            '.Dataset.Open
            '.OptionEnabled = True
            sw_nuevo_item = False
        End If
        rsSolDet.Close
        
    End With
    
End Sub

Private Sub abofechaentrega_Click()
    abofechaentrega.Value = Now
End Sub

Private Sub abofechaentrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
        dxDBGrid1.Columns.FocusedIndex = 1
    End If

End Sub

Private Sub atbmenu_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim resp    As Integer
    On Error Resume Next
    Select Case Tool.Id
        Case "ID_Nuevo":
            inicio = True
            Me.MousePointer = vbHourglass
            Wnuevo = True
            If swGrabacion = True Then
                resp = MsgBox("La Orden no ha sido grabada. ¿Desea grabarla ahora?", vbYesNo + vbQuestion, "Sistema de Logística")
                If resp = vbYes Then
                    MODIFICAR_OC
                End If
            End If
            '
            sw_nuevo_documento = False
            Limpia_Orden
            Limpiar
            'AdicionaItem
            AdicionaItemGrid
            sw_nuevo_documento = True
            SendKeys "{tab}"
            loc = 1
            
            Me.MousePointer = vbDefault
            'atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = False
            atbmenu.Tools.ITEM("ID_Anular").Visible = False
            atbmenu.Tools.ITEM("ID_Eliminar").Visible = False
            atbmenu.Tools.ITEM("ID_Imprimir").Visible = False
'            atbmenu.Tools.ITEM("ID_Aprobacion").Visible = False
        Case "ID_Grabar":
            If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
                Grid.Dataset.Post
                sw_detalle = True
            End If

            If MsgBox("¿Desea Grabar la " & Me.Caption & "?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                Me.MousePointer = vbHourglass
                GrabarOC
                'ActualizarNumOrd
                Me.MousePointer = vbDefault
            End If
    
        Case "ID_Imprimir":
           Me.MousePointer = vbHourglass
            If Len(Trim(Txt_NumOC.Text)) > 0 Then
                Select Case wrucempresa
                Case "20381208835"
                    Imprime_Orden_Electrica 1
                Case "20434047171"
                    Imprime_Orden_Ditec 1
                Case Else
                    Imprime_Orden 1
                End Select
            Else
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
            End If
            Me.MousePointer = vbDefault
        Case "ID_RenovarOrden"
            'MsgBox "renovar"
            Me.MousePointer = 11
            If Rs.State = 1 Then Rs.Close
            sql = "select * from if4orden where f4numord='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            Rs.Open sql, cnn_dbbancos, 3, 1
            If Rs.RecordCount > 0 Then
                If Rs!f4estnul = "S" Then
                    MsgBox "No se puede renovar una Orden de Compra Anulada", vbExclamation, "Sistema de Logística"
                    SwRenovar = False
                    Exit Sub
                End If
            Else
                MsgBox "Error, no existe la orden que quiere renovar", vbCritical, wnomcia
            End If
            SwRenovar = True
            If MsgBox("¿Desea Renovar la Orden de Compra?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                Me.MousePointer = vbHourglass
                GrabarOC
                'ActualizarNumOrd
                'ANULANDO ANTERIOR
                If Trim$(Txt_NumOC.Text) = "" Then
                    MsgBox "No existe Orden de Compra", vbInformation, "wnomcia"
                    Exit Sub
                Else
                    eliminar_sin_preguntar
                End If
                SwRenovar = False
                Txt_TOC.Text = TOC
                Me.Txt_NumOC.Text = wNumOc
                Call Txt_NumOC_KeyPress(13)
                
            End If
            Me.MousePointer = 1
            
        Case "ID_Anular":
            If Trim$(Txt_NumOC.Text) = "" Then
                MsgBox "No existe Orden de Compra", vbInformation, "Sistema de Logística"
                Exit Sub
            Else
                eliminar
            End If
        Case "ID_Eliminar"
            Dim strReq As String
            If MsgBox("¿Desea Eliminar la Orden de Compra?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                eliminar_sin_preguntar
                
                strReq = ""
                For i = 1 To Grid.Dataset.RecordCount
                    Grid.Dataset.RecNo = i
                    If strReq <> Grid.Columns.ColumnByFieldName("cod_solicitud").Value & "" Then
                        strReq = Grid.Columns.ColumnByFieldName("cod_solicitud").Value & ""
                        VerificaAtencionDeRequerimiento (strReq)
                    End If
                    
                Next
                csql = "delete * from if4orden where f4local='" & TOC & "' and f4numord='" & Txt_NumOC.Text & "'"
                cnn_dbbancos.Execute csql
                AlmacenaQuery_sql csql, cnn_dbbancos
                Actualiza_Log csql, cnn_dbbancos.ConnectionString
                Me.Hide
                lista_oc.dxDBGrid1.Dataset.ADODataset.Requery
            End If
        Case "ID_Email":
            EMAIL
            
        Case "ID_CtasxPagar"
            If Len(Trim(Txt_NumOC.Text)) > 0 Then
                If rsif4orden.State = adStateOpen Then rsif4orden.Close
                rsif4orden.Open "SELECT F4CORRELA FROM IF4ORDEN WHERE F4NUMORD=" & Txt_NumOC.Text & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsif4orden.EOF Then
                    If Val("" & rsif4orden.Fields("F4CORRELA")) > 0 Then
                        MsgBox "La orden de compra ya fue trasladada a cuentas por pagar.", vbInformation, "Atención"
                    Else
                        If MsgBox("Está seguro(a) de trasladar la Orden de Compra a Cuentas por Pagar ?", vbYesNo, "Atención") = vbYes Then
                            TRASLADA_CTASXPAGAR Txt_NumOC.Text
                        End If
                    End If
                End If
                rsif4orden.Close
            Else
                MsgBox "La Orden de Compra no ha sido grabada.", vbInformation, "Atención"
            End If
        Case "ID_Aprobacion"
            If MsgBox("¿Desea solicitar la aprobación de la Orden de Compra " & Txt_NumOC.Text & "?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
                Call EnviaMail(Txt_NumOC.Text)
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
            If Trim(prSol!numorden) <> "" And Trim(prSol!numorden) <> Trim(ordendecompra.Txt_NumOC.Text) Then
                cadena = "" & prSol!numorden & " , " & Trim(ordendecompra.Txt_NumOC.Text)
            Else
                cadena = Trim(ordendecompra.Txt_NumOC.Text)
            End If
            sql = "Update TB_CABSOLICITUD set NumOrden='" & left(Trim(cadena), 255) & "' where Cod_Solicitud='" & Trim(psolicitud) & "'"
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
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
    Label11.Visible = True
    lblmoneda(0).Visible = True
    lblmoneda(1).Visible = True
    lblmoneda(2).Visible = True
    txtmonto.Visible = True
    txtbase.Visible = True
    TxtIgv.Visible = True
    
End Sub

Sub Invisi()

    Cmbmone.ListIndex = 1
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.left = 5000
    lblmoneda(0).Visible = False
    lblmoneda(1).Visible = False
    lblmoneda(2).left = 5600
    txtmonto.Visible = False
    txtbase.Visible = False
    TxtIgv.Visible = False
    txttotal.left = 4905
    
End Sub

Sub Forma_Imp()

    Invisi
    
End Sub

Private Sub cmdcerrar_Click()
'pnlcosto.Visible = False
End Sub

Private Sub cmbTipDoc_Click()
With Grid.Columns
    If right(CmbTipDoc.Text, 2) = "02" Then
        .ColumnByFieldName("F3SINIGV").Caption = "Precio s/Ret"
        .ColumnByFieldName("F3CONIGV").Caption = "Precio c/Ret"
        .ColumnByFieldName("F3IGV").Caption = "Retención"
    Else
        .ColumnByFieldName("F3SINIGV").Caption = "Precio s/IGV"
        .ColumnByFieldName("F3CONIGV").Caption = "Precio c/IGV"
        .ColumnByFieldName("F3IGV").Caption = "I.G.V."
    End If
End With

 Dim nPorc As Double
            If right(CmbTipDoc.Text, 2) = "02" Then
                nPorc = gretenc
                If Grid.Columns.ColumnByFieldName("f3sinigv").Value >= 700 Then
                    Grid.Dataset.Edit
                    Grid.Columns.ColumnByFieldName("F5afecto").Value = True
                    Grid.Dataset.Post
                Else
                    Grid.Dataset.Edit
                    Grid.Columns.ColumnByFieldName("F5afecto").Value = False
                    Grid.Dataset.Post
                End If
                    
            Else
                nPorc = wwigv
            End If
            Grid.Dataset.Edit
            'Grid.Columns.ColumnByFieldName("F3COLMOD").Value = UCase(Grid.Columns.FocusedColumn.FieldName)
            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
                'Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value * (1 + (nPorc / 100))
                If right(CmbTipDoc.Text, 2) = "02" Then
                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
                Else
                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value * (1# + (nPorc / 100))
                End If
                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
                
            Else
                Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
                
                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
            End If
            Grid.Dataset.Post
            CalculaTotal
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
            lblmoneda(0).Caption = "S/."
            lblmoneda(1).Caption = "S/."
            lblmoneda(2).Caption = "S/."
            lblmoneda(3).Caption = "S/."
            lblmoneda(4).Caption = "S/."
            Me.txttotal.BackColor = &HC0FFFF
            Me.TxtIgv.BackColor = &HC0FFFF
            Me.txtbase.BackColor = &HC0FFFF
            Me.txtmonto.BackColor = &HC0FFFF
            Me.TxtRnd.BackColor = &HC0FFFF
        Case 1:
            lblmoneda(0).Caption = "US$"
            lblmoneda(1).Caption = "US$"
            lblmoneda(2).Caption = "US$"
            lblmoneda(3).Caption = "US$"
            lblmoneda(4).Caption = "US$"
            Me.txttotal.BackColor = &HC0FFC0
            Me.TxtIgv.BackColor = &HC0FFC0
            Me.txtbase.BackColor = &HC0FFC0
            Me.txtmonto.BackColor = &HC0FFC0
            Me.TxtRnd.BackColor = &HC0FFC0
    End Select
    If Not inicio Then swGrabacion = True
    Grid.Dataset.Refresh
End Sub

Private Sub Cmbmone_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub calcula()
On Error GoTo HNDERR
Dim afecto      As Double
Dim inafecto    As Double
Dim IGV         As Double
Dim sql         As String

    'If cmbtipopera.ListIndex = 0 Then
        sql = "select sum(iif(f5afecto='*',f5valvta)) as afecto, " _
        & "sum(iif(isnull(f5afecto),f5valvta)) as inafecto, sum(f3igv) as igv from tmpOrdendeCompra"
        If rst.State = adStateOpen Then rst.Close
        
        If rst.State = adStateOpen Then rst.Close
        rst.Open sql, cnn_form, adOpenStatic, adLockOptimistic
        If Not (rst.EOF) Then
            afecto = IIf(IsNull(rst.Fields("afecto")), 0, rst.Fields("afecto"))
            inafecto = IIf(IsNull(rst.Fields("inafecto")), 0, rst.Fields("inafecto"))
            IGV = IIf(IsNull(rst.Fields("igv")), 0, rst.Fields("igv"))
            
            txtbase.Text = Format$(afecto, "####,##0.00")
            txtmonto.Text = Format$(inafecto, "####,##0.00")
            TxtIgv.Text = Format(IGV, "###,###,##0.00")
            
            txttotal.Text = Format$(afecto + inafecto + IGV + (Me.TxtRnd.Text), "###,##0.00")
        End If
        rst.Close
    'End If
    
    Exit Sub
    
HNDERR:
    Select Case Err.Number
        Case -2147217865
            Resume Next
    End Select
    
End Sub

Private Sub dxCheckBox1_Click()
    If dxCheckBox1.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f3gasto").Visible = True
        Grid.Columns.ColumnByFieldName("f5codcta").Visible = True
    Else
        Grid.Columns.ColumnByFieldName("f3gasto").Visible = False
        Grid.Columns.ColumnByFieldName("f5codcta").Visible = False
    End If
End Sub

Private Sub dxCheckBox2_Click()
    If dxCheckBox2.Checked = 1 Then
        Grid.Columns.ColumnByFieldName("f3sinigv").DecimalPlaces = 6
    Else
        Grid.Columns.ColumnByFieldName("f3sinigv").DecimalPlaces = 2
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
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
              '  dxDBGrid1.Columns.ColumnByFieldName("f5c").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").Value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value = wcodfab
                
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = wmarca
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = wafecto

                Rem EMB dxDBGrid1.Dataset.FieldValues("f5valvta") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.Value, "dd/mm/yyyy"))
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

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Me.Height = 8085
    'Me.Width = 12015
    Me.left = Menu.dxSideBar1.Width
    Me.top = 1200
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strReq As String
    strReq = ""
    For i = 1 To Grid.Dataset.RecordCount
        Grid.Dataset.RecNo = i
        If strReq <> Grid.Columns.ColumnByFieldName("cod_solicitud").Value & "" Then
            strReq = Grid.Columns.ColumnByFieldName("cod_solicitud").Value & ""
            '******************* NUEVO CAMBIO ****************
            'VerificaAtencionDeRequerimiento (strReq)
            '*************************************************
        End If
        
    Next


    sw_nuevo_item = True
    dxDBGrid1.Dataset.Close
    lista_oc.dxDBGrid1.Dataset.Active = False
    lista_oc.dxDBGrid1.Dataset.Refresh
    lista_oc.dxDBGrid1.Dataset.Active = True

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
Dim fec     As Date
    SwRenovar = False
    Me.MousePointer = 11
    num_solcomp = ""
    If wTipoOC = 1 Or TOC = "OC" Then
        Me.Caption = "Orden de Compra"
    Else
        Me.Caption = "Orden de Servicio"
    End If

    
    Set rst = New ADODB.Recordset
    Set rsOrdenCab = New ADODB.Recordset
    Set rsOrdenDet = New ADODB.Recordset
    Set rsproductos = New ADODB.Recordset
    Set rssolcab = New ADODB.Recordset
    Set rsSolDet = New ADODB.Recordset
    Set rstaux = New ADODB.Recordset
    
    sw_ayuda = False
    inicio = True
    swGrabacion = False
    sw_activate = False
    CargaDocumentos CmbTipDoc
'    loc = 1
    If loc = 2 Then
        Call define_cabecera
        txtmonto.Visible = False
        TxtIgv.Visible = False
        txttotal.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        lblmoneda(0).Visible = False
        lblmoneda(1).Visible = False
        lblmoneda(2).Visible = False
    Else
        loc = 1
    End If
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    fec = txt_fecha.Value
    Wnuevo = True
    flawigv = False
    SWcondipago = 0
    
    Set rst = Af.OpenSQLForwardOnly("select F1IGV, F1RETENC from param_com where f1codemp='" & UCase(wempresa) & "'", cconex_ctrcom)
    If Not (rst.EOF) Then
         wwigv = rst.Fields("F1IGV")
         gretenc = rst.Fields("F1RETENC")
    End If
    rst.Close
    
    Txt_Prove.Enabled = True
    If FlagGeneraOC = False Then
        Wnuevo = True
    End If
     
    jc = 0
    
    sw_nuevo_item = False
    If Dir(wrutatemp & "tmp_logistica_" & wempresa & ".mdb") <> "" Then
        cnombase = "tmp_logistica_" & wempresa & ".mdb"
    Else
        cnombase = "tmp_bancos.MDB"
    End If
    cnomtabla = "tmpOrdendeCompra"
    
'    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
'    If cnn_form.State = adStateOpen Then cnn_form.Close
'    cnn_form.Open cconex_form
    
    StrCn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ";Persist Security Info=False"
    If CnTmp.State = 1 Then CnTmp.Close
    CnTmp.Open StrCn
    
 '   Call CONFIGURA_GRID
    Call CONFIGURA_GRID_OC
        
    
    If sw_nuevo_documento = True Then
'        DELETEREC_LOG cnomtabla, cnn_form
        'DELETEREC_LOG cnomtabla, CnTmp
        Limpiar
'        AdicionaItem
        sw_nuevo_documento = False
        AdicionaItemGrid
        sw_nuevo_documento = True
    Else
        inicio = True
        MODIFICAR_OC
        sw_nuevo_documento = False
        inicio = False
'        atbmenu.Tools.ITEM("ID_Grabar").Visible = True
'        atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
'        atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
'        atbmenu.Tools.ITEM("ID_Anular").Visible = True
    End If
        atbmenu.Tools.ITEM("ID_Grabar").Visible = True
        'atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
        'atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
'        atbmenu.Tools.ITEM("ID_Anular").Visible = True
    
    
    Me.MousePointer = 1
   
End Sub

Private Sub CargaDocumentos(pCombo As ComboBox)
Dim TbDocumento1 As New ADODB.Recordset
    SqlCad = "Select * from documentos"
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


Sub Limpiar()

    SWcondipago = 0
    Txt_NumOC = ""
    Txt_NumSolComp = ""
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    abofechaentrega.CheckBox = True
    abofechaentrega.Value = Empty 'Format("", "dd/MM/yyyy")
    aBoHoraEntrega.Value = Time
    FrameOC.Caption = ""
    txtcontacto.Text = ""
    txtcodsoli = wusuario
    Cmbmone.ListIndex = 0
    TxtCodCosto.Text = ""
    PnlNomCosto.Caption = ""
    txtcodforma = ""
    pnlnomforma = ""
           
    txt_tc.Text = Format(traerCampo("CAMBIOS", "CAMBIO", "FECHA", Me.txt_fecha.Value, ""), "0.000")

    
    Txt_Referencia = ""
    
    txtbase = "0.00"
    txtmonto = "0.00"
    TxtIgv.Text = "0.00"
    txttotal = "0.00"
       
    
       
    SWcondipago = 0
    txtempresa.Text = UCase$(wnomcia)
    
    txtCotizacion.Text = ""
    txtlugar_entrega.Text = ""
    
    
    
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
    
    txt_tc.Text = Format(traerCampo("CAMBIOS", "CAMBIO", "FECHA", Me.txt_fecha.Value, ""), "0.000")
    txttotal.Text = "0.00"
    TxtIgv.Text = "0.00"
    txtbase.Text = "0.00"
    txtmonto.Text = "0.00"
    
    wgraba = 1
    
    
End Sub

Sub Visi()

    txtbase.Visible = True
    TxtIgv.Visible = True
    txttotal.Visible = True

End Sub

Sub LLENA_TEMPCAB()
Dim cnn         As ADODB.Connection
Dim tempocompra As ADODB.Recordset
Dim X           As Integer
Dim rsprod      As New ADODB.Recordset

    'Nueva Versión
    Set cnn = New ADODB.Connection
    Set tempocompra = New ADODB.Recordset
    If Dir(wrutatemp & "tmp_logistica_" & wempresa & ".mdb") <> "" Then
        cnombase = "tmp_logistica_" & wempresa & ".mdb"
    Else
        cnombase = "tmp_bancos.MDB"
    End If
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ";Persist Security Info=False"
    
    sql = "delete * from tmpocompra"
    cnn.Execute sql
    'AlmacenaQuery_sql sql, cnn
    
    If tempocompra.State = adStateOpen Then tempocompra.Close
    tempocompra.Open "tmpocompra", cnn, adOpenStatic, adLockOptimistic
    
    With dxDBGrid1
        If .Dataset.RecordCount = 0 Then
            tempocompra.Close
            cnn.Close
            Exit Sub
        End If
        .Dataset.First
        If Not (.Dataset.EOF) Then
            .Dataset.First
            Do While Not (.Dataset.EOF)
                If Val(IIf(IsNull(.Dataset.FieldValues("f3precos")), 0, _
                .Dataset.FieldValues("f3precos"))) > 0 Then
                    tempocompra.AddNew
                    tempocompra!Orden = Format(Txt_NumOC.Text, "0000000")
                    tempocompra!PROVEEDOR = pnlnomprv.Caption
                    tempocompra!direccion = pnldireprv.Caption
                    tempocompra!ruc = Txt_Prove.Text
                    tempocompra!CLIENTE = txtFechaPago.Text
                    tempocompra!CODCONTA = IIf(ChK_regularizacion.Checked = True, 1, 0)
                    'tempocompra!CONTACTO = txtcontacto.Text
                    tempocompra!fecha = txt_fecha.Value
                    tempocompra!FORPAG = pnlnomforma.Caption
                    tempocompra!Moneda = Cmbmone.Text
                    tempocompra!referencia = Txt_Referencia.Text
                    'tempocompra!Centro = txtcodcosto.Text
                    'tempocompra!nomcentro = pnlnomcosto.Caption
                    tempocompra!OBSERVA = txtobserva.Text
                    tempocompra!SUBTOTAL = txtbase.Text
                    tempocompra!MONTOINA = txtmonto.Text
                    tempocompra!IGV = TxtIgv.Text
                    tempocompra!TOTAL = txttotal.Text
                    tempocompra!empresa = txtempresa.Text
                    tempocompra!ss = Txt_NumSolComp.Text
                    tempocompra!Codigo = "" & .Dataset.FieldValues("f3codpro")
                    tempocompra!Descripcion = "" & .Dataset.FieldValues("f5nompro")
                    tempocompra!cantidad = .Dataset.FieldValues("f3canpro")
                    tempocompra!costo = .Dataset.FieldValues("f3precos")
                    tempocompra!descuento = .Dataset.FieldValues("f3pordct")
                    tempocompra!Precio = .Dataset.FieldValues("f3preuni")
                    
                    
                    If rsprod.State = adStateOpen Then rsprod.Close
                    rsprod.Open "SELECT F7CODMED from if5pla where f5codpro='" & .Dataset.FieldValues("f3codpro") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not (rsprod.EOF) Then
                        tempocompra!unidad = rsprod.Fields("F7CODMED") & ""
                    End If
                    rsprod.Close
                        
                    tempocompra.Update
                End If
                .Dataset.Next
            Loop
            .Dataset.First
        End If
        tempocompra.Close
        cnn.Close
    End With
    
End Sub



Private Sub Grid_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
   If sw_nuevo_item = False Then
        If Action = daInsert Then
            Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("ITEM").Value = Grid.Dataset.RecordCount + 1
            Grid.Columns.FocusedIndex = 1
        End If
        
    End If
    If Action = daEdit Then
    If Grid.Columns.ColumnByFieldName("f3redondeo").Value > 0 And Grid.Columns.ColumnByFieldName("f3canpro").Value > Grid.Columns.ColumnByFieldName("f3redondeo").Value Then
        MsgBox "No puede poner una cantidad mayor a la del requerimiento", vbCritical
        Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("f3canpro").Value = Grid.Columns.ColumnByFieldName("f3redondeo").Value
        Grid.Dataset.Post
        Grid.Dataset.Edit
    End If
End If
End Sub

Private Sub Grid_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
Dim nPorc As Double
If right(CmbTipDoc.Text, 2) = "02" Then
    nPorc = gretenc
Else
    nPorc = wwigv
End If
Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
Case "F5AFECTO"
    Grid.Dataset.Edit
    If Grid.Columns.ColumnByFieldName("F5AFECTO").Value = True Then
        Grid.Columns.ColumnByFieldName("F5AFECTO").Value = False
        'If UCase(Grid.Columns.ColumnByFieldName("F3COLMOD").Value & "") = "F3SINIGV" Then
            Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
        'Else
        '    Grid.Columns.ColumnByFieldName("F3sinIGV").Value = Grid.Columns.ColumnByFieldName("F3conIGV").Value
        'End If
        Grid.Columns.ColumnByFieldName("F3baseimp").Value = 0
        Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3SINIGV").Value * Grid.Columns.ColumnByFieldName("F3canpro").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = 0
    Else
        Grid.Columns.ColumnByFieldName("F5AFECTO").Value = True
        'If UCase(Grid.Columns.ColumnByFieldName("F3COLMOD").Value & "") = "F3SINIGV" Then
            Grid.Columns.ColumnByFieldName("F3CONIGV").Value = Grid.Columns.ColumnByFieldName("F3SINIGV").Value * (1 + (nPorc / 100))
        'Else
            'Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value / (1 + (nPorc / 100))
        '    Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value * nPorc / 9
        'End If
        Grid.Columns.ColumnByFieldName("F3baseimp").Value = Grid.Columns.ColumnByFieldName("F3SINIGV").Value * Grid.Columns.ColumnByFieldName("F3canpro").Value
        Grid.Columns.ColumnByFieldName("F3monina").Value = 0
        Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3baseimp").Value * nPorc / 100
    End If
    Grid.Dataset.Post
End Select


CalculaTotal
End Sub




Private Sub CalculaTotal()
On Error GoTo Errores:
With Grid
    If .Dataset.Active = True Then
    
 'Grid.Dataset.Edit
    If .Columns.ColumnByFieldName("F5afecto").Value = True Then
        .Dataset.Edit
        .Columns.ColumnByFieldName("F3VALDESC").Value = (.Columns.ColumnByFieldName("F3BASEIMP").Value + .Columns.ColumnByFieldName("f3igv").Value) * .Columns.ColumnByFieldName("F3PORDESC").Value / 100
    Else
        .Dataset.Edit
        .Columns.ColumnByFieldName("F3VALDESC").Value = .Columns.ColumnByFieldName("F3monina").Value * .Columns.ColumnByFieldName("F3PORDESC").Value / 100
    End If
    'Grid.Dataset.Post
    If right(CmbTipDoc.Text, 2) = "02" Then
        .Columns.ColumnByFieldName("f3total").Value = .Columns.ColumnByFieldName("F3baseimp").Value + .Columns.ColumnByFieldName("F3monina").Value - .Columns.ColumnByFieldName("F3igv").Value - .Columns.ColumnByFieldName("f3VALDESC").Value
    Else
        .Columns.ColumnByFieldName("f3total").Value = .Columns.ColumnByFieldName("F3baseimp").Value + .Columns.ColumnByFieldName("F3monina").Value + .Columns.ColumnByFieldName("F3igv").Value - .Columns.ColumnByFieldName("f3VALDESC").Value
    End If
    .Dataset.Post
    End If
End With
Exit Sub

Errores:


End Sub

Private Sub Grid_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case "F3BASEIMP", "F3MONINA", "F3REDONDEO", "F3TOTAL", "F3IGV"
    Text = Format(Text, "###,###,##0.00")
Case "F3SINIGV", "F3CONIGV"
    If dxCheckBox2.Checked = False Then
        Text = Format(Text, "###,###,##0.0000")
    End If
End Select
End Sub

Private Sub Grid_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
If UCase(Column.FieldName) = "F3VALDESC" Or UCase(Column.FieldName) = "F3PORDESC" Or UCase(Column.FieldName) = "F3BASEIMP" Or UCase(Column.FieldName) = "F3MONINA" Or UCase(Column.FieldName) = "F3REDONDEO" Or UCase(Column.FieldName) = "F3TOTAL" Or UCase(Column.FieldName) = "F3IGV" Then
    If Mid(Cmbmone.Text, 1, 1) = "S" Then
        Color = &HC0FFFF
    Else
        Color = &HC0FFC0
    End If
    Font.Bold = True
    Text = Format(Text, "###,###,##0.00")
End If
End Sub

Private Sub Grid_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim sql         As String
    
    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)

        Case "F3CODPRO":
        
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
                Grid.Dataset.Edit
                Grid.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                Grid.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
                Grid.Columns.ColumnByFieldName("f3codmedida").Value = wmedida
                Grid.Columns.ColumnByFieldName("f3desmedida").Value = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida, "T", cnn_dbbancos)
                Grid.Columns.ColumnByFieldName("f5afecto").Value = IIf(wafecto = "*", True, False)
                Grid.Dataset.FieldValues("f3SINIGV") = Format(wvv_prod, "###,##0.0000")
                Grid.Dataset.FieldValues("f3canpro") = Format(0, "###,##0.0000")
                Grid.Dataset.FieldValues("f3redondeo") = 0
                Grid.Dataset.FieldValues("f3sinigv") = Format(0, "###,##0.0000")
                Grid.Dataset.FieldValues("f3conigv") = Format(0, "###,##0.0000")
                Grid.Dataset.FieldValues("f3baseimp") = Format(0, "###,##0.00")
                Grid.Dataset.FieldValues("f3monina") = Format(0, "###,##0.00")
            End If
        Case "F5DESCOSTO":
            wcodcosto = "": wdescosto = "": wunicosto = "":
            
            Ayuda_Centros.Show 1
                      
            If Len(Trim(wcodcosto)) > 0 Then
                Grid.Dataset.Edit
                Grid.Columns.ColumnByFieldName("F5CODCOSTO").Value = wcodcosto
                Grid.Columns.ColumnByFieldName("F5DESCOSTO").Value = wdescosto
                Grid.Dataset.Post
               
            End If
            'End If
            '****************
        Case "F5CODCTA"
            Dim gassto As String
            Dim rsgasto As New ADODB.Recordset
            Dim amovs(0 To 3)  As a_grabacion
            wctacont = "": wnomctacont = ""
            Ayuda_PlanCta.Show 1
            If Len(Trim(wctacont)) > 0 Then
                gassto = ObtenerCampo("BF9GIN", "CODIGO", "CUENTA", wctacont, "T", cnn_dbbancos)
                If Len(Trim(gassto)) = 0 Then
                    csql = "SELECT TOP 1 Val(CODIGO) AS num From BF9GIN ORDER BY Val(CODIGO) DESC"
                    If rsgasto.State = 1 Then rsgasto.Close
                    rsgasto.Open csql, cnn_dbbancos, 3, 1
                    If rsgasto.RecordCount > 0 Then
                        gassto = Format(CStr(rsgasto.Fields("num").Value + 1), "000")
                    End If
                    amovs(0).campo = "CODIGO": amovs(0).valor = gassto: amovs(0).TIPO = "T"
                    amovs(1).campo = "BASE": amovs(1).valor = "G": amovs(1).TIPO = "T"
                    amovs(2).campo = "NOMBRE": amovs(2).valor = wnomctacont: amovs(2).TIPO = "T"
                    amovs(3).campo = "CUENTA": amovs(3).valor = wctacont: amovs(3).TIPO = "T"
                    GRABA_REGISTRO amovs(), "BF9GIN", "A", 3, StrConexDbBancos, ""
                End If
                Grid.Dataset.Edit
                Grid.Columns.ColumnByFieldName("F3GASTO").Value = gassto
                Grid.Columns.ColumnByFieldName("F5CODCTA").Value = wctacont
                Grid.Dataset.Post
            End If
        
        Case "COD_SOLICITUD"
            
        FlagAcceso = False
        flagwin = True
        whelp_solicitud = 4
        FlagAcceso = False
        num_solcomp = ""
        ayuda_solicitudes_OC.Show 1
        Unload ayuda_solicitudes_OC
        Set ayuda_solicitudes_OC = Nothing
        
        If Len(Trim(num_solcomp)) > 0 Then
            Txt_Prove.Enabled = True
            Call MostrarDatos
            Dim nPorc As Double
            If right(CmbTipDoc.Text, 2) = "02" Then
                nPorc = gretenc
                If Grid.Columns.ColumnByFieldName("f3sinigv").Value >= 700 Then
                    Grid.Dataset.Edit
                    Grid.Columns.ColumnByFieldName("F5afecto").Value = True
                    Grid.Dataset.Post
                Else
                    Grid.Dataset.Edit
                    Grid.Columns.ColumnByFieldName("F5afecto").Value = False
                    Grid.Dataset.Post
                End If
                    
            Else
                nPorc = wwigv
            End If
            Grid.Dataset.First
            Do While Not Grid.Dataset.EOF
                Grid.Dataset.Edit
                If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
                    If right(CmbTipDoc.Text, 2) = "02" Then
                        Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
                    End If
                    Grid.Columns.ColumnByFieldName("F3monina").Value = 0
                    Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
                    Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
                Else
                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
                    Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
                    Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
                    Grid.Columns.ColumnByFieldName("F3igv").Value = 0
                End If
                Grid.Dataset.Post
                CalculaTotal
                Grid.Dataset.Next
            Loop
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
            calcula
            sw_nuevo_item = False
        End If
    End If
    Select Case UCase(Grid.Columns.FocusedColumn.Caption)
        Case "?"
        Codigo_producto = Grid.Columns.ColumnByFieldName("F3CODPRO").Value
        ayuda_prov_prod.Show 1
        Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("F3sinigv").Value = wvv_prod
            Grid.Columns.ColumnByFieldName("F3conIGV").Value = wpv_prod
            If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
                If right(CmbTipDoc.Text, 2) = "02" Then
                    Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
                End If
                Grid.Columns.ColumnByFieldName("F3monina").Value = 0
                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
                Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (wwigv / 100)
            Else
                Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
                Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
                Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
                Grid.Columns.ColumnByFieldName("F3igv").Value = 0
            End If
            Grid.Columns.ColumnByFieldName("f3total").Value = Grid.Columns.ColumnByFieldName("F3baseimp").Value + Grid.Columns.ColumnByFieldName("F3monina").Value + Grid.Columns.ColumnByFieldName("F3igv").Value - Grid.Columns.ColumnByFieldName("f3VALDESC").Value
        Grid.Dataset.Post
    End Select
End Sub

Private Sub Grid_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim nPorc As Double
Dim nValor As Double
If right(CmbTipDoc.Text, 2) = "02" Then
    If left(Cmbmone.Text, 1) = "D" Then
        nValor = Val(Grid.Columns.ColumnByFieldName("f3sinigv").Value & "") * Val(txt_tc.Text)
    Else
        nValor = Val(Grid.Columns.ColumnByFieldName("f3sinigv").Value & "")
    End If
    nPorc = gretenc
    'If nValor >= 700 Then
    '    Grid.Dataset.Edit
    '    Grid.Columns.ColumnByFieldName("F5afecto").Value = True
    '    Grid.Dataset.Post
    'Else
    '    Grid.Dataset.Edit
    '    Grid.Columns.ColumnByFieldName("F5afecto").Value = False
    '    Grid.Dataset.Post
    'End If
        
Else
    nPorc = wwigv
End If
Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
Case "F3CANPRO"
    Grid.Dataset.Edit
    If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
        Grid.Columns.ColumnByFieldName("F3monina").Value = 0
        Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
    Else
        Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
        Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = 0
    End If
    Grid.Dataset.Post
Case "F3CONIGV"
    Grid.Dataset.Edit
    Grid.Columns.ColumnByFieldName("F3COLMOD").Value = UCase(Grid.Columns.FocusedColumn.FieldName)
    If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
        If right(CmbTipDoc.Text, 2) = "02" Then
            Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value * nPorc / 9
        Else
            Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value / (1# + (nPorc / 100))
        End If
        Grid.Columns.ColumnByFieldName("F3monina").Value = 0
        Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
    Else
        Grid.Columns.ColumnByFieldName("F3SINIGV").Value = Grid.Columns.ColumnByFieldName("F3CONIGV").Value
        Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
        Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = 0
    End If
    Grid.Dataset.Post
Case "F3SINIGV"
    Grid.Dataset.Edit
    Grid.Columns.ColumnByFieldName("F3COLMOD").Value = UCase(Grid.Columns.FocusedColumn.FieldName)
    If Grid.Columns.ColumnByFieldName("F5afecto").Value = True Then
        If right(CmbTipDoc.Text, 2) = "02" Then
            Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value - (Grid.Columns.ColumnByFieldName("F3sinIGV").Value * ((nPorc / 100)))
        Else
            Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value * (1# + (nPorc / 100))
        End If
        Grid.Columns.ColumnByFieldName("F3monina").Value = 0
        Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = Grid.Columns.ColumnByFieldName("F3BASEIMP").Value * (nPorc / 100)
    Else
        Grid.Columns.ColumnByFieldName("F3conIGV").Value = Grid.Columns.ColumnByFieldName("F3sinIGV").Value
        Grid.Columns.ColumnByFieldName("F3BASEIMP").Value = 0
        Grid.Columns.ColumnByFieldName("F3monina").Value = Grid.Columns.ColumnByFieldName("F3CANPRO").Value * Grid.Columns.ColumnByFieldName("F3sinigv").Value
        Grid.Columns.ColumnByFieldName("F3igv").Value = 0
    End If
    Grid.Dataset.Post
Case "F3PORDESC"
    
End Select
CalculaTotal
End Sub





Private Sub txt_fecha_LostFocus()

    If IsDate(txt_fecha.Value) Then
        If Val(txt_tc.Text & "") = 0# Then
            If rscambios.State = adStateOpen Then rscambios.Close
            If ctipoadm_bd = "M" Then
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE FECHA='" & txt_fecha.Value & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            Else
                rscambios.Open "SELECT CAMBIO FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & txt_fecha.Value & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            End If
            If Not rscambios.EOF Then
                txt_tc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
            Else
                txt_tc.Text = Format(2.81, "0.000")
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

Private Sub Txt_Prove_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub Txt_Prove_GotFocus()

    Txt_Prove.SelStart = 0: Txt_Prove.SelLength = Len(Txt_Prove)
    
End Sub

Private Sub Txt_Prove_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(Txt_Prove.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2NOMPROV,F2DIRPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & Trim(Txt_Prove.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rst.EOF Then
                pnlnomprv.Caption = "" & rst.Fields("F2NOMPROV")
                pnldireprv.Caption = "" & rst.Fields("F2DIRPROV")
                GRABA_GRID Trim(Txt_Prove.Text)
            Else
                MsgBox "El proveedor no existe. Verifique.", vbInformation, "Atención"
                Txt_Prove.SetFocus
            End If
            If rst.State = adStateOpen Then rst.Close
        End If
    End If

End Sub

Private Sub Txt_Referencia_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txt_tc_Change()

    If Not inicio Then swGrabacion = True
    
    If txt_tc.Text = "" Then
        txt_tc.Text = "2.810"
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
End Sub

Private Sub txtcodcosto_KeyPress(KeyAscii As Integer)

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
                'SendKeys "{tab}"
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

    If sw_ayuda = False Then
        If Len(Trim(txtcodforma.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG='" & Trim(txtcodforma.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rst.EOF Then
                pnlnomforma.Caption = Trim("" & rst!F2DESPAG)
            Else
                pnlnomforma.Caption = ""
                MsgBox "Còdigo de forma de pago no existe. Verifique.", vbInformation, "Atenciòn"
                txtcodforma.SetFocus
            End If
            rst.Close
        End If
    End If

End Sub

Private Sub txtcodsoli_Change()
txtcodsoli_LostFocus
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

    If Len(Trim(txtcodsoli.Text)) > 0 Then
        If rst.State = adStateOpen Then rst.Close
        rst.Open "SELECT * FROM ef2users WHERE f2coduser='" & Trim(txtcodsoli.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rst.EOF Then
            pnlnomsoli.Caption = "" & rst.Fields("f2nomuser")
        Else
            pnlnomsoli.Caption = ""
            MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
            'txtcodsoli.SetFocus
        End If
        rst.Close
    End If

End Sub

Private Sub txtcontacto_GotFocus()
    txtcontacto.SelStart = 0: txtcontacto.SelLength = Len(txtcontacto)

End Sub

Private Sub txtcontacto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtCotizacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtempresa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtlugar_entrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtobserva_Change()

    If Not inicio Then swGrabacion = True

End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
    
End Sub

Private Sub Txt_Fecha_Change()
    
    wgraba = 0
    If Not inicio Then swGrabacion = True
    txt_tc.Text = Format(ObtenerCampo("CAMBIOS", "CAMBIO", "FECHA", txt_fecha.Value, "F", cnn_dbbancos), "0.000")

End Sub

Private Sub Txt_Fecha_GotFocus()
    
    'txt_fecha.FocusSelect = True
    
End Sub

Private Sub Txt_Fecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
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
                SendKeys "{tab}"
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
        
            Call MostrarDatos
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
        Call MostrarDatos
        Txt_Prove.Text = ""
        pnlnomprv.Caption = ""
        pnldireprv.Caption = ""
        
        End If
        SendKeys "{tab}"
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
    
    If loc = 1 Then
        With rsOrdenCab
            If Not (.EOF) Then
                txtempresa = !F4EMPRESA & ""
                If Txt_NumOC = "" Then
                    !F4NUMORD = " "
                Else
                    Txt_NumOC = (!F4NUMORD & "")
                End If
                Txt_NumSolComp = !F4CODSOLICITUD & ""
                txt_fecha.Value = !F4FECEMI
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
                txtCotizacion.Text = !F4numcotiza & ""
                TxtCodCosto.Text = !F4CENTRO & ""
                'txtcodcosto_KeyPress 13
                abofechaentrega.Value = Format(!F4FECENT, "DD/MM/YYYY")
                
                If loc = 2 Then
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                Else
                    TxtIgv = Format$(!F4IGV & "", "#,##0.00")
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
                If loc = 1 Then
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
                        rst.Open "SELECT P.f5nompro,P.f5codfab,P.F7codmed,M.F2DESMAR from if5pla P, EF2MARCAS M where P.f5codpro='" & rsOrdenDet!f3codpro & "' AND P.F5MARCA=M.F2CODMAR", cnn_dbbancos, adOpenDynamic, adLockOptimistic
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
                            dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.Value, "dd/mm/yyyy"))
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
    If loc = 1 Then
        With rsOrdenCab
            If Not (.EOF) Then
                txtempresa = !F4EMPRESA & ""
                If Txt_NumOC = "" Then
                    !F4NUMORD = " "
                Else
                    Txt_NumOC = (!F4NUMORD & "")
                End If
                Txt_NumSolComp = !F4CODSOLICITUD & ""
                txt_fecha.Value = !F4FECEMI
                txtobserva.Text = rsOrdenCab!F4OBSERVA & ""
                txtcontacto.Text = rsOrdenCab!F4CONTACTO & ""
                If !F4TIPMON = "S" Then
                    Cmbmone.ListIndex = 0
                Else
                    Cmbmone.ListIndex = 1
                End If
                FrameOC.FontBold = True
'                Select Case !f4estnul & ""
'                Case "S"
'                    FrameOC.Caption = "Anulado"
'                    FrameOC.ForeColor = vbRed
'                    atbmenu.Tools.ITEM("ID_Nuevo").Visible = True
'                    atbmenu.Tools.ITEM("ID_Grabar").Visible = False
'                    atbmenu.Tools.ITEM("ID_Anular").Visible = False
'                    atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
'                    atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
'                    atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = False
'                    atbmenu.Tools.ITEM("ID_Aprobacion").Visible = False
'                Case "P"
'                    FrameOC.Caption = "Pendiente de Aprobación"
'                    FrameOC.ForeColor = &HC0FFC0
'                    atbmenu.Tools.ITEM("ID_Nuevo").Visible = True
'                    atbmenu.Tools.ITEM("ID_Grabar").Visible = False
'                    atbmenu.Tools.ITEM("ID_Anular").Visible = False
'                    atbmenu.Tools.ITEM("ID_Eliminar").Visible = False
'                    atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
'                    atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = False
'                    atbmenu.Tools.ITEM("ID_Aprobacion").Visible = False
'                Case "N"
'                    FrameOC.Caption = "Sin Aprobación"
'                    FrameOC.ForeColor = vbBlue
'                    atbmenu.Tools.ITEM("ID_Nuevo").Visible = True
'                    atbmenu.Tools.ITEM("ID_Grabar").Visible = True
'                    atbmenu.Tools.ITEM("ID_Anular").Visible = False
'                    atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
'                    atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
'                    atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = True
'                    atbmenu.Tools.ITEM("ID_Aprobacion").Visible = True
'                Case "R", "A"
'                    If !f4estnul & "" = "R" Then
'                        FrameOC.Caption = "Rechazado"
'                        FrameOC.ForeColor = vbMagenta
'                    Else
'                        FrameOC.Caption = "Aprobado"
'                        FrameOC.ForeColor = vbBlack
'                    End If
'
'                    atbmenu.Tools.ITEM("ID_Nuevo").Visible = True
'                    atbmenu.Tools.ITEM("ID_Grabar").Visible = False
'                    atbmenu.Tools.ITEM("ID_Anular").Visible = True
'                    atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
'                    atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
'                    atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = True
'                    atbmenu.Tools.ITEM("ID_Aprobacion").Visible = False
'                End Select
                    atbmenu.Tools.ITEM("ID_Nuevo").Visible = True
                    atbmenu.Tools.ITEM("ID_Grabar").Visible = True
                    atbmenu.Tools.ITEM("ID_Anular").Visible = True
                    atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
                    atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
'                    atbmenu.Tools.ITEM("ID_RenovarOrden").Visible = True
'                    atbmenu.Tools.ITEM("ID_Aprobacion").Visible = True

                txt_tc = Format$(!F4TIPCAM, "0.000") & ""
                txtcodforma = !F4FORPAG & ""
                Txt_Referencia = !F4REFERE & ""
                txtcodsoli = !F4CODSOL & ""
                txtCotizacion.Text = !F4numcotiza & ""
                TxtCodCosto.Text = !F4CENTRO & ""
                PnlNomCosto = ObtenerCampo("centros", "F3DESCRIP", "f3costo", TxtCodCosto.Text, "T", cnn_dbbancos)
                If Not IsNull(!F4FECENT) Then
                    abofechaentrega.Value = Format(!F4FECENT, "DD/MM/YYYY")
                Else
                    abofechaentrega.CheckBox = True
                    abofechaentrega.Value = Empty
                End If
                'aBoHoraEntrega.Value = Format(!F4FECENT, "hh:mm:ss")
                txtFechaPago.Text = !F4DIAPAGO & ""
                If !F4REGULARIZA = "1" Then
                    ChK_regularizacion.Checked = True
                Else
                    ChK_regularizacion.Checked = False
                End If
                If !F4PAGOPARCIAL = True Then
                    Chk_pagoparcial.Checked = True
                Else
                    Chk_pagoparcial.Checked = False
                End If
                If loc = 2 Then
                    txtbase = Format$(!F4BASIMP & "", "#,##0.00")
                Else
                    TxtIgv = Format$(!F4IGV & "", "#,##0.00")
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
                
                xnombre = rsOrdenCab!F4CODSOL & ""
                csql = "SELECT F2NOMUSER from ef2userS where f2coduser='" & UCase(Trim(xnombre)) & "'"
                Set RsU = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
                If Not (RsU.EOF) Then
                    txtcodsoli = UCase(xnombre)
                    pnlnomsoli.Caption = RsU!F2NOMUSER & ""
                End If
                RsU.Close
                
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT F2DESPAG from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not (rst.EOF) Then
                    pnlnomforma.Caption = "" & rst.Fields("F2DESPAG")
                    wgraba = 0
                End If
                rst.Close
                
                SeleccionaEnComboRight rsOrdenCab!f4tipdoc & "", CmbTipDoc
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
        sql = "SELECT * from if3orden where f4numord='" & GOC & "' AND F4local = '" & TOC & "' ORDER BY val(ITEM)"
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
                If loc = 1 Then
                    If rsOrdenDet.Fields("f4numord") = GOC Then
                        Amov(0).campo = "item": Amov(0).valor = i & "": Amov(0).TIPO = "N"
                        Amov(1).campo = "f3codpro": Amov(1).valor = .Fields("f3codpro") & "": Amov(1).TIPO = "T"
                        Amov(2).campo = "f5codcosto": Amov(2).valor = .Fields("F3CENCOS") & "": Amov(2).TIPO = "T"
                        
                        RsCC.Filter = adFilterNone
                        RsCC.Filter = "f3costo='" & .Fields("F3CENCOS") & "" & "'"
                        If RsCC.RecordCount > 0 Then
                            Amov(15).campo = "f5descosto": Amov(15).valor = RsCC!f3abrev & "": Amov(15).TIPO = "T"
                        Else
                            Amov(15).campo = "f5descosto": Amov(15).valor = "": Amov(15).TIPO = "T"
                        End If
                        'If rst.State = adStateOpen Then rst.Close
                        'rst.Open "SELECT P.f5nompro,P.f5codfab,P.F7codmed,M.F2DESMAR from if5pla P, EF2MARCAS M where P.f5codpro='" & rsOrdenDet!f3codpro & "' AND P.F5MARCA=M.F2CODMAR", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        'If Not (rst.EOF) Then
                         '   Amov(3).campo = "f5nompro": Amov(3).valor = rst.Fields("f5nompro") & "": Amov(3).TIPO = "T"
                          '  Amov(4).campo = "f3codmedida": Amov(4).valor = rst!f7codmed & "": Amov(4).TIPO = "T"
                        'Else
                        Amov(3).campo = "f5nompro": Amov(3).valor = .Fields("f5nompro") & "": Amov(3).TIPO = "T"
                        Amov(4).campo = "f3codmedida": Amov(4).valor = .Fields("UNIDAD") & "": Amov(4).TIPO = "T"
                        RsMed.Filter = adFilterNone
                        RsMed.Filter = "f7sigmed='" & .Fields("UNIDAD") & "" & "'"
                        If RsMed.RecordCount > 0 Then
                            Amov(14).campo = "f3desmedida": Amov(14).valor = RsMed!f7sigmed & "": Amov(14).TIPO = "T"
                        Else
                            Amov(14).campo = "f3desmedida": Amov(14).valor = "": Amov(14).TIPO = "T"
                        End If
                        'End If
                        'rst.Close
                                                                           
                        Amov(5).campo = "f3canpro": Amov(5).valor = .Fields("f3canpro") & "": Amov(5).TIPO = "N"
                        Amov(6).campo = "f3sinigv": Amov(6).valor = Val("" & .Fields("f3precos")): Amov(6).TIPO = "N"
                        Amov(7).campo = "f3conigv": Amov(7).valor = Val("" & .Fields("f3preuni")): Amov(7).TIPO = "N"
                        Amov(8).campo = "f5afecto": Amov(8).valor = IIf(.Fields("f5afecto") & "" = "*", -1, 0): Amov(8).TIPO = "N"
                                                                           
                        If Trim(.Fields("f5afecto") & "") = "*" Then
                            Amov(9).campo = "f3baseimp": Amov(9).valor = Val("" & .Fields("f5valvta")): Amov(9).TIPO = "N"
                            Amov(10).campo = "f3monina": Amov(10).valor = "0": Amov(10).TIPO = "N"
                        Else
                            Grid.Dataset.FieldValues("f3baseimp") = 0
                            Grid.Dataset.FieldValues("f3monina") = Val("" & .Fields("f5valvta"))
                            Amov(9).campo = "f3baseimp": Amov(9).valor = 0: Amov(9).TIPO = "N"
                            Amov(10).campo = "f3monina": Amov(10).valor = Val("" & .Fields("f5valvta")): Amov(10).TIPO = "N"
                        End If
                        
                        Amov(11).campo = "f3igv": Amov(11).valor = Val("" & .Fields("f3igv")): Amov(11).TIPO = "N"
                        Amov(12).campo = "f3total": Amov(12).valor = Val("" & .Fields("f3total")): Amov(12).TIPO = "N"
                        Amov(13).campo = "f3colmod": Amov(13).valor = ("" & .Fields("f3backorder")): Amov(13).TIPO = "T"
                        Amov(16).campo = "f3valdesc": Amov(16).valor = Val("" & .Fields("f3totdct")): Amov(16).TIPO = "N"
                        Amov(17).campo = "f3pordesc": Amov(17).valor = Val("" & .Fields("f3pordct")): Amov(17).TIPO = "N"
                        Amov(18).campo = "f3observa": Amov(18).valor = "" & .Fields("f3observa"): Amov(18).TIPO = "T"
                        Amov(19).campo = "cod_solicitud": Amov(19).valor = "" & .Fields("cod_solicitud"): Amov(19).TIPO = "T"
                    Else
                        Exit Do
                    End If
                    
                End If
                GRABA_REGISTRO_noenvia Amov, "tmpOrdendeCompra", "A", 19, CnTmp, ""
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
    Grid.Dataset.ADODataset.CommandText = "select * from tmpOrdendeCompra order by item"
    Grid.Dataset.Active = True
    dxDBGrid1.KeyField = "item"
    Grid.Dataset.Close
    Grid.Dataset.Open
    If Grid.Dataset.RecordCount = 0 Then
        AdicionaItemGrid
    End If
    
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


Private Sub CONFIGURA_GRID_OC()
    
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
Dim StrCodCosto As String

    StrCodCosto = ObtenerCampo("centros", "cconcar", "f3costo", TxtCodCosto.Text, "T", cnn_dbbancos)
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
        wmes = Month(txt_fecha.Value)
        Orden = Format(Val(rst.Fields("MAYOR") & "") + 1, "000000000000") ' & "/" & StrCodCosto
    Else
        'Orden = wanno & "-" & Format(1, "00000") & "/0"
        wmes = Month(txt_fecha.Value)
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
Dim amovs_cab(0 To 31)  As a_grabacion
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
    sql = "select sum(f3total) as ztotal from tmpOrdendeCompra "
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
    If loc = 1 Then
        Select Case jc
            Case 0
            Txt_NumOC.Text = Nueva_orden
            Txt_TOC.Text = TOC
        End Select
    End If
    
    
    If loc = 1 Then
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
            amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = wNumOc: amovs_cab(0).TIPO = "T"
        Else
            amovs_cab(0).campo = "F4NUMORD": amovs_cab(0).valor = Txt_NumOC.Text: amovs_cab(0).TIPO = "T"
        End If
        If ctipo = "A" Then
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = "N": amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = "1": amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = 0: amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECGRA": amovs_cab(4).valor = Format(Now, "dd/MM/yyyy hh:mm:ss"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEGRA": amovs_cab(5).valor = wusuario: amovs_cab(5).TIPO = "T"
        Else
            amovs_cab(1).campo = "F4ESTNUL": amovs_cab(1).valor = rsOrdenCab.Fields("F4ESTNUL"): amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4FALTA": amovs_cab(2).valor = rsOrdenCab.Fields("F4FALTA"): amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4ESTVAL": amovs_cab(3).valor = rsOrdenCab.Fields("F4ESTVAL"): amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECMOD": amovs_cab(4).valor = Format(Now, "dd/MM/yyyy hh:mm:ss"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4USEMOD": amovs_cab(5).valor = wusuario: amovs_cab(5).TIPO = "T"
        End If
        
        amovs_cab(6).campo = "F4CODSOL": amovs_cab(6).valor = txtcodsoli.Text: amovs_cab(6).TIPO = "T"
        amovs_cab(7).campo = "F4FECEMI": amovs_cab(7).valor = Format(txt_fecha.Value, "DD/MM/YYYY"): amovs_cab(7).TIPO = "F"
        amovs_cab(8).campo = "F4CODPRV": amovs_cab(8).valor = Txt_Prove: amovs_cab(8).TIPO = "T"
        amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = txt_tc.Text: amovs_cab(9).TIPO = "N"
        amovs_cab(10).campo = "F4FORPAG": amovs_cab(10).valor = txtcodforma.Text: amovs_cab(10).TIPO = "T"
        amovs_cab(11).campo = "F4REFERE": amovs_cab(11).valor = Txt_Referencia.Text: amovs_cab(11).TIPO = "T"
        amovs_cab(12).campo = "F4OBSERVA": amovs_cab(12).valor = txtobserva.Text: amovs_cab(12).TIPO = "T"
        amovs_cab(13).campo = "F4CODSOLICITUD": amovs_cab(13).valor = Trim(Txt_NumSolComp.Text): amovs_cab(13).TIPO = "T"
        amovs_cab(14).campo = "F4TIPMON": amovs_cab(14).valor = IIf(Cmbmone.ListIndex = 0, "S", "D"): amovs_cab(14).TIPO = "T"
        amovs_cab(15).campo = "F4IGV": amovs_cab(15).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3igv").SummaryFooterValue, "0.00")): amovs_cab(15).TIPO = "N"
        amovs_cab(16).campo = "F4MONINA": amovs_cab(16).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3monina").SummaryFooterValue, "0.00")): amovs_cab(16).TIPO = "N"
        amovs_cab(17).campo = "F4BASIMP": amovs_cab(17).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3baseimp").SummaryFooterValue, "0.00")): amovs_cab(17).TIPO = "N"
        amovs_cab(18).campo = "F4MONTO": amovs_cab(18).valor = Val(Format(Grid.Columns.ColumnByFieldName("f3total").SummaryFooterValue, "0.00")): amovs_cab(18).TIPO = "N"
        amovs_cab(19).campo = "F4LOCAL": amovs_cab(19).valor = Txt_TOC.Text: amovs_cab(19).TIPO = "T"
        amovs_cab(20).campo = "F4EMPRESA": amovs_cab(20).valor = txtempresa.Text: amovs_cab(20).TIPO = "T"
        amovs_cab(21).campo = "F4NUMCOTIZA": amovs_cab(21).valor = txtCotizacion.Text: amovs_cab(21).TIPO = "T"
        amovs_cab(22).campo = "F4LUGAR_ENTREGA": amovs_cab(22).valor = txtlugar_entrega.Text: amovs_cab(22).TIPO = "T"
        amovs_cab(23).campo = "F4CONTACTO": amovs_cab(23).valor = txtcontacto.Text: amovs_cab(23).TIPO = "T"
        amovs_cab(24).campo = "F4FECENT": amovs_cab(24).valor = Format(abofechaentrega.Value, "DD/MM/YYYY"): amovs_cab(24).TIPO = "F"
        amovs_cab(25).campo = "F4RND": amovs_cab(25).valor = 0: amovs_cab(25).TIPO = "N"
        amovs_cab(26).campo = "F4CENTRO": amovs_cab(26).valor = (TxtCodCosto.Text & ""): amovs_cab(26).TIPO = "T"
        amovs_cab(27).campo = "F4TIPDOC": amovs_cab(27).valor = right((CmbTipDoc.Text & ""), 2): amovs_cab(27).TIPO = "T"
        amovs_cab(28).campo = "F4REGULARIZA": amovs_cab(28).valor = IIf(ChK_regularizacion.Checked = False, 0, 1): amovs_cab(28).TIPO = "T"
        amovs_cab(29).campo = "F4DIAPAGO": amovs_cab(29).valor = txtFechaPago.Text: amovs_cab(29).TIPO = "T"
        amovs_cab(30).campo = "F4PAGOPARCIAL": amovs_cab(30).valor = IIf(Chk_pagoparcial.Checked = False, 0, 1): amovs_cab(30).TIPO = "T"
        amovs_cab(31).campo = "F4NOMPROV": amovs_cab(31).valor = pnlnomprv.Caption: amovs_cab(31).TIPO = "T"
        rsOrdenCab.Close
        
        
        GRABA_REGISTRO_logistica amovs_cab(), "IF4ORDEN", ctipo, 31, cnn_dbbancos, "F4NUMORD = '" & Txt_NumOC.Text & "' AND F4local = '" & TOC & "'"
        
        
    End If
    
    '---------- GRABANDO EL DETALLE DE LA ORDEN DE COMPRA ----------------------'
    If ctipoadm_bd = "M" Then
        If SwRenovar = True Then
            sql = ("delete from if3orden where f4numord= '" & wNumOc & "'  AND F4local = '" & TOC & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        Else
            sql = ("delete from if3orden where f4numord= '" & Txt_NumOC.Text & "'  AND F4local = '" & TOC & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        End If
    Else
        If SwRenovar = True Then
            sql = ("delete * from if3orden where f4numord= '" & wNumOc & "'  AND F4local = '" & TOC & "'")
             cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
           Actualiza_Log sql, cnn_dbbancos.ConnectionString
        Else
            sql = ("delete * from if3orden where f4numord= '" & Txt_NumOC.Text & "'  AND F4local = '" & TOC & "'")
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
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
                            ocompra = Val(rstaux.Fields("f3ocompra").Value)
                            rstaux.Fields("f3ocompra").Value = ocompra + wcantidad
                        Else             'Modifica
                            rstaux.Fields("f3ocompra").Value = wcantidad
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
                        sql = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD"
                        sql = sql & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL"
                        sql = sql & ",F3FENTREGA,item,F5NOMPRO, F4LOCAL,F3CENCOS,F3OBSERVA,cod_solicitud) VALUES "
                        sql = sql & "('" & wNumOc & "','" & .Fields("f3codpro") & "','" & .Fields("f3codpro") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & ",'','" & .Fields("f3codmedida") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & "," & .Fields("f3COnigv") & "," & .Fields("f3SInigv") & ","
                        sql = sql & Val(.Fields("f3pordesc") & "") & "," & Val(.Fields("f3valdesc") & "") & ","
                        sql = sql & Val(.Fields("f3baseimp") & "") + Val(.Fields("f3monina") & "") & ",'" & IIf((.Fields("f5afecto" & "")) = False, " ", "*") & "',"
                        sql = sql & Val(.Fields("f3igv") & "") & "," & Val(.Fields("f3total") & "") & ",null," & Val(.Fields("item") & "") & ",'"
                        sql = sql & .Fields("f5nompro") & "','1','"
                        If TxtCodCosto.Text = "998" Then
                            sql = sql & .Fields("f3cencos") & "','"
                        Else
                            sql = sql & TxtCodCosto.Text & "','"
                        End If
                        sql = sql & left(.Fields("f3OBSERVA"), 255) & "','" & .Fields("cod_solicitud") & "')"
                    Else
                        sql = "INSERT INTO IF3ORDEN (F4NUMORD,F3CODPRO,F3CODFAB,F3CANPRO,F5MARCA,UNIDAD"
                        sql = sql & ",F3CANFAL,F3PREUNI,F3PRECOS,F3PORDCT,F3TOTDCT,F5VALVTA,F5AFECTO,F3IGV,F3TOTAL"
                        sql = sql & ",F3FENTREGA,item,F5NOMPRO, F4LOCAL,F3CENCOS,F3OBSERVA,cod_solicitud) VALUES "
                        sql = sql & "('" & Txt_NumOC.Text & "','" & .Fields("f3codpro") & "','" & .Fields("f3codpro") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & ",'','" & .Fields("f3codmedida") & "',"
                        sql = sql & Val(.Fields("f3canpro") & "") & "," & Val(.Fields("f3COnigv") & "") & "," & Val(.Fields("f3SInigv") & "") & ","
                        sql = sql & Val(.Fields("f3pordesc") & "") & "," & Val(.Fields("f3valdesc") & "") & ","
                        sql = sql & Val(.Fields("f3baseimp") & "") + Val(.Fields("f3monina") & "") & ",'" & IIf((.Fields("f5afecto" & "")) = False, " ", "*") & "',"
                        sql = sql & Val(.Fields("f3igv") & "") & "," & Val(.Fields("f3total") & "") & ",null," & Val(.Fields("item") & "") & ",'"
                        sql = sql & .Fields("f5nompro") & "','" & TOC & "','"
                        If TxtCodCosto.Text = "998" Then
                            sql = sql & .Fields("f3cencos") & "','"
                        Else
                            sql = sql & TxtCodCosto.Text & "','"
                        End If
                        sql = sql & left(.Fields("f3OBSERVA"), 255) & "','" & .Fields("cod_solicitud") & "')"
                    End If
                        
                    cnn_dbbancos.Execute sql
                    AlmacenaQuery_sql sql, cnn_dbbancos
                    Actualiza_Log sql, cnn_dbbancos.ConnectionString

                If Not IsNull(.Fields("cod_solicitud")) And right(soli_acum, 12) <> .Fields("cod_solicitud") Then
                    If Len(Trim(soli_acum)) > 1 Then
                        soli_acum = soli_acum & ", " & .Fields("cod_solicitud")
                        sql = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC & "' where cod_solicitud='" & .Fields("cod_solicitud") & "'"
                        cnn_dbbancos.Execute sql
                        AlmacenaQuery_sql sql, cnn_dbbancos
                        Actualiza_Log sql, cnn_dbbancos.ConnectionString
                    Else
                        soli_acum = .Fields("cod_solicitud")
                        sql = "update tb_cabsolicitud set cs_orden='" & Txt_NumOC & "' where cod_solicitud='" & .Fields("cod_solicitud") & "'"
                        cnn_dbbancos.Execute sql
                        AlmacenaQuery_sql sql, cnn_dbbancos
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
                    AlmacenaQuery_sql sql, cnn_dbbancos
                    Actualiza_Log sql, cnn_dbbancos.ConnectionString
    
                    '**********************************************************
                    rst.Close
                
                    If rst.State = adStateOpen Then rst.Close
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
        AlmacenaQuery_sql sql, cnn_dbbancos
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
                        AlmacenaQuery_sql sql, cnn_dbbancos
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
''                AlmacenaQuery_sql sql, cnn_dbbancos
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
                CODPROV = Txt_Prove.Text
                NOMPROV = pnlnomprv.Caption
                codprod = .Fields("f3codpro") & ""
                NomProd = .Fields("f5nompro") & ""
                cmoneda = IIf(Cmbmone.ListIndex = 0, "S", "D")
                dfecha = Format(txt_fecha.Value, "DD/MM/YYYY")
                'nprecos = Val("" & .Fields("F3PRECOS"))
                If rsproductos.State = adStateOpen Then rsproductos.Close
                rsproductos.Open "SELECT F5CODFAB,F7codmed FROM IF5PLA WHERE F5CODPRO='" & codprod & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsproductos.EOF Then
                    ccodfab = left("" & rsproductos.Fields("F5CODFAB"), 15)
                    ccodmed = "" & rsproductos.Fields("F7codmed")
                End If
                rsproductos.Close
                    
                If rst.State = adStateOpen Then rst.Close
                rst.Open "SELECT * FROM EF2PROD_PROV WHERE F5CODPRO='" & codprod & "' AND " _
                & "F2CODPRV='" & CODPROV & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If rst.RecordCount = 0 Then
'                    rst.AddNew
'                    rst!F2CODPRV = CodProv
'                    rst!F2NOMPRV = NOMPROV
'                    rst!f5codpro = codprod
'                    rst!f5nompro = NomProd
'                    rst!f5valvta = nprecos
'                    rst.Fields("F2MONEDA") = cmoneda
'                    rst.Fields("F2FECHA") = dfecha
'                    rst!f5codfab = ccodfab
'                    rst!f7codmed = ccodmed
'                    rst.Fields("F2COND_PAGO") = txtcodforma.Text
'                    rst.Fields("F2FORPAG") = txtcodforma.Text
'                    rst.Update
                    If ctipoadm_bd = "M" Then
                        sql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F2MONEDA,F2FECHA,F5CODFAB,F7CODMED,F2COND_PAGO,F2FORPAG) VALUES " _
                            & "('" & CODPROV & "','" & NOMPROV & "','" & codprod & "'," & nprecos & ",'" & cmoneda & "','" & dfecha & "','" _
                            & ccodfab & "','" & ccodmed & "','" & txtcodforma.Text & "','" & txtcodforma.Text & "')"
                    Else
                        sql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F2MONEDA,F2FECHA,F5CODFAB,F7CODMED,F2COND_PAGO,F2FORPAG) VALUES " _
                            & "('" & CODPROV & "','" & NOMPROV & "','" & codprod & "'," & nprecos & ",'" & cmoneda & "',CVDATE('" & dfecha & "'),'" _
                            & ccodfab & "','" & ccodmed & "','" & txtcodforma.Text & "','" & txtcodforma.Text & "')"
                    End If
                Else
                    If ctipoadm_bd = "M" Then
                        sql = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA='" & dfecha & "' WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CODPROV & "'"
                    Else
                        sql = "UPDATE EF2PROD_PROV SET F5VALVTA=" & nprecos & ",F2MONEDA='" & cmoneda & "',F2FECHA=CVDATE('" & dfecha & "') WHERE F5CODPRO='" & codprod & "' AND F2CODPRV='" & CODPROV & "'"
                    End If
                    cnn_dbbancos.Execute (sql)
                    AlmacenaQuery_sql sql, cnn_dbbancos
                    Actualiza_Log sql, cnn_dbbancos.ConnectionString
                End If
                rst.Close
                .MoveNext
            Loop
        End With
    End If
    rsdetaoc.Close
    
End Sub

Private Sub Txt_Prove_DblClick()

    Txt_Prove_KeyDown 113, 0
        'atbmenu.Tools.ITEM("ID_Grabar").Visible = True
        'atbmenu.Tools.ITEM("ID_Imprimir").Visible = True
        'atbmenu.Tools.ITEM("IDEmail").Enabled = True
        'atbmenu.Tools.ITEM("ID_Eli").Visible = True
End Sub

Private Sub Txt_Prove_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim q As Integer
    If KeyCode = 113 Then
        sw_ayuda = True
        sw_ocompra = False
        'hlp_proveedores.Show 1
        ayuda_proveedores_ocl.Show 1
        sw_ayuda = False
        Txt_Prove.Text = wrucprov
        pnlnomprv.Caption = wnomprov
        pnldireprv.Caption = wdirprov
        txtcontacto.Text = wcontacto
        For q = 0 To CmbTipDoc.ListCount - 1
            CmbTipDoc.ListIndex = q
            If right(CmbTipDoc.Text, 2) = wdcto Then 'wdcto
                CmbTipDoc.ListIndex = q
                Exit For
            End If
        Next
        If Len(Trim(wfpagoprov)) > 0 Then
            txtcodforma.Text = wfpagoprov
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT * from ef2forpag where f2forpag='" & txtcodforma.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                pnlnomforma.Caption = Trim("" & rst.Fields("F2DESPAG"))
            End If
            rst.Close
        End If
        Txt_Prove_KeyPress 13
    End If
    
End Sub

Private Sub Txt_Prove_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub Txt_Referencia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
        SendKeys "{tab}"
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
        SendKeys "{tab}"
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
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub imprimir()
    
    LLENA_TEMPCAB
    acr_ocompra.Show 1

End Sub

Private Sub eliminar()
Dim gcodigo     As String
Dim gcant       As Double
    
    If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
    rsOrdenCab.Open "SELECT * from if4orden where f4numord='" & Txt_NumOC & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsOrdenCab.EOF Then
        If MsgBox("¿Desea Anular La Orden de Compra?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Logística") = 6 Then
            
            sql = "Update if4ORDEN set f4estado=5,f4estnul='S',f4monto=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            sql = "Update if3ORDEN set f3canpro=0,f3igv=0,f3preuni=0,f5valvta=0,f3precos=0,f3total=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            AlmacenaQuery_sql sql, cnn_dbbancos
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
'                        AlmacenaQuery_sql sql, cnn_dbbancos
'                    Else
'                        sql = "update tb_cabsolicitud set cs_estado='2' where cod_solicitud='" & Txt_NumSolComp & "'"
'                        cnn_dbbancos.Execute sql
'                        AlmacenaQuery_sql sql, cnn_dbbancos
'                    End If
'                    rsOrdenDet.Close
                    MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " ha sido Anulada", vbInformation, App.Title
'                    Call Visi
'                    Call Limpia_Orden
'                    sw_nuevo_documento = False
'                    AdicionaItemGrid
'
'                    sw_nuevo_documento = True
'                    Call Limpiar
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
            
            AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            sql = "Update if3ORDEN set f3canpro=0,f3igv=0,f3preuni=0,f5valvta=0,f3precos=0,f3total=0 where F4NUMORD='" & Txt_NumOC.Text & "' and f4local='" & TOC & "'"
            cnn_dbbancos.Execute sql
            
            AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
            
            sql = "update tb_cabsolicitud set cs_estado='2', cs_orden = '', numorden = '' where cod_solicitud='" & Txt_NumSolComp & "'"
            cnn_dbbancos.Execute sql
            AlmacenaQuery_sql sql, cnn_dbbancos
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
''                        AlmacenaQuery_sql sql, cnn_dbbancos
''                    Else
''                        sql = "update tb_cabsolicitud set cs_estado='2' where cod_solicitud='" & Txt_NumSolComp & "'"
''                        cnn_dbbancos.Execute sql
''                        AlmacenaQuery_sql sql, cnn_dbbancos
''                    End If
''                    rsOrdenDet.Close
'                   ' MsgBox "La Orden de Compra Nº " & Txt_NumOC.Text & " ha sido Anulada", vbInformation, App.Title
''                    Call Visi
''                    Call Limpia_Orden
''                    sw_nuevo_documento = False
''                    AdicionaItemGrid
''
''                    sw_nuevo_documento = True
''                    Call Limpiar
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
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
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
            sql = "SELECT B.F5CODPRO,B.F5TEXTO,B.F5NOMPRO,B.F5AFECTO,B.F5CODFAB,B.F5VALVTA,B.F7CODMED,B.F5MARCA FROM EF2PROD_PROV AS A,IF5PLA AS B WHERE A.F2CODPRV='" & wrucprov & "' AND B.F5CODFAB='" & dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value & "' ORDER BY B.F5CODPRO"
            rsproduc.Open sql, cnn_dbbancos, adOpenStatic, adLockReadOnly
            If Not rsproduc.EOF Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = rsproduc.Fields("F5CODPRO") & ""
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value = rsproduc.Fields("F5CODFAB") & ""
                If Len(Trim(rsproduc.Fields("F5TEXTO")) & "") > 0 Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = rsproduc.Fields("F5TEXTO") & ""
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = rsproduc.Fields("F5NOMPRO") & ""
                End If
                dxDBGrid1.Columns.ColumnByFieldName("ds_unidmed").Value = rsproduc.Fields("F7CODMED") & ""
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & rsproduc.Fields("F5MARCA") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsproduc.Fields("F5MARCA") & ""
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = rsproduc.Fields("F5AFECTO") & ""
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Val(rsproduc.Fields("F5VALVTA") & "")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format$(abofechaentrega.Value, "DD/MM/YYYY"))
                dxDBGrid1.Dataset.FieldValues("check") = True
                dxDBGrid1.Dataset.Post
            End If
            rsproduc.Close
            Set rsproduc = Nothing
            dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("check").ColIndex - 1
        Case "f3preuni"
                Dim cantidad As Double
                Dim totdcto As Double
                Dim ValVta As Double
                Dim IGV  As Double
                With dxDBGrid1
                    cantidad = Val(Format(dxDBGrid1.Columns.ColumnByFieldName("F3CANPRO").Value, "0.00"))
                    If cantidad > 0 Then
                        .Dataset.Edit
                                totdcto = 0
                                
                                ValVta = Val(Format(cantidad * Val(Format("" & .Columns.ColumnByFieldName("F3PREuni").Value, "0.0000")) - totdcto, "0.00"))
                                
                                .Columns.ColumnByFieldName("F5VALVTA").Value = Format$(ValVta, "###,##0.00")
                                IGV = 0
                                .Columns.ColumnByFieldName("F3precos").Value = Val(Format("" & .Columns.ColumnByFieldName("F3PREuni").Value, "0.0000"))
                                If .Columns.ColumnByFieldName("F5AFECTO").Value = "*" Then     'Afecto
                                    
                                    .Columns.ColumnByFieldName("F3precos").Value = Val(Format("" & .Columns.ColumnByFieldName("F3PREuni").Value, "0.0000")) / (1 + (wwigv / 100))
                                    IGV = (.Columns.ColumnByFieldName("F3precos").Value * (wwigv / 100)) * cantidad
                                End If
                                .Columns.ColumnByFieldName("F3IGV").Value = Format$(IGV, "#,##0.00")
                                .Columns.ColumnByFieldName("F3TOTAL").Value = Format$(ValVta, "###,##0.00")
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
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F3CODPRO").Value & "")) = 0 Then
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
    dxDBGrid1.Dataset.ADODataset.CommandText = "select * from tmpOrdendeCompra"

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
            .FieldValues("f3fentrega") = Format$(abofechaentrega.Value, "dd/mm/yyyy")
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
    Grid.Dataset.ADODataset.CommandText = "select * from tmpOrdendeCompra"
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
            ayuda_productos.Show 1
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("f3codpro").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("f3medida").Value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value = wcodfab
                
                If rsmarcas.State = adStateOpen Then rsmarcas.Close
                rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsmarcas.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = wmarca
                End If
                dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value = wafecto

                Rem EMB dxDBGrid1.Dataset.FieldValues("f5valvta") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3PRECOS") = Format(wvv_prod, "###,##0.00")
                dxDBGrid1.Dataset.FieldValues("f3fentrega") = CVDate(Format(abofechaentrega.Value, "dd/mm/yyyy"))
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
                dxDBGrid1.Columns.ColumnByFieldName("f3cencos").Value = wcodcosto
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
            CalculaTotal
            sw_nuevo_item = False
        End If
    End If


End Sub

Private Sub txtplazo_entrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{tab}"
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
    SendKeys "{TAB}"
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
    Txt_TOC.Text = TOC
    Txt_NumOC.Text = GOC
    With rsOrdenCab
        If rsOrdenCab.State = adStateOpen Then rsOrdenCab.Close
        sql = "SELECT * from if4orden where f4numord='" & GOC & "' AND f4LOCAL = '" & TOC & "'"
        rsOrdenCab.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not (.EOF) Then
'            MostrarDatosOC
            MostrarDatosOC_Grid
            ExisteOrdenCompra = True
        Else
            ExisteOrdenCompra = False
        End If
        .Close
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
        SendKeys "{tab}"
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
                cnn_form.Execute (csql)
                AlmacenaQuery_sql sql, cnn_form
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
    
    amovs_cab(0).campo = "VIA_INGR": amovs_cab(0).valor = "1": amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "CORRELA": amovs_cab(1).valor = ncorre_d: amovs_cab(1).TIPO = "N"
    amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = cnro_comp: amovs_cab(2).TIPO = "T"
    amovs_cab(3).campo = "FCH_COMP": amovs_cab(3).valor = dfechamov: amovs_cab(3).TIPO = "F"
    amovs_cab(4).campo = "PROVEEDORO": amovs_cab(4).valor = ccodprov: amovs_cab(4).TIPO = "T"
    amovs_cab(5).campo = "RUC": amovs_cab(5).valor = cruc: amovs_cab(5).TIPO = "T"
    amovs_cab(6).campo = "MONEDAO": amovs_cab(6).valor = Moneda: amovs_cab(6).TIPO = "T"
    amovs_cab(7).campo = "TOTALO": amovs_cab(7).valor = ntotal: amovs_cab(7).TIPO = "N"
    amovs_cab(8).campo = "TCAMBIOO": amovs_cab(8).valor = ntc: amovs_cab(8).TIPO = "N"
    amovs_cab(9).campo = "PROVEEDOR": amovs_cab(9).valor = ccodprov: amovs_cab(9).TIPO = "T"
    amovs_cab(10).campo = "MONEDA": amovs_cab(10).valor = Moneda: amovs_cab(10).TIPO = "T"
    amovs_cab(11).campo = "TCAMBIO": amovs_cab(11).valor = ntc: amovs_cab(11).TIPO = "N"
    amovs_cab(12).campo = "TOTAL": amovs_cab(12).valor = ntotal: amovs_cab(12).TIPO = "N"
    amovs_cab(13).campo = "SALDO": amovs_cab(13).valor = ntotal: amovs_cab(13).TIPO = "N"
    amovs_cab(14).campo = "DEB_HAB": amovs_cab(14).valor = "H": amovs_cab(14).TIPO = "T"
    amovs_cab(15).campo = "REFERENCIA": amovs_cab(15).valor = cdetal: amovs_cab(15).TIPO = "T"
    amovs_cab(16).campo = "NOMPROV": amovs_cab(16).valor = cnomprov: amovs_cab(16).TIPO = "T"
    amovs_cab(17).campo = "CONCEPTO": amovs_cab(17).valor = cdetal: amovs_cab(17).TIPO = "T"
    amovs_cab(18).campo = "FCH_VCTO": amovs_cab(18).valor = dfechamov: amovs_cab(18).TIPO = "F"
    
    GRABA_REGISTRO_logistica amovs_cab(), "PAG_DCTO", "A", 18, cnn_ctaspag, ""
    
    cnn_ctaspag.Close
    sql = ("UPDATE IF4ORDEN SET F4CORRELA=" & ncorre_d & " WHERE F4NUMORD=" & pnumero & "")
    cnn_dbbancos.Execute sql
    AlmacenaQuery_sql sql, cnn_dbbancos
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
                editTo = Rs!mail
            Else
                editTo = editTo & ";" & Rs!mail
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
            StrDetLine = Format(Det!ITEM & "", "000") & ".- (" & Val(Det!F3CANPRO & "") & " " & left(Det!f7sigmed & Space(3), 3) & ") "
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
            If RsV!vbjefecc = True Then
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



