VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F7E69521-3C28-11D2-B3E7-00AA00B42B7C}#3.1#0"; "fpTab30.ocx"
Begin VB.Form parametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros del Sistema"
   ClientHeight    =   8490
   ClientLeft      =   1650
   ClientTop       =   1395
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   10980
   Begin TabproADOLib.fpTabProADO fpTabProADO1 
      Height          =   7035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10590
      _Version        =   196609
      _ExtentX        =   18680
      _ExtentY        =   12409
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCount        =   8
      ApplyTo         =   2
      OffsetFromClientTop=   -1  'True
      BookRingShowHole=   -1  'True
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      DataField       =   ""
      DataMember      =   ""
      TabCaption      =   "parametros.frx":0000
      PageEarMarkPictureNext=   "parametros.frx":0494
      PageEarMarkPicturePrev=   "parametros.frx":04B0
      EarMarkPictureNext=   "parametros.frx":04CC
      EarMarkPicturePrev=   "parametros.frx":04E8
      Begin VB.Frame Frame13 
         Enabled         =   0   'False
         Height          =   2715
         Left            =   -25019
         TabIndex        =   107
         Top             =   -18914
         Visible         =   0   'False
         Width           =   9420
         Begin Threed.SSCheck chkimpreso 
            Height          =   285
            Left            =   480
            TabIndex        =   108
            Top             =   720
            Width           =   5460
            _Version        =   65536
            _ExtentX        =   9631
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Impresión de Documentos para Puertos de Impresora definidos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkelimina 
            Height          =   285
            Left            =   480
            TabIndex        =   109
            Top             =   1080
            Width           =   7140
            _Version        =   65536
            _ExtentX        =   12594
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Eliminación de Item's al momento de modificar el Código de Producto ó la Cantidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkfacturar 
            Height          =   285
            Left            =   480
            TabIndex        =   110
            Top             =   1440
            Width           =   5940
            _Version        =   65536
            _ExtentX        =   10477
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Facturación Diaria y Separación de Documentos "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkcontrol 
            Height          =   285
            Left            =   480
            TabIndex        =   111
            Top             =   1800
            Width           =   5940
            _Version        =   65536
            _ExtentX        =   10477
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Control del Menu para los diferentes módulos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkmantenimiento 
            Height          =   285
            Left            =   480
            TabIndex        =   112
            Top             =   2160
            Width           =   5940
            _Version        =   65536
            _ExtentX        =   10477
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Mantenimiento de productos con alteración del Valor de Venta"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.ComboBox cmbmodo 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "parametros.frx":0504
         Left            =   -19814
         List            =   "parametros.frx":050E
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   -19529
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame7 
         Enabled         =   0   'False
         Height          =   4470
         Left            =   -25184
         TabIndex        =   67
         Top             =   -20624
         Visible         =   0   'False
         Width           =   9945
         Begin VB.Frame Frame11 
            Caption         =   " Cuentas  Contables "
            Height          =   1890
            Left            =   225
            TabIndex        =   69
            Top             =   555
            Width           =   9510
            Begin VB.TextBox txtctadescuentos 
               Height          =   330
               Left            =   1665
               MaxLength       =   12
               TabIndex        =   76
               Top             =   1215
               Width           =   1275
            End
            Begin VB.TextBox txtctaredondeomenos 
               Height          =   330
               Left            =   7935
               MaxLength       =   12
               TabIndex        =   75
               Top             =   810
               Width           =   1275
            End
            Begin VB.TextBox txtctaredondeomas 
               Height          =   330
               Left            =   4620
               MaxLength       =   12
               TabIndex        =   74
               Top             =   810
               Width           =   1275
            End
            Begin VB.TextBox txtctaotros_imp 
               Height          =   330
               Left            =   1665
               MaxLength       =   12
               TabIndex        =   73
               Top             =   810
               Width           =   1275
            End
            Begin VB.TextBox txtcta4ta 
               Height          =   330
               Left            =   7935
               MaxLength       =   12
               TabIndex        =   72
               Top             =   420
               Width           =   1275
            End
            Begin VB.TextBox txtctaies 
               Height          =   330
               Left            =   4620
               MaxLength       =   12
               TabIndex        =   71
               Top             =   420
               Width           =   1275
            End
            Begin VB.TextBox txtctaigv 
               Height          =   330
               Left            =   1665
               MaxLength       =   12
               TabIndex        =   70
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Descuentos"
               Height          =   210
               Left            =   270
               TabIndex        =   83
               Top             =   1290
               Width           =   870
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Redondeo (-)"
               Height          =   210
               Left            =   6540
               TabIndex        =   82
               Top             =   885
               Width           =   960
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Redondeo (+)"
               Height          =   210
               Left            =   3390
               TabIndex        =   81
               Top             =   885
               Width           =   990
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Otros Impuestos"
               Height          =   210
               Left            =   270
               TabIndex        =   80
               Top             =   885
               Width           =   1185
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Retención de 4ta."
               Height          =   210
               Left            =   6540
               TabIndex        =   79
               Top             =   495
               Width           =   1260
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "I.E.S."
               Height          =   210
               Left            =   3435
               TabIndex        =   78
               Top             =   495
               Width           =   360
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V."
               Height          =   210
               Left            =   270
               TabIndex        =   77
               Top             =   495
               Width           =   405
            End
         End
         Begin VB.TextBox txtorigen 
            Height          =   330
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   68
            Top             =   2760
            Width           =   510
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Origen"
            Height          =   210
            Left            =   540
            TabIndex        =   84
            Top             =   2835
            Width           =   480
         End
      End
      Begin VB.Frame Frame10 
         Enabled         =   0   'False
         Height          =   2715
         Left            =   -24824
         TabIndex        =   61
         Top             =   -19019
         Visible         =   0   'False
         Width           =   9420
         Begin Threed.SSCheck chkuupp 
            Height          =   285
            Left            =   450
            TabIndex        =   62
            Top             =   585
            Width           =   2220
            _Version        =   65536
            _ExtentX        =   3916
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Ingresar UUPP"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkccosto 
            Height          =   285
            Left            =   450
            TabIndex        =   63
            Top             =   945
            Width           =   2220
            _Version        =   65536
            _ExtentX        =   3916
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Ingresar Centro de Costo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   6300
         Left            =   -25184
         TabIndex        =   44
         Top             =   -21914
         Width           =   9870
         Begin VB.TextBox txtobsgen_oci 
            Height          =   450
            Left            =   3165
            TabIndex        =   99
            Top             =   5760
            Width           =   6405
         End
         Begin VB.TextBox txtemitido_oci 
            Height          =   450
            Left            =   3165
            TabIndex        =   97
            Top             =   5280
            Width           =   6405
         End
         Begin VB.TextBox txtnota_oci 
            Height          =   330
            Left            =   3165
            MultiLine       =   -1  'True
            TabIndex        =   95
            Top             =   4920
            Width           =   6405
         End
         Begin VB.TextBox txtobsfec_oci 
            Height          =   450
            Left            =   3165
            TabIndex        =   93
            Top             =   4440
            Width           =   6405
         End
         Begin VB.TextBox txtobsgen_oc 
            Height          =   450
            Left            =   3165
            TabIndex        =   91
            Top             =   3960
            Width           =   6405
         End
         Begin VB.TextBox txtemitido_oc 
            Height          =   450
            Left            =   3165
            TabIndex        =   89
            Top             =   3480
            Width           =   6405
         End
         Begin VB.TextBox txtnota_oc 
            Height          =   330
            Left            =   3165
            TabIndex        =   87
            Top             =   3120
            Width           =   6405
         End
         Begin VB.TextBox txtobsfec_oc 
            Height          =   450
            Left            =   3165
            TabIndex        =   85
            Top             =   2640
            Width           =   6405
         End
         Begin VB.TextBox txttexto_oc 
            Height          =   570
            Left            =   3120
            MaxLength       =   150
            TabIndex        =   48
            Top             =   1470
            Width           =   6405
         End
         Begin VB.TextBox txtasunto_oc 
            Height          =   330
            Left            =   3150
            MaxLength       =   50
            TabIndex        =   47
            Top             =   1065
            Width           =   6405
         End
         Begin VB.TextBox txtemailcc_oc 
            Height          =   330
            Left            =   3150
            MaxLength       =   30
            TabIndex        =   46
            Top             =   660
            Width           =   6405
         End
         Begin VB.TextBox txtemail_oc 
            Height          =   330
            Left            =   3150
            MaxLength       =   30
            TabIndex        =   45
            Top             =   255
            Width           =   6405
         End
         Begin Threed.SSCheck chkvisualiza_dsctos 
            Height          =   285
            Left            =   195
            TabIndex        =   66
            Top             =   2160
            Width           =   2190
            _Version        =   65536
            _ExtentX        =   3863
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Visualiza descuentos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Observación Gen. de O.C. Internacional"
            Height          =   210
            Left            =   240
            TabIndex        =   100
            Top             =   5880
            Width           =   2865
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Datos a emitir a de O.C. Internacional"
            Height          =   210
            Left            =   240
            TabIndex        =   98
            Top             =   5370
            Width           =   2655
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Nota de la O.C. Internacional"
            Height          =   210
            Left            =   240
            TabIndex        =   96
            Top             =   4920
            Width           =   2040
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Observación de Fecha O.C.Internacional"
            Height          =   210
            Left            =   240
            TabIndex        =   94
            Top             =   4530
            Width           =   2925
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Observación Gen. de O.C. Nacional"
            Height          =   210
            Left            =   240
            TabIndex        =   92
            Top             =   4080
            Width           =   2565
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Datos a emitir a de O.C. Nacional"
            Height          =   210
            Left            =   240
            TabIndex        =   90
            Top             =   3600
            Width           =   2355
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Nota de la O.C. Nacional"
            Height          =   210
            Left            =   240
            TabIndex        =   88
            Top             =   3210
            Width           =   1740
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones de Fecha O.C. Nacional"
            Height          =   210
            Left            =   240
            TabIndex        =   86
            Top             =   2730
            Width           =   2850
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Texto de la orden de compra"
            Height          =   210
            Left            =   225
            TabIndex        =   52
            Top             =   1560
            Width           =   2070
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Asunto de la orden de compra"
            Height          =   210
            Left            =   225
            TabIndex        =   51
            Top             =   1155
            Width           =   2190
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Email para enviar la o/c con copia"
            Height          =   210
            Left            =   225
            TabIndex        =   50
            Top             =   750
            Width           =   2415
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Email para enviar la orden de compra"
            Height          =   210
            Left            =   225
            TabIndex        =   49
            Top             =   345
            Width           =   2670
         End
      End
      Begin VB.Frame Frame8 
         Enabled         =   0   'False
         Height          =   4245
         Left            =   -25094
         TabIndex        =   39
         Top             =   -21359
         Width           =   9780
         Begin VB.Frame Frame9 
            Caption         =   " Importar datos para el Registro de Compras "
            Height          =   870
            Left            =   270
            TabIndex        =   57
            Top             =   1260
            Width           =   9105
            Begin Threed.SSOption optimporta 
               Height          =   240
               Index           =   0
               Left            =   495
               TabIndex        =   58
               Top             =   450
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "Orden de Compra"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optimporta 
               Height          =   240
               Index           =   1
               Left            =   3960
               TabIndex        =   59
               Top             =   450
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "Vale de Ingreso"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optimporta 
               Height          =   240
               Index           =   2
               Left            =   7290
               TabIndex        =   60
               Top             =   405
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "Ninguno"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
         End
         Begin VB.TextBox txtformato_rc 
            Height          =   330
            Left            =   2880
            TabIndex        =   42
            Top             =   765
            Width           =   555
         End
         Begin VB.TextBox txtformato_voucher 
            Height          =   330
            Left            =   2880
            TabIndex        =   40
            Top             =   360
            Width           =   555
         End
         Begin Threed.SSCheck chkctaspag 
            Height          =   285
            Left            =   270
            TabIndex        =   53
            Top             =   2340
            Width           =   4380
            _Version        =   65536
            _ExtentX        =   7726
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Interconexión con el módulo de Cuentas por Pagar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkshowcodigo 
            Height          =   285
            Left            =   270
            TabIndex        =   54
            Top             =   2700
            Width           =   4740
            _Version        =   65536
            _ExtentX        =   8361
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " No mostrar el código de la cuenta contable del proveedor"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chknumera 
            Height          =   285
            Left            =   270
            TabIndex        =   55
            Top             =   3780
            Visible         =   0   'False
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Numeración de Comercial Alimenticia"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkbancos 
            Height          =   285
            Left            =   270
            TabIndex        =   56
            Top             =   3180
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   " Interconexión con el módulo de Bancos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Formato del Registro de Compras"
            Height          =   210
            Left            =   315
            TabIndex        =   43
            Top             =   855
            Width           =   2400
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Formato del Voucher"
            Height          =   210
            Left            =   315
            TabIndex        =   41
            Top             =   450
            Width           =   1515
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Porcentajes "
         Enabled         =   0   'False
         Height          =   870
         Left            =   -25049
         TabIndex        =   32
         Top             =   -16949
         Width           =   9735
         Begin VB.TextBox txt4ta 
            Height          =   330
            Left            =   8820
            TabIndex        =   38
            Top             =   315
            Width           =   555
         End
         Begin VB.TextBox txties 
            Height          =   330
            Left            =   5085
            TabIndex        =   36
            Top             =   315
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtigv 
            Height          =   330
            Left            =   1260
            TabIndex        =   34
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Retención de 4ta."
            Height          =   210
            Left            =   7425
            TabIndex        =   37
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "I.E.S."
            Height          =   210
            Left            =   4545
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "I.G.V."
            Height          =   210
            Left            =   720
            TabIndex        =   33
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   2670
         Left            =   -25229
         TabIndex        =   23
         Top             =   -19064
         Width           =   9870
         Begin VB.TextBox txttexto_sol 
            Height          =   690
            Left            =   3150
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1710
            Width           =   6405
         End
         Begin VB.TextBox txtasunto_sol 
            Height          =   330
            Left            =   3150
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1305
            Width           =   6405
         End
         Begin VB.TextBox txtemailcc_sol 
            Height          =   330
            Left            =   3150
            MaxLength       =   30
            TabIndex        =   27
            Top             =   900
            Width           =   6405
         End
         Begin VB.TextBox txtemail_sol 
            Height          =   330
            Left            =   3150
            MaxLength       =   30
            TabIndex        =   25
            Top             =   495
            Width           =   6405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Texto de la solicitud"
            Height          =   210
            Left            =   225
            TabIndex        =   30
            Top             =   1800
            Width           =   1425
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Asunto de la solicitud"
            Height          =   210
            Left            =   225
            TabIndex        =   28
            Top             =   1395
            Width           =   1545
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Email para enviar la solicitud con copia"
            Height          =   210
            Left            =   225
            TabIndex        =   26
            Top             =   990
            Width           =   2775
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Email para enviar la solicitud"
            Height          =   210
            Left            =   225
            TabIndex        =   24
            Top             =   585
            Width           =   2025
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Conceptos "
         Enabled         =   0   'False
         Height          =   5460
         Left            =   -25229
         TabIndex        =   10
         Top             =   -21494
         Width           =   10005
         Begin VB.Frame Frame14 
            Height          =   3300
            Left            =   4905
            TabIndex        =   113
            Top             =   1080
            Width           =   4560
            Begin VB.Frame Frame15 
               Caption         =   "Descripción y longitud de campo"
               Height          =   2130
               Left            =   225
               TabIndex        =   118
               Top             =   675
               Width           =   4155
               Begin VB.TextBox txtnivel1 
                  Height          =   315
                  Index           =   4
                  Left            =   3465
                  MaxLength       =   1
                  TabIndex        =   132
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox txtdescri1 
                  Height          =   315
                  Index           =   4
                  Left            =   720
                  TabIndex        =   131
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   2670
               End
               Begin VB.TextBox txtnivel1 
                  Height          =   315
                  Index           =   3
                  Left            =   3465
                  MaxLength       =   1
                  TabIndex        =   129
                  Top             =   1305
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox txtdescri1 
                  Height          =   315
                  Index           =   3
                  Left            =   720
                  TabIndex        =   128
                  Top             =   1305
                  Visible         =   0   'False
                  Width           =   2670
               End
               Begin VB.TextBox txtnivel1 
                  Height          =   315
                  Index           =   2
                  Left            =   3465
                  MaxLength       =   1
                  TabIndex        =   126
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox txtdescri1 
                  Height          =   315
                  Index           =   2
                  Left            =   720
                  TabIndex        =   125
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   2670
               End
               Begin VB.TextBox txtnivel1 
                  Height          =   315
                  Index           =   1
                  Left            =   3465
                  MaxLength       =   1
                  TabIndex        =   123
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox txtdescri1 
                  Height          =   315
                  Index           =   1
                  Left            =   720
                  TabIndex        =   122
                  Top             =   585
                  Visible         =   0   'False
                  Width           =   2670
               End
               Begin VB.TextBox txtdescri1 
                  Height          =   315
                  Index           =   0
                  Left            =   720
                  TabIndex        =   121
                  Top             =   225
                  Visible         =   0   'False
                  Width           =   2670
               End
               Begin VB.TextBox txtnivel1 
                  Height          =   315
                  Index           =   0
                  Left            =   3465
                  MaxLength       =   1
                  TabIndex        =   119
                  Top             =   225
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel 5"
                  Height          =   210
                  Index           =   4
                  Left            =   180
                  TabIndex        =   133
                  Top             =   1710
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel 4"
                  Height          =   210
                  Index           =   3
                  Left            =   180
                  TabIndex        =   130
                  Top             =   1350
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel 3"
                  Height          =   210
                  Index           =   2
                  Left            =   180
                  TabIndex        =   127
                  Top             =   990
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel 2"
                  Height          =   210
                  Index           =   1
                  Left            =   180
                  TabIndex        =   124
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Nivel 1"
                  Height          =   210
                  Index           =   0
                  Left            =   180
                  TabIndex        =   120
                  Top             =   270
                  Visible         =   0   'False
                  Width           =   480
               End
            End
            Begin VB.TextBox txtniveles 
               Height          =   315
               Left            =   1710
               MaxLength       =   1
               TabIndex        =   115
               Top             =   270
               Width           =   555
            End
            Begin VB.TextBox txtcodprod 
               Height          =   315
               Left            =   2385
               MaxLength       =   4
               TabIndex        =   114
               Top             =   2880
               Width           =   555
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Niveles a usar"
               Height          =   210
               Left            =   540
               TabIndex        =   117
               Top             =   315
               Width           =   1035
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "Longitud campo de producto"
               Height          =   210
               Left            =   225
               TabIndex        =   116
               Top             =   2925
               Width           =   2055
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Columna a Visualizar en los Vales"
            Height          =   870
            Left            =   360
            TabIndex        =   101
            Top             =   4380
            Width           =   9105
            Begin Threed.SSOption optcolumna 
               Height          =   240
               Index           =   0
               Left            =   495
               TabIndex        =   102
               Top             =   450
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "Código Interno"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
            Begin Threed.SSOption optcolumna 
               Height          =   240
               Index           =   1
               Left            =   3960
               TabIndex        =   103
               Top             =   480
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "Código de Fabricante"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optcolumna 
               Height          =   240
               Index           =   2
               Left            =   7290
               TabIndex        =   104
               Top             =   405
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "Ambos"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.ComboBox cmbmoneda 
            Height          =   330
            ItemData        =   "parametros.frx":0522
            Left            =   3495
            List            =   "parametros.frx":052C
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   3510
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Height          =   1905
            Left            =   315
            TabIndex        =   14
            Top             =   1080
            Width           =   4110
            Begin VB.TextBox txtsalidaalm 
               Height          =   315
               Left            =   2610
               MaxLength       =   4
               TabIndex        =   19
               Top             =   1215
               Width           =   1140
            End
            Begin VB.TextBox txtsalidaxtransf 
               Height          =   315
               Left            =   2610
               MaxLength       =   11
               TabIndex        =   16
               Top             =   405
               Width           =   1140
            End
            Begin VB.TextBox txtingobra 
               Height          =   315
               Left            =   2610
               MaxLength       =   4
               TabIndex        =   15
               Top             =   810
               Width           =   1140
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Salida del almacén"
               Height          =   210
               Left            =   270
               TabIndex        =   20
               Top             =   1260
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Salida por Transferencia"
               Height          =   210
               Left            =   270
               TabIndex        =   18
               Top             =   450
               Width           =   1785
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Ingreso a la obra"
               Height          =   210
               Left            =   270
               TabIndex        =   17
               Top             =   855
               Width           =   1215
            End
         End
         Begin Threed.SSCheck chktiposalida 
            Height          =   420
            Left            =   315
            TabIndex        =   13
            Top             =   720
            Visible         =   0   'False
            Width           =   9420
            _Version        =   65536
            _ExtentX        =   16616
            _ExtentY        =   741
            _StockProps     =   78
            Caption         =   " Empresa Constructora, genera una salida por cada ingreso y si el centro de costo tiene un almacen genera una transferencia"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtingalmxoc 
            Height          =   330
            Left            =   3510
            MaxLength       =   100
            TabIndex        =   11
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda para generar los vales de salida"
            Height          =   210
            Left            =   360
            TabIndex        =   21
            Top             =   3555
            Width           =   2940
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso al almacén por Orden de Compra"
            Height          =   210
            Left            =   315
            TabIndex        =   12
            Top             =   360
            Width           =   2955
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3165
         Left            =   360
         TabIndex        =   6
         Top             =   1215
         Width           =   9735
         Begin VB.TextBox txtdireccion2 
            Height          =   330
            Left            =   2250
            MaxLength       =   100
            TabIndex        =   5
            Top             =   2115
            Width           =   7260
         End
         Begin VB.TextBox txtdireccion1 
            Height          =   330
            Left            =   2250
            MaxLength       =   100
            TabIndex        =   4
            Top             =   1665
            Width           =   7260
         End
         Begin VB.TextBox txtnomempresa 
            Height          =   330
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   1
            Top             =   405
            Width           =   7260
         End
         Begin VB.TextBox txtanno 
            Height          =   315
            Left            =   2250
            MaxLength       =   4
            TabIndex        =   3
            Top             =   1260
            Width           =   1140
         End
         Begin VB.TextBox txtruc 
            Height          =   315
            Left            =   2250
            MaxLength       =   11
            TabIndex        =   2
            Top             =   855
            Width           =   1140
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Dirección 2"
            Height          =   210
            Left            =   315
            TabIndex        =   65
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Dirección 1"
            Height          =   210
            Left            =   315
            TabIndex        =   64
            Top             =   1710
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de la Empresa"
            Height          =   210
            Left            =   315
            TabIndex        =   9
            Top             =   450
            Width           =   1620
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año de proceso"
            Height          =   210
            Left            =   315
            TabIndex        =   8
            Top             =   1305
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "R.U.C."
            Height          =   210
            Left            =   315
            TabIndex        =   7
            Top             =   900
            Width           =   450
         End
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Facturación de ..."
         Enabled         =   0   'False
         Height          =   210
         Left            =   -17384
         TabIndex        =   106
         Top             =   -19424
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   405
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "parametros.frx":0540
      ToolBars        =   "parametros.frx":511C
   End
End
Attribute VB_Name = "parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmesproame      As String
Dim cnn_cont As New ADODB.Connection
Private Sub Form_Load()
    Me.MousePointer = 11
    cmbmodo.Clear
    cmbmodo.AddItem "Tienda "
    cmbmodo.AddItem "Proyectos "
    cmbmodo.ListIndex = 0
    Me.left = 900
    Me.top = 650

    VISUALIZA_PARAMETROS
    Me.MousePointer = 1
End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Grabar":
            GRABAR_PARAMETROS
            MsgBox "Los parámetros fueron actualizados", vbInformation, "Sistema de Logística"
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub

Private Sub GRABAR_PARAMETROS()
Dim I               As Integer
Dim csql            As String
Dim ctiposalida     As String
Dim cmoneda         As String
Dim coc             As String
Dim ctrasoc         As String
Dim cbancos         As String
Dim cctaspag        As String
Dim cnumera         As String
Dim cshowcodigo     As String
Dim cuupp           As String
Dim cproame         As String
Dim ccosto          As String
Dim cdirec1         As String
Dim cdirec2         As String
Dim cvisualiza_dsctos   As String
Dim cvisualiza_columna  As String

    Me.MousePointer = 11
    cmoneda = IIf(cmbmoneda.ListIndex = 0, "S", "D")
    ctiposalida = IIf(chktiposalida.Value = True, "*", " ")
    
    sql = ""
    For I = 1 To 5
        If txtnivel1(I - 1).Visible = True Then
            sql = sql + " ,f1nivel0" & I & "= '" & txtdescri1(I - 1).Text & "'" _
                & " ,f1lonniv" & I & "= " & txtnivel1(I - 1).Text
        End If
    Next
    
    csql = "UPDATE SF1PARAIN SET F1CONC_COMPRA='" & txtingalmxoc.Text & "',F1CONC_SALXTRANSF='" & _
           txtsalidaxtransf.Text & "',F1CONC_SALIDA='" & txtsalidaalm.Text & "',F1CONC_ING_OBRA='" & _
           txtingobra.Text & "',F1TIPOSALIDA='" & ctiposalida & "',F1MONEDA_PRODUCTOS='" & cmoneda & "'" & _
           ", f1numord = " & txtniveles.Text & ", f1loncod = " & txtcodprod.Text & _
           sql & " WHERE F1CODEMP ='" & wempresa & "'"
    cnn_cont.Execute (csql)
    AlmacenaQuery_sql csql, cnn_cont
    
    
    coc = "": ctrasoc = ""
    If optimporta(0).Value = True Then
        coc = "O"
        ctrasoc = "*"
    End If
    If optimporta(1).Value = True Then
        coc = "V"
        ctrasoc = "*"
    End If
    If optimporta(2).Value = True Then
        coc = " "
        ctrasoc = " "
    End If
    cbancos = IIf(chkbancos.Value = True, "*", " ")
    cctaspag = IIf(chkctaspag.Value = True, "*", " ")
    cnumera = IIf(chknumera.Value = True, "*", " ")
    cshowcodigo = IIf(chkshowcodigo.Value = True, "*", " ")
    cuupp = IIf(chkuupp.Value = True, "*", " ")
    cproame = txtanno.Text & cmesproame
    ccosto = IIf(chkccosto.Value = True, "S", "N")
    cdirec1 = Trim(txtdireccion1.Text)
    cdirec2 = Trim(txtdireccion2.Text)
    cvisualiza_dsctos = IIf(chkvisualiza_dsctos.Value = True, "*", " ")
    'If optcolumna(0).Value = True Then wvisualiza_cod = "I"
    'If optcolumna(1).Value = True Then wvisualiza_cod = "F"
    'If optcolumna(2).Value = True Then wvisualiza_cod = "T"
    
    csql = "UPDATE PARAM_COM SET F1NOMEMP='" & txtnomempresa.Text & "',F1RUCEMP='" & txtruc.Text & _
           "',F1EMAIL_SOLICITUD='" & txtemail_sol.Text & "',F1EMAIL_CCSOL='" & txtemailcc_sol.Text & _
           "',F1ASUNTO_SOL='" & txtasunto_sol.Text & "',F1TEXTO_SOL='" & txttexto_sol.Text & _
           "',F1EMAIL_OC='" & txtemail_oc.Text & "',F1EMAIL_CCOC='" & txtemailcc_oc.Text & _
           "',F1ASUNTO_OC='" & txtasunto_oc.Text & "',F1TEXTO_OC='" & txttexto_oc.Text & _
           "',F1TRASOC='" & ctrasoc & "',F1TIPDOC_ASOC='" & coc & "',F1BANCOS='" & cbancos & _
           "',F1CTAPAG='" & cctaspag & "',F1NUMERA='" & cnumera & "',F1VISCOD='" & cshowcodigo & _
           "',F1FORMATOV='" & txtformato_voucher.Text & "',F1IGV=" & txtigv.Text & ",F1FONAVI=" & txties.Text & _
           ",F1RETENC=" & txt4ta.Text & ",F1ORIGEN='" & txtorigen.Text & "',F1CTAIGV='" & txtctaigv.Text & _
           "',F1CTAOTROS='" & txtctaotros_imp.Text & "',F1REDSUMA='" & txtctaredondeomas.Text & _
           "',F1REDRESTA='" & txtctaredondeomenos.Text & "',F1CTARET='" & txtcta4ta.Text & _
           "',F1CTAFONAVI='" & txtctaies.Text & "',F1FORMATORC='" & txtformato_rc.Text & _
           "',F1DCTO='" & txtctadescuentos.Text & "',F1UUPP='" & cuupp & "',F1PROAME='" & cproame & _
           "',F1SHOW_CCOSTO='" & ccosto & "',F1DIREMP='" & cdirec1 & "',F1DIREMP2='" & cdirec2 & "',F1VISUALIZA_DCTOS='" & cvisualiza_dsctos & _
           "',F1OBSFECENT_OC='" & txtobsfec_oc.Text & "',F1NOTA_OC='" & txtnota_oc.Text & "',F1EMITIDO_OC='" & txtemitido_oc.Text & _
           "',F1OBSGEN_OC='" & txtobsgen_oc.Text & "',F1OBSFECENT_OCI='" & txtobsfec_oci.Text & "',F1NOTA_OCI='" & txtnota_oci.Text & "',F1EMITIDO_OCI='" & txtemitido_oci.Text & "',F1OBSGEN_OCI='" & txtobsgen_oci.Text & _
           "',F1VISUALIZA_COD='" & Trim(UCase("I")) & "'WHERE F1CODEMP ='" & wempresa & "'"
    
    cnn_ctrcom.Execute (csql)
    AlmacenaQuery_sql csql, cnn_ctrcom
    '----------------------------------------------------------
    If rsparam_com.State = adStateOpen Then rsparam_com.Close
    rsparam_com.Open "SELECT * FROM PARAM_COM WHERE F1CODEMP ='" & wempresa & "'", cnn_ctrcom
    If Not rsparam_com.EOF Then
        wnomcia = Trim(rsparam_com.Fields("F1NOMEMP") & "")
        wrucempresa = Trim(rsparam_com.Fields("F1RUCEMP") & "")
        wemailsol = Trim(rsparam_com.Fields("F1EMAIL_SOLICITUD") & "")
        wemailccsol = Trim(rsparam_com.Fields("F1EMAIL_CCSOL") & "")
        wasuntosol = Trim(rsparam_com.Fields("F1ASUNTO_SOL") & "")
        wtextosol = Trim(rsparam_com.Fields("F1TEXTO_SOL") & "")
        wemailoc = Trim(rsparam_com.Fields("F1EMAIL_OC") & "")
        wemailccoc = Trim(rsparam_com.Fields("F1EMAIL_CCOC") & "")
        wasuntooc = Trim(rsparam_com.Fields("F1ASUNTO_OC") & "")
        wtextooc = Trim(rsparam_com.Fields("F1TEXTO_OC") & "")
        
        ' Right(rsparam_com.Fields("f1proame") & "", 2)
        'wocompra = Trim(rsparam_com.Fields("f1ocompra") & "")  '---- MODIFICAR POR wf1trasoc
        'wf1tipdoc_asoc = rsparam_com.Fields("F1TIPDOC_ASOC") & ""
        'wf1inggasto = rsparam_com.Fields("f1inggasto") & ""
        'wbancos = Trim(rsparam_com.Fields("f1bancos") & "")
        'wf1traslado = rsparam_com.Fields("f1traslado") & ""
        'gctapag = "" & rsparam_com.Fields("F1CTAPAG")
        'wf1formatov = rsparam_com.Fields("f1formatov") & ""
        'wf1numera = rsparam_com.Fields("f1numera") & ""
        'wf1viscod = rsparam_com.Fields("f1viscod") & ""
        'wf1trasoc = rsparam_com.Fields("F1TRASOC") & ""
       '
        wIgv = Val("" & rsparam_com.Fields("F1IGV"))
       '
       ' gfonavi = Val("" & rsparam_com.Fields("F1FONAVI"))
        gretenc = Val("" & rsparam_com.Fields("F1RETENC"))        'wingobra = Trim(rsparam_com.Fields("f1ingobra") & "")
        worigen = Trim(rsparam_com.Fields("f1origen") & "")
        wctaigv = Trim(rsparam_com.Fields("f1ctaigv") & "")
        'wctaotros = Trim(rsparam_com.Fields("f1ctaotros") & "")
        'wredsuma = Trim(rsparam_com.Fields("f1redsuma") & "")
        'wredresta = Trim(rsparam_com.Fields("f1redresta") & "")
        'wctaret = Trim(rsparam_com.Fields("f1ctaret") & "")
        'wctafon = Trim(rsparam_com.Fields("f1ctafonavi") & "")
        'wf1formato = rsparam_com.Fields("f1formatorc") & ""
        'wanno = Left(rsparam_com.Fields("f1proame") & "", 4)
        'wdcto = Trim(rsparam_com.Fields("f1dcto") & "")
        'wf1cnting = rsparam_com.Fields("f1cnting") & ""
        'wf1uupp = Trim(rsparam_com.Fields("F1UUPP") & "")
        'wf1show_ccosto = Trim(rsparam_com.Fields("F1SHOW_CCOSTO") & "")
        'wf1direc1 = Trim(rsparam_com.Fields("F1DIREMP") & "")
        'wf1direc2 = Trim(rsparam_com.Fields("F1DIREMP2") & "")
        'cvisualiza_dsctos = Trim(rsparam_com.Fields("F1VISUALIZA_DCTOS") & "")
        
    End If
    rsparam_com.Close
    '----------------------------------------------------------
    Me.MousePointer = 1
End Sub

Private Sub VISUALIZA_PARAMETROS()
    Dim I As Integer
    
    
    If cnn_cont.State = 1 Then cnn_cont.Close
    cnn_cont.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CONTROL.MDB" & ";Persist Security Info=False"
    If rscontrol.State = adStateOpen Then rscontrol.Close
    rscontrol.Open "SELECT * FROM SF1PARAIN WHERE F1CODEMP ='" & wempresa & "'", cnn_cont, adOpenDynamic, adLockOptimistic
    If Not rscontrol.EOF Then
        txtingalmxoc.Text = rscontrol.Fields("F1CONC_COMPRA") & ""
        txtsalidaxtransf.Text = rscontrol.Fields("F1CONC_SALXTRANSF") & ""
        txtsalidaalm.Text = rscontrol.Fields("F1CONC_SALIDA") & ""
        txtingobra.Text = rscontrol.Fields("F1CONC_ING_OBRA") & ""
        chktiposalida.Value = IIf(Trim(rscontrol.Fields("F1TIPOSALIDA") & "") = "*", True, False)
        cmbmoneda.ListIndex = IIf(Trim(rscontrol.Fields("F1MONEDA_PRODUCTOS") & "") = "S", 0, 1)
        txtniveles.Text = rscontrol.Fields("f1niveles")
        For I = 1 To rscontrol.Fields("f1niveles")
            Label43(I - 1).Visible = True
            txtdescri1(I - 1).Visible = True
            txtnivel1(I - 1).Visible = True
            txtdescri1(I - 1).Text = rscontrol.Fields("f1nivel0" & I)
            txtnivel1(I - 1).Text = rscontrol.Fields("f1lonniv" & I)
        Next
        txtcodprod.Text = rscontrol.Fields("f1loncod")
    End If
    rscontrol.Close
    
    If rscontrol.State = adStateOpen Then rscontrol.Close
    rscontrol.Open "SELECT * FROM SF1PARAM WHERE F1CODEMP ='" & wempresa & "'", cnn_cont, adOpenDynamic, adLockOptimistic
    If Not rscontrol.EOF Then
        'chkimpreso.Value = rscontrol.Fields("F1IMPRESORAS") & ""
        'chkimpreso.Value = rscontrol.Fields("F1ELIM_ITEM_DCTO") & ""
        'chkimpreso.Value = rscontrol.Fields("F1FACTURAR_DIARIO_CORRELA") & ""
        'chkimpreso.Value = rscontrol.Fields("F1CONtrol_menu") & ""
        'chkimpreso.Value = IIf(Trim(rscontrol.Fields("F1mant_productos") & "") = "*", True, False)
        'cmbmodo.ListIndex = IIf(Trim(rscontrol.Fields("F1sistema_venta_proyectos") & "") = "V", 0, 1)
    End If
    rscontrol.Close
    
    If rsparam_com.State = adStateOpen Then rsparam_com.Close
    rsparam_com.Open "SELECT * FROM PARAM_COM WHERE F1CODEMP ='" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
    If Not rsparam_com.EOF Then
        txtnomempresa.Text = Trim(rsparam_com.Fields("F1NOMEMP") & "")
        txtruc.Text = Trim(rsparam_com.Fields("F1RUCEMP") & "")
        txtemail_sol.Text = Trim(rsparam_com.Fields("F1EMAIL_SOLICITUD") & "")
        txtemailcc_sol.Text = Trim(rsparam_com.Fields("F1EMAIL_CCSOL") & "")
        txtasunto_sol.Text = Trim(rsparam_com.Fields("F1ASUNTO_SOL") & "")
        txttexto_sol.Text = Trim(rsparam_com.Fields("F1TEXTO_SOL") & "")
        txtemail_oc.Text = Trim(rsparam_com.Fields("F1EMAIL_OC") & "")
        txtemailcc_oc.Text = Trim(rsparam_com.Fields("F1EMAIL_CCOC") & "")
        txtasunto_oc.Text = Trim(rsparam_com.Fields("F1ASUNTO_OC") & "")
        txttexto_oc.Text = Trim(rsparam_com.Fields("F1TEXTO_OC") & "")
        chkvisualiza_dsctos.Value = IIf(Trim(rsparam_com.Fields("F1VISUALIZA_DCTOS") & "") = "*", True, False)
        txtobsfec_oc.Text = Trim(rsparam_com.Fields("F1OBSFECENT_OC") & "")
        txtobsfec_oci.Text = Trim(rsparam_com.Fields("F1OBSFECENT_OCI") & "")
        txtnota_oc.Text = Trim(rsparam_com.Fields("F1NOTA_OC") & "")
        txtnota_oci.Text = Trim(rsparam_com.Fields("F1NOTA_OCI") & "")
        txtemitido_oc.Text = Trim(rsparam_com.Fields("F1EMITIDO_OC") & "")
        txtemitido_oci.Text = Trim(rsparam_com.Fields("F1EMITIDO_OCI") & "")
        txtobsgen_oc.Text = Trim(rsparam_com.Fields("F1OBSGEN_OC") & "")
        txtobsgen_oci.Text = Trim(rsparam_com.Fields("F1OBSGEN_OCI") & "")
        
        Select Case UCase(Trim(rsparam_com.Fields("F1VISUALIZA_COD") & ""))
            Case "I"
                optcolumna(0).Value = True
                optcolumna(1).Value = False
                optcolumna(2).Value = False
            Case "F"
                optcolumna(1).Value = True
                optcolumna(0).Value = False
                optcolumna(2).Value = False
            Case "T"
                optcolumna(2).Value = True
                optcolumna(1).Value = False
                optcolumna(0).Value = False
        End Select
        
        If rsparam_com.Fields("F1TRASOC") & "" = "*" Then
            If rsparam_com.Fields("F1TIPDOC_ASOC") & "" = "V" Then
                optimporta(1).Value = True
            Else
                optimporta(0).Value = True
            End If
        Else
            optimporta(2).Value = True
        End If
        
        chkbancos.Value = IIf(Trim(rsparam_com.Fields("f1bancos") & "") = "*", True, False)
        chkctaspag.Value = IIf("" & rsparam_com.Fields("F1CTAPAG") = "*", True, False)
        txtformato_voucher.Text = rsparam_com.Fields("f1formatov") & ""
        chknumera.Value = IIf(rsparam_com.Fields("f1numera") & "" = "*", True, False)
        chkshowcodigo.Value = IIf(rsparam_com.Fields("f1viscod") & "" = "*", True, False)
        
        
        txtigv.Text = Val("" & rsparam_com.Fields("F1IGV"))
        txties.Text = Val("" & rsparam_com.Fields("F1FONAVI"))
        txt4ta.Text = Val("" & rsparam_com.Fields("F1RETENC"))
        txtorigen.Text = Trim(rsparam_com.Fields("f1origen") & "")
        txtctaigv.Text = Trim(rsparam_com.Fields("f1ctaigv") & "")
        txtctaotros_imp.Text = Trim(rsparam_com.Fields("f1ctaotros") & "")
        txtctaredondeomas.Text = Trim(rsparam_com.Fields("f1redsuma") & "")
        txtctaredondeomenos.Text = Trim(rsparam_com.Fields("f1redresta") & "")
        txtcta4ta.Text = Trim(rsparam_com.Fields("f1ctaret") & "")
        txtctaies.Text = Trim(rsparam_com.Fields("f1ctafonavi") & "")
        txtformato_rc.Text = rsparam_com.Fields("f1formatorc") & ""
        txtanno.Text = left(rsparam_com.Fields("f1proame") & "", 4)
        txtctadescuentos.Text = Trim(rsparam_com.Fields("f1dcto") & "")
        chkuupp.Value = IIf(Trim(rsparam_com.Fields("F1UUPP") & "") = "*", True, False)
        cmesproame = Mid(rsparam_com.Fields("f1proame") & "", 5, 2)
        chkccosto.Value = IIf(Trim(rsparam_com.Fields("F1SHOW_CCOSTO") & "") = "N", False, True)
        txtdireccion1.Text = Trim(rsparam_com.Fields("F1DIREMP") & "")
        txtdireccion2.Text = Trim(rsparam_com.Fields("F1DIREMP2") & "")
        
    End If
    rsparam_com.Close
    
End Sub

Private Sub txt4ta_GotFocus()

    txt4ta.SelStart = 0: txt4ta.SelLength = Len(txt4ta.Text)

End Sub

Private Sub txtanno_GotFocus()

    txtanno.SelStart = 0: txtanno.SelLength = Len(txtanno.Text)
    
End Sub

Private Sub txtasunto_oc_GotFocus()

    txtasunto_oc.SelStart = 0: txtasunto_oc.SelLength = Len(txtasunto_oc.Text)
    
End Sub

Private Sub txtasunto_sol_GotFocus()

    txtasunto_sol.SelStart = 0: txtasunto_sol.SelLength = Len(txtasunto_sol.Text)
    
End Sub

Private Sub txtcodprod_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 49 To 53, 13:
        Case Else
        KeyAscii = 0
    End Select

End Sub

Private Sub txtcta4ta_GotFocus()

    txtcta4ta.SelStart = 0: txtcta4ta.SelLength = Len(txtcta4ta.Text)
    
End Sub

Private Sub txtctadescuentos_GotFocus()

    txtctadescuentos.SelStart = 0: txtctadescuentos.SelLength = Len(txtctadescuentos.Text)
    
End Sub

Private Sub txtctaies_GotFocus()

    txtctaies.SelStart = 0: txtctaies.SelLength = Len(txtctaies.Text)
    
End Sub

Private Sub txtctaigv_GotFocus()

    txtctaigv.SelStart = 0: txtctaigv.SelLength = Len(txtctaigv.Text)
    
End Sub

Private Sub txtctaotros_imp_GotFocus()

    txtctaotros_imp.SelStart = 0: txtctaotros_imp.SelLength = Len(txtctaotros_imp.Text)
    
End Sub

Private Sub txtctaredondeomas_GotFocus()

    txtctaredondeomas.SelStart = 0: txtctaredondeomas.SelLength = Len(txtctaredondeomas.Text)
    
End Sub

Private Sub txtctaredondeomenos_GotFocus()

    txtctaredondeomenos.SelStart = 0: txtctaredondeomenos.SelLength = Len(txtctaredondeomenos.Text)
    
End Sub

Private Sub txtemail_oc_GotFocus()

    txtemail_oc.SelStart = 0: txtemail_oc.SelLength = Len(txtemail_oc.Text)
    
End Sub

Private Sub txtemail_sol_GotFocus()

    txtemail_sol.SelStart = 0: txtemail_sol.SelLength = Len(txtemail_sol.Text)
    
End Sub

Private Sub txtemailcc_oc_GotFocus()

    txtemailcc_oc.SelStart = 0: txtemailcc_oc.SelLength = Len(txtemailcc_oc.Text)
    
End Sub

Private Sub txtemailcc_sol_GotFocus()

    txtemailcc_sol.SelStart = 0: txtemailcc_sol.SelLength = Len(txtemailcc_sol.Text)
    
End Sub

Private Sub txtemitido_oc_GotFocus()
    txtemitido_oc.SelStart = 0: txtemitido_oc.SelLength = Len(txtemitido_oc.Text)
End Sub

Private Sub txtemitido_oci_GotFocus()
    txtemitido_oci.SelStart = 0: txtemitido_oci.SelLength = Len(txtemitido_oci.Text)
End Sub

Private Sub txtformato_rc_GotFocus()

    txtformato_rc.SelStart = 0: txtformato_rc.SelLength = Len(txtformato_rc.Text)
    
End Sub

Private Sub txtformato_voucher_GotFocus()

    txtformato_voucher.SelStart = 0: txtformato_voucher.SelLength = Len(txtformato_voucher.Text)
    
End Sub

Private Sub txties_GotFocus()

    txties.SelStart = 0: txties.SelLength = Len(txties.Text)
    
End Sub

Private Sub TxtIgv_GotFocus()

    txtigv.SelStart = 0: txtigv.SelLength = Len(txtigv.Text)
    
End Sub

Private Sub txtingalmxoc_GotFocus()

    txtingalmxoc.SelStart = 0: txtingalmxoc.SelLength = Len(txtingalmxoc.Text)
    
End Sub

Private Sub txtingobra_GotFocus()

    txtingobra.SelStart = 0: txtingobra.SelLength = Len(txtingobra.Text)
    
End Sub

Private Sub txtniveles_KeyPress(KeyAscii As Integer)
Dim I As Integer
Dim J As Integer
    Select Case KeyAscii
        Case 8, 49 To 53:
        Case 13
            I = 0: J = 0
            For I = 1 To txtniveles.Text
                Label43(I - 1).Visible = True
                txtdescri1(I - 1).Visible = True
                txtnivel1(I - 1).Visible = True
            Next
            For J = I To 5
                Label43(J - 1).Visible = False
                txtdescri1(J - 1).Visible = False
                txtnivel1(J - 1).Visible = False
            Next
        Case Else
        KeyAscii = 0
    End Select
        
End Sub

Private Sub txtnomempresa_GotFocus()

    txtnomempresa.SelStart = 0: txtnomempresa.SelLength = Len(txtnomempresa.Text)
    
End Sub

Private Sub txtnota_oc_GotFocus()
    txtnota_oc.SelStart = 0: txtnota_oc.SelLength = Len(txtnota_oc.Text)
End Sub

Private Sub txtnota_oci_GotFocus()
    txtnota_oci.SelStart = 0: txtnota_oci.SelLength = Len(txtnota_oci.Text)
End Sub

Private Sub txtobsfec_oc_GotFocus()
    txtobsfec_oc.SelStart = 0: txtobsfec_oc.SelLength = Len(txtobsfec_oc.Text)
End Sub

Private Sub txtobsfec_oci_GotFocus()
    txtobsfec_oci.SelStart = 0: txtobsfec_oci.SelLength = Len(txtobsfec_oci.Text)
End Sub

Private Sub txtobsgen_oc_GotFocus()
    txtobsgen_oc.SelStart = 0: txtobsgen_oc.SelLength = Len(txtobsgen_oc.Text)
End Sub

Private Sub txtobsgen_oci_GotFocus()
    txtobsgen_oci.SelStart = 0: txtobsgen_oci.SelLength = Len(txtobsgen_oci.Text)
End Sub

Private Sub txtorigen_GotFocus()
    
    txtorigen.SelStart = 0: txtorigen.SelLength = Len(txtorigen.Text)

End Sub

Private Sub txtruc_GotFocus()

    txtruc.SelStart = 0: txtruc.SelLength = Len(txtruc.Text)
    
End Sub

Private Sub txtsalidaalm_GotFocus()

    txtsalidaalm.SelStart = 0: txtsalidaalm.SelLength = Len(txtsalidaalm.Text)
    
End Sub

Private Sub txtsalidaxtransf_GotFocus()

    txtsalidaxtransf.SelStart = 0: txtsalidaxtransf.SelLength = Len(txtsalidaxtransf.Text)
    
End Sub

Private Sub txttexto_oc_GotFocus()

    txttexto_oc.SelStart = 0: txttexto_oc.SelLength = Len(txttexto_oc.Text)
    
End Sub

Private Sub txttexto_sol_GotFocus()

    txttexto_sol.SelStart = 0: txttexto_sol.SelLength = Len(txttexto_sol.Text)
    
End Sub
