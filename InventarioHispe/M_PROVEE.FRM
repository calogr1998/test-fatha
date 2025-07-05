VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmProveedor 
   Appearance      =   0  'Flat
   Caption         =   "Actualizacion de Proveedores"
   ClientHeight    =   5895
   ClientLeft      =   3120
   ClientTop       =   1710
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   6090
   Begin Threed.SSCommand BtnExitO 
      Height          =   330
      Left            =   1740
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   90
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   78
      Picture         =   "M_PROVEE.frx":0000
   End
   Begin Threed.SSCommand BtnPrint 
      Height          =   330
      Left            =   1395
      TabIndex        =   18
      ToolTipText     =   "Imprimir"
      Top             =   90
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   78
      Picture         =   "M_PROVEE.frx":015A
   End
   Begin Threed.SSCommand btnnuevo 
      Height          =   330
      Left            =   315
      TabIndex        =   16
      ToolTipText     =   "Nuevo"
      Top             =   90
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   78
      Picture         =   "M_PROVEE.frx":069C
   End
   Begin Threed.SSCommand BtnDelO 
      Height          =   330
      Left            =   1035
      TabIndex        =   17
      ToolTipText     =   "Eliminar"
      Top             =   90
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   78
      Picture         =   "M_PROVEE.frx":0BDE
   End
   Begin Threed.SSCommand BtnEditO 
      Height          =   330
      Left            =   675
      TabIndex        =   15
      ToolTipText     =   "Grabar"
      Top             =   90
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   78
      Picture         =   "M_PROVEE.frx":0CE8
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5310
      Left            =   225
      TabIndex        =   21
      Top             =   480
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   9366
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtNewRuc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3735
         MaxLength       =   11
         TabIndex        =   34
         Top             =   2745
         Visible         =   0   'False
         Width           =   1575
      End
      Begin Threed.SSPanel TXTDESZON 
         Height          =   285
         Left            =   1935
         TabIndex        =   33
         Top             =   1305
         Width           =   3450
         _Version        =   65536
         _ExtentX        =   6085
         _ExtentY        =   503
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   810
         Left            =   165
         TabIndex        =   32
         Top             =   4335
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   1429
         _StockProps     =   14
         Caption         =   "Tipo "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption opttipo 
            Height          =   255
            Index           =   1
            Left            =   3495
            TabIndex        =   14
            Top             =   360
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Extranjero"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption opttipo 
            Height          =   255
            Index           =   0
            Left            =   735
            TabIndex        =   13
            Top             =   360
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Nacional"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   855
         Left            =   165
         TabIndex        =   31
         Top             =   3450
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   1508
         _StockProps     =   14
         Caption         =   "Moneda "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         Begin Threed.SSOption moneda 
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   12
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Ambos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption moneda 
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   11
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Dólares"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption moneda 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Soles"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSCheck chkorden 
         Height          =   255
         Left            =   195
         TabIndex        =   9
         Top             =   3120
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Orden de Compra"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Txtcodzon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1290
         Width           =   465
      End
      Begin VB.TextBox TxtCate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   8
         Top             =   2730
         Width           =   780
      End
      Begin VB.TextBox TxtRefe 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Top             =   2370
         Width           =   4000
      End
      Begin VB.TextBox TxtCodigo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   1
         Top             =   210
         Width           =   600
      End
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Top             =   570
         Width           =   4000
      End
      Begin VB.TextBox TxtDire 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         TabIndex        =   3
         Top             =   930
         Width           =   4000
      End
      Begin VB.TextBox TxtRUC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         MaxLength       =   11
         TabIndex        =   0
         Top             =   225
         Width           =   1305
      End
      Begin VB.TextBox TxtFono 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         TabIndex        =   5
         Top             =   1650
         Width           =   2730
      End
      Begin VB.TextBox TxtFax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         TabIndex        =   6
         Top             =   2010
         Width           =   2730
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nuevo RUC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2565
         TabIndex        =   35
         Top             =   2775
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4170
         TabIndex        =   30
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Razon Social:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   29
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   28
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "R. U. C. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   26
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   2100
         Width           =   300
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   24
         Top             =   2415
         Width           =   825
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ter.Contable:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   2775
         Width           =   960
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cód.Postal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   22
         Top             =   1380
         Width           =   810
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   270
      TabIndex        =   20
      Top             =   45
      Width           =   1995
      _Version        =   65536
      _ExtentX        =   3519
      _ExtentY        =   741
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin VB.Data DataProveedores 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Menu menuedit 
      Caption         =   "&Edición"
      Visible         =   0   'False
      Begin VB.Menu menu210 
         Caption         =   "&Ver Productos"
      End
   End
End
Attribute VB_Name = "FrmProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wselec      As Integer
Dim wvalor      As Integer
Dim wprovee     As String
Dim dbempre     As DAO.Database
Dim TbProv      As DAO.Recordset
Dim TbZonas     As DAO.Recordset
Dim sw          As Boolean
Dim wcodigo     As String

Private Sub Actualiza_Proveedores()

    txtcodigo.Text = "" & DataProveedores.Recordset.Fields("F2CODPROV")
    TxtNombre.Text = "" & DataProveedores.Recordset.Fields("F2NOMPROV")
    TxtDire.Text = "" & DataProveedores.Recordset.Fields("F2DIRPROV")
    txtruc.Text = "" & DataProveedores.Recordset.Fields("F2NEWRUC")
    Rem NSE txtNewRuc.Text = "" & DataProveedores.Recordset.Fields("F2NEWRUC")
    TxtFono.Text = "" & DataProveedores.Recordset.Fields("F2TELPROV")
    TxtFax.Text = "" & DataProveedores.Recordset.Fields("F2FAXPROV")
    Txtcodzon.Text = "" & DataProveedores.Recordset.Fields("F7CODPOS")
    TxtCate.Text = "" & DataProveedores.Recordset.Fields("F2CODCON")
    If Trim(DataProveedores.Recordset.Fields("F2TIPMON")) = "S" Then
        Moneda(0).Value = True
    Else
        If Trim(DataProveedores.Recordset.Fields("F2TIPMON")) = "D" Then
            Moneda(1).Value = True
        Else
            Moneda(2).Value = True
        End If
    End If
    chkorden.Value = DataProveedores.Recordset.Fields("f2orden")
    If Trim(DataProveedores.Recordset.Fields("F2TIPPROV") & "") = "N" Or Len(Trim(DataProveedores.Recordset.Fields("F2TIPPROV") & "")) = 0 Then
        opttipo(0).Value = True
    Else
        opttipo(1).Value = True
    End If
    txtcodigo.Enabled = False
    txtruc.Enabled = False
    
End Sub

Private Sub BtnDelO_Click()

    Me.MousePointer = 11
    Elimina_Proveedores
    Nuevo_Proveedores
    Me.MousePointer = 1

End Sub

Private Sub BtnEditO_Click()

    Me.MousePointer = 11
    Valida_Datos
    If wvalor = 1 Then
       Graba_Proveedores
    End If
    Me.MousePointer = 1

End Sub

Private Sub BtnExitO_Click()

    Unload Me

End Sub

Private Sub btnnuevo_Click()

    sw = False
    Me.MousePointer = 11
    Nuevo_Proveedores
    'HABILITAR_DATOS
    txtruc.Enabled = True
    txtruc.SetFocus
    txtcodigo.Enabled = True
    Me.MousePointer = 1

End Sub

Private Sub BtnPrint_Click()
Dim Fila        As Integer
Dim TbConsulta  As DAO.Recordset

    If MsgBox("Desea mostrar el Listado de Proveedores por Pantalla ?", 36, "Atención") = 6 Then
        Me.MousePointer = 11
        Rem NSE FrmListProveedor.Show 1
        Me.MousePointer = 1
        Exit Sub
    End If
       
    Me.MousePointer = 11
    
    Set dbcontrol = OpenDatabase(App.Path & "\CTRCOM.MDB")
    Set tbparametro1 = dbcontrol.OpenRecordset("PARAM_COM")
    tbparametro1.Index = "IDCODEMP"
    tbparametro1.Seek "=", Trim(wempresa)
    
    Set TbConsulta = dbempre.OpenRecordset("select * from ef2proveedores order by f2nomprov")

    Printer.ScaleMode = 4
    Printer.FontName = "Courier new" 'Printer.Fonts(tbparametro1.Fields("f1fonnam"))

    Rem NSE cabeceras 80
    Printer.FontSize = 12
    Printer.FontBold = True
    writexy "LISTADO DE PROVEEDORES", 5, 533, 0

    Printer.FontSize = 9
    writexy "RUC", 11, 4, 0
    writexy "NOMBRE O RAZON SOCIAL", 8, 16, 0
    writexy "DIRECCION", 8, 45, 0
    writexy "TELEFONO", 8, 73, 0
    writexy "FAX", 8, 84, 0
    Printer.FontBold = False
    Printer.Line (1, 9)-(90, 9)
    Fila = 10
    
    Do While Not TbConsulta.EOF
        writexy "" & TbConsulta.Fields("F2NEWRUC"), Fila, 2, 0
        writexy "" & Left(TbConsulta.Fields("F2NOMPROV"), 35), Fila, 10, 0
        writexy "" & Left(TbConsulta.Fields("F2DIRPROV"), 35), Fila, 42, 0
        writexy "" & Left(TbConsulta.Fields("F2TELPROV"), 8), Fila, 73, 0
        writexy "" & Left(TbConsulta.Fields("F2FAXPROV"), 8), Fila, 82, 0
        Fila = Fila + 1
        If Fila = 60 Then
            Printer.NewPage
            Rem NSE cabeceras 80
            Printer.FontSize = 12
            Printer.FontBold = True
            writexy "LISTADO DE PROVEEDORES", 5, 30, 0

            Printer.FontSize = 9
            writexy "RUC", 8, 4, 0
            writexy "NOMBRE O RAZON SOCIAL", 8, 13, 0
            writexy "DIRECCION", 8, 45, 0
            writexy "TELEFONO", 8, 73, 0
            writexy "FAX", 8, 84, 0
            Printer.FontBold = False
            Printer.Line (1, 9)-(90, 9)
            Fila = 10
        End If
        TbConsulta.MoveNext
    Loop
    Printer.EndDoc

    TbConsulta.Close
    tbparametro1.Close
    dbcontrol.Close

    Me.MousePointer = 1

End Sub

Private Sub Cal_codigo()
Dim wnum    As Integer
    
    If DataProveedores.Recordset.RecordCount = 0 Then
       wnum = 1
    Else
       DataProveedores.Recordset.MoveLast
       Rem If wprovee = "4" Then
           Do While Not DataProveedores.Recordset.Bof
               If Left(DataProveedores.Recordset.Fields("f2codprov"), 1) = "X" Then
                  DataProveedores.Recordset.MovePrevious
               Else
                  Exit Do
               End If
           Loop
           If DataProveedores.Recordset.Fields("f2codprov") = "9999" Then
              DataProveedores.Recordset.MovePrevious
           End If
       Rem Else
       Rem    If dataproveedores.Recordset.fields("f2codprov") = "999" Then
       Rem       dataproveedores.Recordset.MovePrevious
       Rem    End If
       Rem End If
       wnum = Val(DataProveedores.Recordset.Fields("f2codprov")) + 1
    End If
    Rem If wprovee = "4" Then
        wcodigo = Format(wnum, "0000")
    Rem Else
    Rem    wcodigo = Format(wnum, "000")
    Rem End If
    txtcodigo.Text = wcodigo

End Sub

Private Sub Calcula_Codigo()

    Dim wnum    As Integer
    
    If DataProveedores.Recordset.RecordCount = 0 Then
        wnum = 1
    Else
        DataProveedores.Recordset.FindLast "F2CODPROV like '" + UCase(Left(Trim(TxtNombre.Text), 1)) & "*" + "'"
        If DataProveedores.Recordset.NoMatch Then
            wnum = 1
        Else
            wnum = Val(Right(DataProveedores.Recordset.Fields("F2CODPROV"), 3)) + 1
        End If
        If wnum = 1000 Then
            wnum = 0
        End If
    End If
    wcodigo = UCase(Left(Trim(TxtNombre.Text), 1)) & Format(wnum, "000")
    txtcodigo.Text = wcodigo
    
End Sub


Private Sub datos_prove()
        
    txtcodigo.Text = DataProveedores.Recordset.Fields("F2CODPROV")
    TxtNombre.Text = DataProveedores.Recordset.Fields("F2NOMPROV") & ""
    txtruc.Text = DataProveedores.Recordset.Fields("F2NEWRUC")
    TxtFax.Text = "" & DataProveedores.Recordset.Fields("F2FAXPROV")
    Txtcodzon.Text = "" & DataProveedores.Recordset.Fields("F7codpos")
    TxtFono.Text = "" & DataProveedores.Recordset.Fields("F2TELPROV")
    TxtDire.Text = "" & DataProveedores.Recordset.Fields("F2DIRPROV")
    If Trim("" & DataProveedores.Recordset.Fields("F2TIPMON")) = "S" Then
        Moneda(0).Value = True
    End If
    If Trim("" & DataProveedores.Recordset.Fields("F2TIPMON")) = "D" Then
        Moneda(1).Value = True
    End If
    If Trim("" & DataProveedores.Recordset.Fields("F2TIPMON")) = "A" Then
        Moneda(2).Value = True
    End If
    chkorden.Value = DataProveedores.Recordset.Fields("f2orden")
    TxtCate.Text = DataProveedores.Recordset.Fields("F2CODCON")
    If Trim(DataProveedores.Recordset.Fields("F2TIPPROV") & "") = "N" Or Len(Trim(DataProveedores.Recordset.Fields("F2TIPPROV") & "")) = 0 Then
        opttipo(0).Value = True
    Else
        opttipo(1).Value = True
    End If
        
End Sub

Private Sub Elimina_Proveedores()
    
    Beep
    If MsgBox("Seguro de Eliminar Al Proveedor", 36, "Atención") = 6 Then
       DataProveedores.Recordset.FindFirst "F2NEWRUC= '" + txtruc.Text + "'"
       If Not DataProveedores.Recordset.NoMatch Then
          DataProveedores.Recordset.Delete
          DataProveedores.Refresh
       End If
    End If

End Sub

Private Sub Form_Load()
Dim dbcntbanc As DAO.Database
Dim tbcntbanc As DAO.Recordset
    
    Set dbcntbanc = OpenDatabase(App.Path & "\cnt_banc.mdb")
    Set tbcntbanc = dbcntbanc.OpenRecordset("bf1cnt")
    tbcntbanc.Index = "bf1cnt"
    tbcntbanc.Seek "=", wempresa
    If Not tbcntbanc.NoMatch Then
        wprovee = tbcntbanc.Fields("f1terminal") & ""
    End If
    tbcntbanc.Close
    dbcntbanc.Close
    Set dbempre = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbZonas = dbempre.OpenRecordset("EF2ZONAS")
    Set TbProv = dbempre.OpenRecordset("EF2PROVEEDORES")

    DataProveedores.DatabaseName = wrutabancos & "\DB_BANCOS.MDB"
    DataProveedores.RecordSource = "Select * From Ef2Proveedores Order By F2CODPROV"
    DataProveedores.Refresh
    
    Nuevo_Proveedores

    If Len(Trim(gcodprov)) > 0 Then
        txtcodigo.Text = Trim("" & gcodprov)
        DataProveedores.Recordset.FindFirst "F2CODPROV  = '" + txtcodigo.Text + "'"
        datos_prove
    End If
    
    If rucprv <> "" Then
        txtruc.Text = rucprv
        Nuevo_Proveedores
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TbProv.Close
    TbZonas.Close
    dbempre.Close

End Sub

Private Sub Frame3D1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        If Len(Trim(txtcodigo.Text)) > 0 Then
            menu210.Enabled = True
        Else
            menu210.Enabled = False
        End If
        PopupMenu menuedit
    End If

End Sub

Private Sub genera_ctacont()
Dim dbconta     As DAO.Database
Dim tbplan      As DAO.Recordset
Dim tbctapars   As DAO.Recordset
Dim ngrado      As Integer
Dim v_f5colbal  As String
Dim v_f5colsit1 As String
Dim v_f5colsit2 As String
Dim nro_cuenta  As String
Dim cta_ref     As String
Dim cta_mon     As String
Dim cod_cta     As String

    Set dbconta = OpenDatabase(wrutaconta & "\db_tabla.mdb")
    Set tbplan = dbconta.OpenRecordset("CF5PLA")
    tbplan.Index = "CF5PLA"
    
    Set tbctapars = dbconta.OpenRecordset("ctapars")
    tbctapars.Index = "ctapars"

    If tbctapars.RecordCount > 0 Then
        tbctapars.MoveFirst
        Do While Not tbctapars.EOF
            If tbctapars.Fields("ctatipo") & "" = "P" Then
                tbplan.Index = "cf5pla"
                tbplan.Seek "=", tbctapars.Fields("CUENTA") & ""
                If Not tbplan.NoMatch Then
                    ngrado = tbplan.Fields("F5GRDCTA")
                    ngrado = ngrado + 1
                    v_f5colbal = tbplan.Fields("F5COLBAL") & ""
                    v_f5colsit1 = tbplan.Fields("f5colsit1") & ""
                    v_f5colsit2 = tbplan.Fields("f5colsit2") & ""
                    nro_cuenta = Trim(tbctapars.Fields("CUENTA") & "")
                    cta_ref = tbctapars.Fields("CTAREF")
                    cta_mon = tbctapars.Fields("CTAMON") & ""
                    '-------------------------------------------
                    cod_cta = nro_cuenta & Trim("" & TxtCate.Text)
                    tbplan.Index = "cf5pla"
                    tbplan.Seek "=", cod_cta ' verificar si existe cod_cta n la tabla cf5pla
                    If tbplan.NoMatch Then
                        tbplan.AddNew
                        tbplan.Fields("f5codcta") = cod_cta
                        tbplan.Fields("f5nomcta") = TxtNombre.Text & ""
                        tbplan.Fields("f5grdcta") = ngrado
                        tbplan.Fields("f5mref") = cta_ref
                        tbplan.Fields("f5moneda") = cta_mon
                        tbplan.Fields("F5COLBAL") = v_f5colbal
                        tbplan.Fields("f5colsit1") = v_f5colsit1
                        tbplan.Fields("f5colsit2") = v_f5colsit2
                        tbplan.Update
                    End If
                    '-------------------------------------------
                End If
            End If
            tbctapars.MoveNext
        Loop
    End If

    tbctapars.Close
    tbplan.Close
    dbconta.Close

End Sub

Private Sub Graba_Proveedores()
Dim sw_nuevo As Integer

    gcodprov = "" & Trim(txtcodigo.Text)
    'If Len(Trim("" & gcodprov)) = 0 Then Calcula_codigo: Txtcodigo.Text = gcodprov
    If Len(Trim("" & gcodprov)) = 0 Then
       Cal_codigo
       txtcodigo.Text = gcodprov
    End If
    DataProveedores.Recordset.FindFirst "F2CODPROV = '" + gcodprov + "'"

    If Not DataProveedores.Recordset.NoMatch Then
        sw_nuevo = False
        DataProveedores.Recordset.Edit
    Else
        sw_nuevo = True
        DataProveedores.Recordset.AddNew
    End If
    DataProveedores.Recordset.Fields("F2CODPROV") = gcodprov
    DataProveedores.Recordset.Fields("F2NOMPROV") = "" & Trim(TxtNombre.Text)
    DataProveedores.Recordset.Fields("F2DIRPROV") = "" & Trim(TxtDire.Text)
    Rem wvr DataProveedores.Recordset.Fields("F2RUCPROV") = "" & Trim(TxtRUC.Text)
    DataProveedores.Recordset.Fields("F2NEWRUC") = "" & Trim(txtruc.Text)
    DataProveedores.Recordset.Fields("F2TELPROV") = "" & Trim(TxtFono.Text)
    DataProveedores.Recordset.Fields("F2FAXPROV") = "" & Trim(TxtFax.Text)
    DataProveedores.Recordset.Fields("F7CODPOS") = "" & Trim(Txtcodzon.Text)
    DataProveedores.Recordset.Fields("F2CODCON") = "" & Trim(TxtCate.Text)
    If Moneda(0).Value = True Then
        DataProveedores.Recordset.Fields("F2TIPMON") = "S"
    Else
        If Moneda(1).Value = True Then
            DataProveedores.Recordset.Fields("F2TIPMON") = "D"
        Else
            DataProveedores.Recordset.Fields("F2TIPMON") = "A"
        End If
    End If
    DataProveedores.Recordset.Fields("f2orden") = chkorden.Value
    If opttipo(0).Value = True Then
        DataProveedores.Recordset.Fields("f2tipprov") = "N"
    Else
        DataProveedores.Recordset.Fields("f2tipprov") = "E"
    End If
    DataProveedores.Recordset.Update

    If sw_nuevo = True Then
        If MsgBox("Desea generar la cuenta contable ? ", 36, "Sistema de Compras") = 6 Then
            genera_ctacont
        End If
    End If
    
    DataProveedores.Refresh
    DataProveedores.Recordset.FindFirst "F2CODPROV = '" + gcodprov + "'"

End Sub

Private Sub HABILITAR_DATOS()
'   txtnombre.Enabled = True
'   Txtdire.Enabled = True
'   txtruc.Enabled = True
'   Txtfono.Enabled = True
'   Txtfax.Enabled = True
'   Txtcodzon.Enabled = True
'   txtcate.Enabled = True
'   datos_prove

End Sub

Private Sub menu210_Click()
    
    Me.MousePointer = 11
    Rem NSE FrmProvProd.Show 1
    Me.MousePointer = 1

End Sub

Private Sub Nuevo_Proveedores()
    
    'Calcula_codigo
    txtcodigo.Text = ""
    TxtNombre.Text = ""
    TxtDire.Text = ""
    If rucprv <> "" Then
        txtruc.Text = rucprv
        Else
        txtruc.Text = ""
    End If
    TxtFono.Text = ""
    TxtFax.Text = ""
    TxtRefe.Text = ""
    TxtCate.Text = ""
    txtNewRuc.Text = ""
    Cal_codigo
    If wprovee = "4" Then
        TxtCate.Text = txtcodigo.Text
    Else
        TxtCate.Text = Format(txtcodigo.Text, "000")
    End If
    Moneda(0).Value = True
    chkorden.Value = False
    opttipo(0).Value = True
    Txtcodzon.Text = ""

End Sub

Private Sub TxtCate_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 65 And KeyAscii <= 122 Then
        KeyAscii = 7
       Exit Sub
    End If
    If KeyAscii = 13 Then
      If Len(TxtCate.Text) = 3 Or Len(TxtCate.Text) = 4 Then
         SendKeys "{TAB}"
      End If
    End If
    

End Sub

Private Sub txtcodigo_DblClick()
    sw = True
    txtcodigo_KeyUp 113, 0
    
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 13 Then
        txtcodigo.Text = Trim("" & txtcodigo.Text)
        DataProveedores.Recordset.FindFirst "F2CODPROV  = '" + txtcodigo.Text + "'"
        If Not DataProveedores.Recordset.NoMatch Then
            If sw = False Then
                MsgBox "Código ya existe.", 16, "Atención"
                datos_prove
                txtcodigo.SetFocus
            End If
        Else
            If wprovee = "4" Then
                TxtCate.Text = txtcodigo.Text
            Else
                TxtCate.Text = Format(txtcodigo.Text, "000")
            End If
            TxtNombre.SetFocus
        End If
    
    End If

End Sub





Private Sub txtcodigo_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo ERROR_AYUDA
    If KeyCode = 113 Then
        FrmAyudaProv.Show 1
        txtruc.Text = gcodprov
        DataProveedores.Recordset.FindFirst "F2CODPROV = '" + txtruc.Text + "'"
        If Not DataProveedores.Recordset.NoMatch Then
            datos_prove
            Me.MousePointer = 1
        End If
    End If
    Exit Sub

ERROR_AYUDA:

   MsgBox "Operación no permitida", 16, "Proveedores"
   Resume Next

End Sub

Private Sub txtcodigo_LostFocus()

    If Len(Trim(txtcodigo.Text)) > 0 Then
        txtcodigo.Text = Trim("" & txtcodigo.Text)
        DataProveedores.Recordset.FindFirst "F2CODPROV  = '" + txtcodigo.Text + "'"
        If Not DataProveedores.Recordset.NoMatch Then

            If sw = False Then
                MsgBox "Código ya existe.", 16, "Atención"
                datos_prove
            End If
        Else
            If wprovee = "4" Then
                TxtCate.Text = txtcodigo.Text
            Else
                TxtCate.Text = Format(txtcodigo.Text, "000")
            End If
               'TXTNOMBRE.SetFocus
               
        End If
    
    End If

End Sub

Private Sub Txtcodzon_DblClick()

    Txtcodzon_KeyUp 113, 0

End Sub

Private Sub Txtcodzon_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And Txtcodzon.Text <> "" Then
        Txtcodzon.Text = Trim("" & Txtcodzon.Text)
        TbZonas.Index = "IDCODZON"
        TbZonas.Seek "=", Txtcodzon.Text
        If Not TbZonas.NoMatch Then
            TXTDESZON.Caption = TbZonas.Fields("F2DESZON")
            SendKeys "{TAB}"
        Else
            MsgBox "No existe este Código Postal ", 16, "Atención"
            TXTDESZON.Caption = ""
            Txtcodzon.Text = ""
            Exit Sub
        End If
    End If
    If KeyAscii = 13 And Txtcodzon.Text = "" Then
        TXTDESZON.Caption = ""
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub Txtcodzon_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodzon = "" & Txtcodzon.Text
        FrmAyudaZon.Show 1
        Txtcodzon.Text = gcodzon
        Me.MousePointer = 1
        Txtcodzon_KeyPress 13
    End If

End Sub

Private Sub TxtDire_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If TxtDire.Text = "" Then
            MsgBox " El campo no puede estar vacio", vbYesNo
            TxtDire.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    End If

End Sub

Private Sub TxtDire_KeyUp(KeyCode As Integer, Shift As Integer)

    txtcodigo_KeyUp KeyCode, Shift

End Sub

Private Sub txtfax_KeyPress(KeyAscii As Integer)
 
 If KeyAscii >= 65 And KeyAscii <= 122 Then
       KeyAscii = 7
       Exit Sub
 End If

   If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub TxtFono_KeyPress(KeyAscii As Integer)
    
 If KeyAscii >= 65 And KeyAscii <= 122 Then
       KeyAscii = 7
       Exit Sub
 End If

 If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then

      If TxtNombre.Text = "" Then
         MsgBox " El campo no puede estar vacio", vbYesNo
         TxtNombre.SetFocus
      Else
        Rem Cal_codigo
        SendKeys "{TAB}"
      End If
    End If

End Sub

Private Sub TxtNombre_KeyUp(KeyCode As Integer, Shift As Integer)

    txtcodigo_KeyUp KeyCode, Shift

End Sub

Private Sub TxtRefe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub txtruc_DblClick()
    
    txtcodigo_KeyUp 113, 0

End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 65 And KeyAscii <= 122 Then
        KeyAscii = 7
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        If txtruc.Text = "" Then
            MsgBox "El campo RUC NO debe estar en blanco", 16, "Atención"
            txtruc.SetFocus
        Else
            'If Len(TxtRUC.Text) < 11 Or Len(TxtRUC.Text) > 11 Then
            '    MsgBox "El RUC debe tener 11 dígitos.", 16, "Atención"
            '    TxtRUC.Text = ""
            '    Exit Sub
            'End If
            TbProv.Index = "IDCODPROV"
            TbProv.Seek "=", txtruc.Text
            If Not TbProv.NoMatch Then
                DataProveedores.Recordset.FindFirst "F2NEWRUC = '" + txtruc.Text + "'"
                If Not DataProveedores.Recordset.NoMatch Then
                    datos_prove
                End If
            Else
                SendKeys "{TAB}"
            End If
        End If
    End If

End Sub

Private Sub txtruc_KeyUp(KeyCode As Integer, Shift As Integer)
    
    txtcodigo_KeyUp KeyCode, Shift

End Sub

Private Sub Valida_Datos()
    
    wvalor = 1
    If Len(Trim(txtruc.Text)) = 0 Then
        MsgBox "Error en código del proveedor", 16, "Inventario"
        wvalor = 0
    End If

End Sub

