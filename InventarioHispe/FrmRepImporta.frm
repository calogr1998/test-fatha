VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRepImporta 
   Caption         =   "Consulta Importación"
   ClientHeight    =   1995
   ClientLeft      =   2790
   ClientTop       =   1800
   ClientWidth     =   5505
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   5505
   Begin MSMask.MaskEdBox TxtDesde 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Threed.SSCommand CmdSalir 
      Height          =   375
      Left            =   2745
      TabIndex        =   7
      Top             =   1530
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salir"
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
      Font3D          =   3
   End
   Begin Threed.SSCommand CmdAceptar 
      Height          =   375
      Left            =   1485
      TabIndex        =   6
      Top             =   1530
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Aceptar"
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
      Font3D          =   3
   End
   Begin Threed.SSPanel PnlNomPrv 
      Height          =   285
      Left            =   2340
      TabIndex        =   5
      Top             =   1080
      Width           =   2940
      _Version        =   65536
      _ExtentX        =   5186
      _ExtentY        =   503
      _StockProps     =   15
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Font3D          =   3
   End
   Begin VB.TextBox TxtCodPrv 
      Height          =   285
      Left            =   1125
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   735
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Rango de Fecha"
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
      ShadowStyle     =   1
      Begin MSMask.MaskEdBox TxtHasta 
         Height          =   285
         Left            =   3465
         TabIndex        =   9
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2835
         TabIndex        =   2
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   1125
      Width           =   735
   End
End
Attribute VB_Name = "FrmRepImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
    If Len(Trim(TxtCodPrv.Text)) = 0 Then
       MsgBox "Error, Ingrese Proveedor", 48
       Exit Sub
    End If
    Me.MousePointer = vbHourglass
    'FrmResInforme.Show 1
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    
    Set Tbproveedor = dbempresa.OpenRecordset("EF2PROVEEDORES")
    Tbproveedor.Index = "IDCODPROV"

    TxtDesde.Text = "" & Format(Now, "dd/mm/yyyy")
    TxtHasta.Text = "" & Format(Now, "dd/mm/yyyy")

    'TxtDesde.Text

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload FrmAyudaProv
End Sub

Private Sub TxtCodPrv_GotFocus()
    TxtCodPrv.SelStart = 0: TxtCodPrv.SelLength = Len(TxtCodPrv.Text)
End Sub

Private Sub TxtCodPrv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Tbproveedor.Seek "=", TxtCodPrv.Text
       If Not Tbproveedor.NoMatch Then
          pnlnomprv.Caption = "" & Tbproveedor.Fields("F2NOMPROV")
       End If
       SendKeys "{Tab}"
    End If

End Sub

Private Sub TxtCodPrv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Me.MousePointer = vbHourglass
       grucprov = TxtCodPrv.Text
       'FrmAyudaProv.Show 1
       TxtCodPrv.Text = grucprov
       Me.MousePointer = vbDefault
       TxtCodPrv_KeyPress 13
    End If

End Sub

Private Sub txtdesde_GotFocus()
  TxtDesde.SelStart = 0: TxtDesde.SelLength = Len(TxtDesde.Text)
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txthasta_GotFocus()
  TxtHasta.SelStart = 0: TxtHasta.SelLength = Len(TxtHasta.Text)
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
