VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSaldoAlm 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos Iniciales"
   ClientHeight    =   1920
   ClientLeft      =   3330
   ClientTop       =   3150
   ClientWidth     =   5820
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   5820
   Begin Threed.SSCommand BtnExit 
      Height          =   300
      Left            =   3072
      TabIndex        =   4
      Top             =   1488
      Width           =   1212
      _Version        =   65536
      _ExtentX        =   2138
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand BtnPrint 
      Height          =   300
      Left            =   1632
      TabIndex        =   3
      Top             =   1488
      Width           =   1212
      _Version        =   65536
      _ExtentX        =   2138
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox FrameCliente 
      Appearance      =   0  'Flat
      FillStyle       =   0  'Solid
      ForeColor       =   &H008080FF&
      Height          =   1275
      Index           =   0
      Left            =   180
      ScaleHeight     =   1245
      ScaleWidth      =   5430
      TabIndex        =   5
      Top             =   45
      Width           =   5460
      Begin MSMask.MaskEdBox Txtfecult 
         Height          =   252
         Left            =   4128
         TabIndex        =   2
         Top             =   912
         Visible         =   0   'False
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txtfecreg 
         Height          =   252
         Left            =   960
         TabIndex        =   1
         Top             =   912
         Visible         =   0   'False
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txtnomalm 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1392
         TabIndex        =   9
         Top             =   528
         Width           =   3756
      End
      Begin VB.TextBox Txtcodalm 
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
         Left            =   945
         MaxLength       =   2
         TabIndex        =   0
         Top             =   530
         Width           =   390
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fec. Ult. Registro:"
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
         Index           =   0
         Left            =   2745
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén:"
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
         Index           =   3
         Left            =   225
         TabIndex        =   7
         Top             =   570
         Width           =   660
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Visible         =   0   'False
         Width           =   492
      End
   End
End
Attribute VB_Name = "FrmSaldoAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbempresa As DAO.Database
Dim TBALMACEN As DAO.Recordset 'FVG

Private Sub BtnExit_Click()

    Unload Me

End Sub

Private Sub BtnPrint_Click()
    
    Me.MousePointer = 11
    gcodalm = Trim(Txtcodalm.Text)
    FrmSaldoIni.Show 1
    Me.MousePointer = 1

End Sub

Private Sub Form_Load()
    
    Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TBALMACEN = dbempresa.OpenRecordset("EF2ALMACENES")
    
    TBALMACEN.Index = "IDCODALM"

    Txtfecreg.Text = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo seguir

    TBALMACEN.Close
    dbempresa.Close
    Set dbempresa = Nothing

    Exit Sub
seguir: Resume Next

End Sub

Private Sub Txtcodalm_Change()
   
    TBALMACEN.Seek "=", Txtcodalm.Text
    If Not TBALMACEN.NoMatch Then
        gcodalm = Trim(Txtcodalm.Text)
        Txtnomalm.Text = "" & TBALMACEN.Fields("F2nomalm")
        Txtfecult.Text = Format(TBALMACEN.Fields("F1ultsal"), "dd/mm/yyyy")
    End If

End Sub

Private Sub Txtcodalm_DblClick()

    Txtcodalm_KeyUp 113, 0

End Sub

Private Sub Txtcodalm_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 13 Then
        Txtcodalm.Text = Trim(Txtcodalm.Text)
        TBALMACEN.Seek "=", Txtcodalm.Text
        If Not TBALMACEN.NoMatch Then
            Txtnomalm.Text = "" & TBALMACEN.Fields("F2nomalm")
            Txtfecult.Text = Format(TBALMACEN.Fields("F1ultsal"), "dd/mm/yyyy")
            BtnPrint.SetFocus 'SendKeys "{tab}"
        Else
            Beep
            MsgBox "Almacén NO existe...", 64, "Atención"
            Txtnomalm = "": Txtcodalm.SetFocus
        End If
    End If
   
End Sub

Private Sub Txtcodalm_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodalm = "" & Txtcodalm.Text
        'frmayudaalm.Show 1
        Txtcodalm.Text = gcodalm
        Me.MousePointer = 1
        Txtcodalm_KeyPress 13
    End If

End Sub

Private Sub Txtfecreg_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

