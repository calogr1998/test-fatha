VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmcnttr 
   Appearance      =   0  'Flat
   Caption         =   "Transferencia de Movimientos"
   ClientHeight    =   2100
   ClientLeft      =   2940
   ClientTop       =   2235
   ClientWidth     =   5250
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   5250
   Begin Threed.SSCommand cmdaceptar 
      Height          =   375
      Left            =   1215
      TabIndex        =   2
      Top             =   1620
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   1455
      Left            =   90
      ScaleHeight     =   1395
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   90
      Width           =   5055
      Begin Threed.SSCheck chktrans 
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1035
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "No transferir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin VB.TextBox txtnummov 
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
         Left            =   2565
         MaxLength       =   7
         TabIndex        =   1
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox txtmes 
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
         Left            =   2565
         TabIndex        =   0
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº Movimiento"
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
         Left            =   1260
         TabIndex        =   6
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   1260
         TabIndex        =   5
         Top             =   315
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   375
      Left            =   2610
      TabIndex        =   3
      Top             =   1620
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
End
Attribute VB_Name = "frmcnttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cmdaceptar_Click()
'Dim dbcompras2      As DAO.Database
'Dim tbcabregis2     As DAO.Recordset
'
'    Set dbcompras2 = OpenDatabase(wrutabancos & "\DB_BANCOS.mdb")
'    Set tbcabregis2 = dbcompras2.OpenRecordset("REGISDOC")
'    tbcabregis2.Index = "IDMESNUM"
'    tbcabregis2.Seek "=", txtmes.Text, txtnummov.Text
'    If Not tbcabregis2.NoMatch Then
'        tbcabregis2.Edit
'        If chktrans.Value = True Then
'            tbcabregis2.Fields("f4contable") = "P"
'        Else
'            tbcabregis2.Fields("f4contable") = " "
'        End If
'        tbcabregis2.Update
'
'        Set dbconta = OpenDatabase(wrutatemp & "\db_conta.mdb")
'        Set tbconta = dbconta.OpenRecordset("contable")
'        tbconta.Index = "contable"
'        tbconta.Seek "=", worigen & right(txtnummov.Text, 5)
'        If Not tbconta.NoMatch Then
'            Do While tbconta.Fields("f3compro") = worigen & right(txtnummov.Text, 5) And Not tbconta.EOF
'                tbconta.Delete
'                tbconta.MoveNext
'                If tbconta.EOF Then Exit Do
'                If tbconta.Fields("f3compro") <> worigen & right(txtnummov.Text, 5) Then Exit Do
'            Loop
'        End If
'        frm_cons.dataconta.Refresh
'
'        tbcabregis2.Close
'        dbcompras2.Close
'    Else
'        MsgBox "El movimiento no ha sido grabado. Verifique.", 48, "Compras"
'        txtmes.SetFocus
'    End If
'
'End Sub
'
'Private Sub cmdsalir_Click()
'
'    Unload Me
'
'End Sub
'
'Private Sub txtmes_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        txtmes.Text = Format(txtmes.Text, "00")
'        txtnummov.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtnummov_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        txtnummov.Text = Format(txtnummov.Text, "0000000")
'        cmdaceptar.SetFocus
'    End If
'
'End Sub
'
