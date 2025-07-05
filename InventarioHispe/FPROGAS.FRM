VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmProveGasto 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedor / Gasto del mes"
   ClientHeight    =   4965
   ClientLeft      =   2565
   ClientTop       =   1935
   ClientWidth     =   6885
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   6885
   Begin Threed.SSCommand btnsalir 
      Height          =   435
      Left            =   3555
      TabIndex        =   10
      Top             =   4455
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand btnejecutar 
      Height          =   435
      Left            =   2145
      TabIndex        =   9
      Top             =   4455
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D2 
      Height          =   4200
      Left            =   90
      ScaleHeight     =   4140
      ScaleWidth      =   6615
      TabIndex        =   11
      Top             =   135
      Width           =   6675
      Begin Threed.SSFrame SSFrame3 
         Height          =   3030
         Left            =   105
         TabIndex        =   14
         Top             =   975
         Width           =   6390
         _Version        =   65536
         _ExtentX        =   11282
         _ExtentY        =   5355
         _StockProps     =   14
         Caption         =   "Datos Generales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin Threed.SSFrame ptipo 
            Height          =   780
            Left            =   165
            TabIndex        =   20
            Top             =   2010
            Width           =   6075
            _Version        =   65536
            _ExtentX        =   10716
            _ExtentY        =   1376
            _StockProps     =   14
            Caption         =   "Tipo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption tipo 
               Height          =   150
               Index           =   1
               Left            =   4530
               TabIndex        =   8
               Top             =   375
               Width           =   1170
               _Version        =   65536
               _ExtentX        =   2053
               _ExtentY        =   275
               _StockProps     =   78
               Caption         =   "Detallado"
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
            Begin Threed.SSOption tipo 
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   7
               Top             =   375
               Width           =   1065
               _Version        =   65536
               _ExtentX        =   1879
               _ExtentY        =   344
               _StockProps     =   78
               Caption         =   "General"
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
         End
         Begin Threed.SSPanel PnlNomPrv 
            Height          =   390
            Left            =   2250
            TabIndex        =   18
            Top             =   1200
            Width           =   4035
            _Version        =   65536
            _ExtentX        =   7117
            _ExtentY        =   688
            _StockProps     =   15
            BackColor       =   -2147483648
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSFrame SSFrame4 
            Height          =   780
            Left            =   165
            TabIndex        =   17
            Top             =   390
            Width           =   6075
            _Version        =   65536
            _ExtentX        =   10716
            _ExtentY        =   1376
            _StockProps     =   14
            Caption         =   "Moneda"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption moneda 
               Height          =   195
               Index           =   1
               Left            =   4515
               TabIndex        =   4
               Top             =   390
               Width           =   1260
               _Version        =   65536
               _ExtentX        =   2222
               _ExtentY        =   344
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
               Height          =   195
               Index           =   0
               Left            =   675
               TabIndex        =   3
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   344
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
               Value           =   -1  'True
            End
         End
         Begin VB.TextBox txtgasto 
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
            Left            =   1185
            MaxLength       =   8
            TabIndex        =   6
            Top             =   1620
            Width           =   1050
         End
         Begin VB.TextBox TxtCodPrv 
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
            Left            =   1185
            MaxLength       =   11
            TabIndex        =   5
            Top             =   1236
            Width           =   1050
         End
         Begin Threed.SSPanel ngasto 
            Height          =   390
            Left            =   2250
            TabIndex        =   19
            Top             =   1590
            Width           =   4035
            _Version        =   65536
            _ExtentX        =   7117
            _ExtentY        =   688
            _StockProps     =   15
            BackColor       =   -2147483648
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   225
            TabIndex        =   16
            Top             =   1260
            Width           =   750
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Gasto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   240
            TabIndex        =   15
            Top             =   1635
            Width           =   435
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   735
         Left            =   105
         TabIndex        =   13
         Top             =   150
         Width           =   3180
         _Version        =   65536
         _ExtentX        =   5609
         _ExtentY        =   1291
         _StockProps     =   14
         Caption         =   "Modo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption modo 
            Height          =   156
            Index           =   1
            Left            =   1680
            TabIndex        =   1
            Top             =   384
            Width           =   1212
            _Version        =   65536
            _ExtentX        =   2138
            _ExtentY        =   275
            _StockProps     =   78
            Caption         =   "Acumulado"
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
         Begin Threed.SSOption modo 
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   0
            Top             =   336
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   444
            _StockProps     =   78
            Caption         =   "Mensual"
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
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   735
         Left            =   3375
         TabIndex        =   12
         Top             =   150
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5524
         _ExtentY        =   1291
         _StockProps     =   14
         Caption         =   "Meses"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox meses 
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
            Height          =   315
            Left            =   810
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "FrmProveGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnejecutar_Click()

    If modo(0).Value = True Then
        xmodo = "0" 'mensual
    Else
        xmodo = "1" 'acumulado
    End If
    xmes = Format(meses.ListIndex + 1, "00")
    If moneda(0).Value = True Then
        xmoneda = "0" 'soles
    Else
        xmoneda = "1" 'dolares
    End If
    
    Tbproveedor.Index = "IDCODPROV"
    Tbproveedor.Seek "=", TxtCodPrv.Text
    If Not Tbproveedor.NoMatch Then
        xpro = "" & Tbproveedor.Fields("f2codprov")
    Else
        xpro = ""
    End If

    If Trim(txtgasto.Text) <> "" Then
        xgasto = "" & Trim(txtgasto.Text)
        xtipo = "2"     'por gasto
    Else
        If TIPO(0).Value = True Then
            xtipo = "0" 'general"
        Else
            xtipo = "1" 'detallado
        End If
    End If
    Me.MousePointer = 11
    FrmReporte.Show 1
    Me.MousePointer = 1
    
End Sub

Private Sub btnsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Set dbcomtab = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbgastos = dbcomtab.OpenRecordset("BF9GIN")
    tbgastos.Index = "idcodigo"
    Set TbDocumento = dbcomtab.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "IDCODDOC"

    Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set Tbproveedor = dbempresa.OpenRecordset("EF2PROVEEDORES")
    Tbproveedor.Index = "IDCODPROV"
    
    meses.Clear
    meses.AddItem "Enero"
    meses.AddItem "Febrero"
    meses.AddItem "Marzo"
    meses.AddItem "Abril"
    meses.AddItem "Mayo"
    meses.AddItem "Junio"
    meses.AddItem "Julio"
    meses.AddItem "Agosto"
    meses.AddItem "Setiembre"
    meses.AddItem "Octubre"
    meses.AddItem "Noviembre"
    meses.AddItem "Diciembre"
    meses.ListIndex = Val(mes) - 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo error_bd

    Tbproveedor.Close
    dbempresa.Close

    tbgastos.Close
    TbDocumento.Close
    dbcomtab.Close

    Exit Sub

error_bd:
    Resume Next

End Sub

Private Sub meses_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then moneda(0).SetFocus

End Sub

Private Sub modo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then meses.SetFocus

End Sub

Private Sub moneda_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then TxtCodPrv.SetFocus

End Sub

Private Sub tipo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then btnejecutar.SetFocus

End Sub

Private Sub TxtCodPrv_DblClick()

     TxtCodPrv_KeyUp 113, 0

End Sub

Private Sub TxtCodPrv_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(TxtCodPrv.Text) = "" Then
            PnlNomPrv.Caption = "Todos"
            txtgasto.SetFocus
        Else
            Tbproveedor.Seek "=", TxtCodPrv.Text
            If Not Tbproveedor.NoMatch Then
                PnlNomPrv.Caption = "" & Tbproveedor.Fields("F2NOMPROV")
                gcodprov = "" & Tbproveedor.Fields("f2codprov")
                txtgasto.SetFocus
            Else
                MsgBox "Proveedor no Existe...", 16, "Atención"
                PnlNomPrv.Caption = "": TxtCodPrv.SetFocus
            End If
        End If
    End If

End Sub

Private Sub TxtCodPrv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
       Me.MousePointer = 11
       grucprov = TxtCodPrv.Text
       FrmAyudaProv.Show 1
       TxtCodPrv.Text = grucprov
       Me.MousePointer = 1
       TxtCodPrv_KeyPress 13
    End If

End Sub

Private Sub TxtCodPrv_LostFocus()

    If Trim(TxtCodPrv.Text) = "" Then
        PnlNomPrv.Caption = "Todos"
    Else
        Tbproveedor.Seek "=", TxtCodPrv.Text
        If Not Tbproveedor.NoMatch Then
            PnlNomPrv.Caption = "" & Tbproveedor.Fields("F2NOMPROV")
            gcodprov = "" & Tbproveedor.Fields("f2codprov")
        Else
            MsgBox "Proveedor no Existe...", 16, "Atención"
            PnlNomPrv.Caption = ""
        End If
    End If
    
End Sub

Private Sub txtgasto_DblClick()

     TxtGASTO_KeyUp 113, 0

End Sub

Private Sub txtgasto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If Trim(txtgasto.Text) = "" Then
            ngasto.Caption = ""
            ptipo.Visible = True
            TIPO(0).SetFocus
       Else
            ptipo.Visible = False
            tbgastos.Index = "idcodigo"
            tbgastos.Seek "=", "G", Trim(txtgasto.Text)
            If Not tbgastos.NoMatch Then
                 ngasto.Caption = "" & tbgastos.Fields("nombre")
                 btnejecutar.SetFocus
            Else
                 MsgBox "Gasto NO existe...", 16, "Atención"
                 ngasto.Caption = "": txtgasto.SetFocus
            End If
       End If
    End If

End Sub

Private Sub TxtGASTO_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
       Me.MousePointer = 11
       GCODGAS = txtgasto.Text
       frmhlpgastos.Show 1
       txtgasto.Text = GCODGAS
       Me.MousePointer = 1
       txtgasto_KeyPress 13
    End If

End Sub

Private Sub txtgasto_LostFocus()

    If Trim(txtgasto.Text) = "" Then
        ngasto.Caption = ""
        ptipo.Visible = True
    Else
        ptipo.Visible = False
        tbgastos.Index = "idcodigo"
        tbgastos.Seek "=", "G", Trim(txtgasto.Text)
        If Not tbgastos.NoMatch Then
            ngasto.Caption = "" & tbgastos.Fields("nombre")
        Else
            MsgBox "Gasto NO existe...", 16, "Atención"
            ngasto.Caption = ""
        End If
    End If
    
            
End Sub

