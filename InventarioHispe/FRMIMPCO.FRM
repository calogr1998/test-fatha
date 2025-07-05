VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmimpco 
   Caption         =   "Impresión de Comprobantes por Rango"
   ClientHeight    =   2550
   ClientLeft      =   1860
   ClientTop       =   2055
   ClientWidth     =   5115
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5115
   Begin Crystal.CrystalReport Report1 
      Left            =   225
      Top             =   2025
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   420
      Left            =   2565
      TabIndex        =   6
      Top             =   2025
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
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
   Begin Threed.SSCommand cmdaceptar 
      Height          =   420
      Left            =   1170
      TabIndex        =   5
      Top             =   2025
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
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
   Begin Threed.SSPanel Panel3D1 
      Height          =   1860
      Left            =   45
      TabIndex        =   7
      Top             =   90
      Width           =   5010
      _Version        =   65536
      _ExtentX        =   8837
      _ExtentY        =   3281
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      Begin Threed.SSFrame Frame3D1 
         Height          =   555
         Left            =   135
         TabIndex        =   11
         Top             =   1125
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   979
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton opttipo 
            Caption         =   "Impresora"
            Height          =   195
            Index           =   1
            Left            =   3015
            TabIndex        =   4
            Top             =   225
            Width           =   1140
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Pantalla"
            Height          =   195
            Index           =   0
            Left            =   675
            TabIndex        =   3
            Top             =   225
            Value           =   -1  'True
            Width           =   870
         End
      End
      Begin VB.TextBox txthasta 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtdesde 
         Height          =   285
         Left            =   1305
         TabIndex        =   1
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtmes 
         Height          =   330
         Left            =   2475
         TabIndex        =   0
         Top             =   90
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   540
         TabIndex        =   9
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   1980
         TabIndex        =   8
         Top             =   135
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmimpco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbtemp      As DAO.Database
Dim TEMP_CABE   As DAO.Recordset
Dim TEMP_DET    As DAO.Recordset

Sub LLENA_TEMPCAB()

    TEMP_CABE.AddNew
    TEMP_CABE.Fields("F4COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
    TEMP_CABE.Fields("F4TOTDEB") = TbCabRegis.Fields("f4total")
    TEMP_CABE.Fields("F4TOTHAB") = TbCabRegis.Fields("f4total")
    TEMP_CABE.Fields("F4CODORI") = worigen
    TEMP_CABE.Fields("F4ORIGEN") = "Registro de Compras"
    TEMP_CABE.Fields("F4FECHA") = TbCabRegis.Fields("f4fecha")
    TEMP_CABE.Fields("F4TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
    TEMP_CABE.Fields("F4PROVE") = TbCabRegis.Fields("f4nomprv")
    TEMP_CABE.Fields("F4IGV") = TbCabRegis.Fields("f4igv")
    TEMP_CABE.Fields("F4TOTAL") = TbCabRegis.Fields("f4total")
    TEMP_CABE.Fields("F4EMPRESA") = wnomcia
    TEMP_CABE.Fields("F4MONEDA") = IIf(TbCabRegis.Fields("f4moneda") = "S", "SOLES", "DOLARES")
    TEMP_CABE.Fields("f4detalle") = TbCabRegis.Fields("f4refere") & ""
    TEMP_CABE.Fields("f4titulo") = "IMPRESION DEL COMPROBANTE"
    TEMP_CABE.Update

End Sub

Sub LLENA_TEMPDET()
Dim cabrev  As String
Dim conta   As Integer

    cabrev = ""
    TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc")
    If Not TbDocumento.NoMatch Then
        cabrev = TbDocumento.Fields("f2abrev")
    End If

    conta = 1
    TbDetRegis.Index = "idmesnum"
    TbDetRegis.Seek "=", txtmes.Text, TbCabRegis.Fields("f4nummov")
    If Not TbDetRegis.NoMatch Then
        Do While Not TbDetRegis.EOF
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = TbDetRegis.Fields("F3CONCEPTO")
            temp_deta.Fields("F3CODGAS") = TbDetRegis.Fields("F3GASTO")
            temp_deta.Fields("F5CODCTA") = TbDetRegis.Fields("F3CTACON")
            If TbDetRegis.Fields("f3debhab") = "D" Then
                temp_deta.Fields("F3DEBE") = Format(TbDetRegis.Fields("f3importe"), "#0.00")
            Else
                temp_deta.Fields("F3HABER") = Format(TbDetRegis.Fields("f3importe"), "#0.00")
            End If
            temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
            temp_deta.Fields("F3TIPDOC") = cabrev
            temp_deta.Fields("F3COSTO") = TbDetRegis.Fields("F3CENCOS")
            temp_deta.Update
            TbDetRegis.MoveNext
            conta = conta + 1
            If TbDetRegis.EOF Then Exit Do
            If TbDetRegis.Fields("f4nummov") <> TbCabRegis.Fields("f4nummov") Or TbDetRegis.Fields("f4mesmov") <> TbCabRegis.Fields("f4mesmov") Then Exit Do
        Loop
        
        If UCase(cabrev) <> "HON" Then
            If TbCabRegis.Fields("f4igv") <> 0 Then
                temp_deta.AddNew
                temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
                temp_deta.Fields("F3ELEMEN") = conta
                temp_deta.Fields("F3DETALL") = "I.G.V."
                temp_deta.Fields("F3CODGAS") = "IGV"
                temp_deta.Fields("F5CODCTA") = wctaigv
                If UCase(cabrev) = "CRE" Then
                    temp_deta.Fields("F3HABER") = TbCabRegis.Fields("f4igv")
                Else
                    temp_deta.Fields("F3DEBE") = TbCabRegis.Fields("f4igv")
                End If
                temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
                temp_deta.Fields("F3TIPDOC") = ""
                temp_deta.Fields("F3COSTO") = ""
                temp_deta.Update
            End If
            '---------------------------------------------------------
            If TbCabRegis.Fields("f4otrimp") <> 0 Then
                conta = conta + 1
                temp_deta.AddNew
                temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
                temp_deta.Fields("F3ELEMEN") = conta
                temp_deta.Fields("F3DETALL") = "OTROS IMPUESTOS"
                temp_deta.Fields("F3CODGAS") = ""
                temp_deta.Fields("F5CODCTA") = wctaotros
                If UCase(cabrev) = "CRE" Then
                    temp_deta.Fields("F3HABER") = TbCabRegis.Fields("f4otrimp")
                Else
                    temp_deta.Fields("F3DEBE") = TbCabRegis.Fields("f4otrimp")
                End If
                temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
                temp_deta.Fields("F3TIPDOC") = ""
                temp_deta.Fields("F3COSTO") = ""
                temp_deta.Update
            End If
            '---------------------------------------------------------
            If Val(TbCabRegis.Fields("f4redsuma") & "") <> 0 Then
                conta = conta + 1
                temp_deta.AddNew
                temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
                temp_deta.Fields("F3ELEMEN") = conta
                temp_deta.Fields("F3DETALL") = "REDONDEO"
                temp_deta.Fields("F3CODGAS") = ""
                temp_deta.Fields("F5CODCTA") = wredsuma
                If UCase(cabrev) = "CRE" Then
                    temp_deta.Fields("F3HABER") = Val(TbCabRegis.Fields("f4redsuma") & "")
                Else
                    temp_deta.Fields("F3DEBE") = Val(TbCabRegis.Fields("f4redsuma") & "")
                End If
                temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
                temp_deta.Fields("F3TIPDOC") = ""
                temp_deta.Fields("F3COSTO") = ""
                temp_deta.Update
            End If
            '---------------------------------------------------------
            If Val(TbCabRegis.Fields("f4redresta") & "") <> 0 Then
                conta = conta + 1
                temp_deta.AddNew
                temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
                temp_deta.Fields("F3ELEMEN") = conta
                temp_deta.Fields("F3DETALL") = "REDONDEO"
                temp_deta.Fields("F3CODGAS") = ""
                temp_deta.Fields("F5CODCTA") = wredresta
                If UCase(cabrev) = "CRE" Then
                    temp_deta.Fields("F3DEBE") = Val(TbCabRegis.Fields("f4redresta") & "")
                Else
                    temp_deta.Fields("F3HABER") = Val(TbCabRegis.Fields("f4redresta") & "")
                End If
                temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
                temp_deta.Fields("F3TIPDOC") = ""
                temp_deta.Fields("F3COSTO") = ""
                temp_deta.Update
            End If
            '---------------------------------------------------------
        Else
            If TbCabRegis.Fields("f4igv") <> 0 Then
                temp_deta.AddNew
                temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
                temp_deta.Fields("F3ELEMEN") = conta
                temp_deta.Fields("F3DETALL") = "RETENCION"
                temp_deta.Fields("F3CODGAS") = ""
                temp_deta.Fields("F5CODCTA") = wctaret
                temp_deta.Fields("F3HABER") = TbCabRegis.Fields("f4igv")
                temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4igv")
                temp_deta.Fields("F3TIPDOC") = ""
                temp_deta.Fields("F3COSTO") = ""
                temp_deta.Update
            End If
            If TbCabRegis.Fields("f4otrimp") <> 0 Then
                conta = conta + 1
                temp_deta.AddNew
                temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
                temp_deta.Fields("F3ELEMEN") = conta
                temp_deta.Fields("F3DETALL") = "I.E.S."
                temp_deta.Fields("F3CODGAS") = ""
                temp_deta.Fields("F5CODCTA") = wctafon
                temp_deta.Fields("F3HABER") = TbCabRegis.Fields("f4otrimp")
                temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
                temp_deta.Fields("F3TIPDOC") = ""
                temp_deta.Fields("F3COSTO") = ""
                temp_deta.Update
            End If
        End If
               
        If TbCabRegis.Fields("f4total") <> 0 Then
            conta = conta + 1
            temp_deta.AddNew
            temp_deta.Fields("F3COMPRO") = txtmes.Text & "/" & TbCabRegis.Fields("f4nummov")
            temp_deta.Fields("F3ELEMEN") = conta
            temp_deta.Fields("F3DETALL") = Left(Trim(TbCabRegis.Fields("f4NomPrv")) & "", 50)
            temp_deta.Fields("F3CODGAS") = ""
            temp_deta.Fields("F5CODCTA") = Trim(TbCabRegis.Fields("f4ctacont") & "")
            If UCase(cabrev) <> "CRE" Then
                temp_deta.Fields("F3HABER") = TbCabRegis.Fields("f4total")
            Else
                temp_deta.Fields("F3DEBE") = TbCabRegis.Fields("f4total")
            End If
            temp_deta.Fields("F3TIPCAMBD") = TbCabRegis.Fields("f4tipcam")
            temp_deta.Fields("F3TIPDOC") = cabrev
            temp_deta.Fields("F3COSTO") = ""
            temp_deta.Fields("F3DOCUM") = cabrev & TbCabRegis.Fields("f4serdoc") & "/" & TbCabRegis.Fields("f4numdoc")
            temp_deta.Update
        End If
    End If

End Sub

Private Sub cmdaceptar_Click()
Dim i As Double
    
    txtmes.Text = Format(txtmes.Text, "00")
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    Set TEMP_CABE = dbtemp.OpenRecordset("vou_cab")
    Set temp_deta = dbtemp.OpenRecordset("vou_det")

    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.mdb")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbOfiRegis = dbcompras.OpenRecordset("REGISOFI")
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"

    For i = Val(TxtDesde.Text) To Val(TxtHasta.Text)
        TbCabRegis.Seek "=", txtmes.Text, Format(i, "0000000")
        If Not TbCabRegis.NoMatch Then
            dbtemp.Execute "DELETE * FROM VOU_CAB"
            dbtemp.Execute "DELETE * FROM VOU_DET"
            LLENA_TEMPCAB
            LLENA_TEMPDET
            Report1.DataFiles(0) = wrutatemp & "\temp_com.mdb"
            If wf1formatov = "2" Then
                Report1.ReportFileName = wrutatemp & "\rpt_vou2.rpt"
            Else
                Report1.ReportFileName = wrutatemp & "\rpt_vou.rpt"
            End If
            If opttipo(0).Value = True Then
                Report1.Destination = 0
            Else
                Report1.Destination = 1
            End If
            Report1.Action = 1
        End If
    Next
    If opttipo(1).Value = True Then
        CmdSalir.SetFocus
    End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
  txtmes.Text = mes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set dbtemp = Nothing
    Set dbcompras = Nothing
    Set dbcomtabla = Nothing
End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If
End Sub

Private Sub txtdesde_GotFocus()
  TxtDesde.SelStart = 0: TxtDesde.SelLength = Len(TxtDesde.Text)
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtDesde.Text = Format(TxtDesde.Text, "0000000")
        TxtHasta.SetFocus
    End If
End Sub

Private Sub txtdesde_LostFocus()
    If Len(Trim(TxtDesde.Text)) > 0 Then
        TxtDesde.Text = Format(TxtDesde.Text, "0000000")
    End If
End Sub

Private Sub txthasta_GotFocus()
TxtHasta.SelStart = 0: TxtHasta.SelLength = Len(TxtHasta.Text)
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(TxtHasta.Text) >= Val(TxtDesde.Text) Then
            TxtHasta.Text = Format(TxtHasta.Text, "0000000")
            cmdaceptar.SetFocus
        Else
            MsgBox "Error en el rango de comprobantes. Verifique.", 48, "Compras"
            TxtHasta.SetFocus
        End If
    End If
End Sub

Private Sub txthasta_LostFocus()
    If Len(Trim(TxtHasta.Text)) > 0 Then
        If Val(TxtHasta.Text) >= Val(TxtDesde.Text) Then
            TxtHasta.Text = Format(TxtHasta.Text, "0000000")
            If Val(TxtHasta.Text) = Val(TxtDesde.Text) Then
                opttipo(0).Enabled = True
                opttipo(0).Value = True
            Else
                opttipo(0).Enabled = False
                opttipo(1).Value = True
            End If
        Else
            TxtHasta.SetFocus
        End If
    End If
End Sub

Private Sub txtmes_GotFocus()
  txtmes.SelStart = 0: txtmes.SelLength = Len(txtmes.Text)
End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtmes.Text = Format(txtmes.Text, "00")
        If Val(txtmes.Text) >= 1 And Val(txtmes.Text) <= 12 Then
            TxtDesde.SetFocus
        Else
            MsgBox "Error en el mes. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If
End Sub

