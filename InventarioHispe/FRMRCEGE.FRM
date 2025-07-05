VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmrcege 
   Caption         =   "Reporte por Centro de Costo"
   ClientHeight    =   4140
   ClientLeft      =   1995
   ClientTop       =   1335
   ClientWidth     =   7305
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7305
   Begin Crystal.CrystalReport cryreporte 
      Left            =   5580
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   420
      Left            =   3690
      TabIndex        =   14
      Top             =   3600
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
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
      Left            =   2385
      TabIndex        =   13
      Top             =   3600
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
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
      Height          =   3345
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   7125
      _Version        =   65536
      _ExtentX        =   12568
      _ExtentY        =   5900
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
      Begin VB.TextBox txtprov 
         Height          =   285
         Left            =   1575
         MaxLength       =   4
         TabIndex        =   16
         Top             =   585
         Width           =   870
      End
      Begin Threed.SSFrame Frame3D3 
         Height          =   645
         Left            =   180
         TabIndex        =   10
         Top             =   2520
         Width           =   6765
         _Version        =   65536
         _ExtentX        =   11933
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Tipo"
         ForeColor       =   8388608
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
         Begin VB.OptionButton opttipo 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   1
            Left            =   4995
            TabIndex        =   12
            Top             =   225
            Width           =   1050
         End
         Begin VB.OptionButton opttipo 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   0
            Left            =   1125
            TabIndex        =   11
            Top             =   225
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin Threed.SSFrame Frame3D2 
         Height          =   645
         Left            =   180
         TabIndex        =   7
         Top             =   1845
         Width           =   6765
         _Version        =   65536
         _ExtentX        =   11933
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Ordenado por"
         ForeColor       =   8388608
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
         Begin VB.OptionButton optorden 
            Caption         =   "Gasto"
            Height          =   195
            Index           =   1
            Left            =   4995
            TabIndex        =   9
            Top             =   225
            Width           =   825
         End
         Begin VB.OptionButton optorden 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   0
            Left            =   1125
            TabIndex        =   8
            Top             =   225
            Value           =   -1  'True
            Width           =   1680
         End
      End
      Begin Threed.SSFrame Frame3D1 
         Height          =   645
         Left            =   180
         TabIndex        =   4
         Top             =   1080
         Width           =   6765
         _Version        =   65536
         _ExtentX        =   11933
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Moneda"
         ForeColor       =   8388608
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
         Begin VB.OptionButton optmoneda 
            Caption         =   "Dólares"
            Height          =   195
            Index           =   1
            Left            =   4995
            TabIndex        =   6
            Top             =   225
            Width           =   915
         End
         Begin VB.OptionButton optmoneda 
            Caption         =   "Soles"
            Height          =   195
            Index           =   0
            Left            =   1125
            TabIndex        =   5
            Top             =   225
            Value           =   -1  'True
            Width           =   1725
         End
      End
      Begin Threed.SSPanel pnlcentro 
         Height          =   285
         Left            =   2565
         TabIndex        =   3
         Top             =   180
         Width           =   4380
         _Version        =   65536
         _ExtentX        =   7726
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtcentro 
         Height          =   285
         Left            =   1575
         MaxLength       =   8
         TabIndex        =   2
         Top             =   180
         Width           =   870
      End
      Begin Threed.SSPanel pnlprov 
         Height          =   285
         Left            =   2565
         TabIndex        =   15
         Top             =   585
         Width           =   4380
         _Version        =   65536
         _ExtentX        =   7726
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         Enabled         =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmrcege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbtemp As DAO.Recordset
Dim xnombre As String
Dim xx      As String
Dim sw      As Integer
Dim Msql    As String

Sub PROCESA(xcentro)
Dim TbConsulta  As DAO.Recordset
Dim tbgasto     As DAO.Recordset
Dim tbtempcab   As DAO.Recordset
Dim nmonto      As Double
Dim nigv        As Double
Dim xnomgasto   As String
Dim cabrev      As String
    
    Me.MousePointer = 11
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("delete * from temp_centros")
    dbtemp.Execute ("delete * from temp_centroscab")
    Set tbtemp = dbtemp.OpenRecordset("temp_centros")
    Set tbtempcab = dbtemp.OpenRecordset("temp_centroscab")
    tbtempcab.Index = "codgasto"

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    Set tbgasto = dbcomtabla.OpenRecordset("BF9GIN")
    
    If Len(txtprov.Text & "") = 0 Then
        Msql = "Select * From REGISMOV WHERE F3CENCOS = '" + xcentro + "' Order by F4NUMMOV"
    Else
        Msql = "Select REGISMOV.F4MESMOV,REGISMOV.F4NUMMOV,REGISMOV.F3GASTO,REGISMOV.F3IMPORTE,REGISMOV.F3CONCEPTO From REGISMOV,REGISDOC WHERE REGISDOC.F4MESMOV=REGISMOV.F4MESMOV AND REGISDOC.F4NUMMOV=REGISMOV.F4NUMMOV AND REGISMOV.F3CENCOS = '" & xcentro & "' AND REGISDOC.F4CODPRV = '" & txtprov.Text & "' Order by REGISMOV.F4MESMOV,REGISMOV.F4NUMMOV"
    End If
    
    Set TbConsulta = dbcompras.OpenRecordset(Msql)

    If TbConsulta.RecordCount > 0 Then
        TbConsulta.MoveFirst
        Do While Not TbConsulta.EOF
            TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
            If Not TbCabRegis.NoMatch Then
                nmonto = 0#:   nigv = 0#:  xnomgasto = ""
                tbtemp.AddNew
                tbgasto.Index = "idcod"
                tbgasto.Seek "=", TbConsulta.Fields("f3gasto") & ""
                If Not tbgasto.NoMatch Then
                    tbtemp.Fields("F3DESGTO") = tbgasto.Fields("nombre") & ""
                    xnomgasto = tbgasto.Fields("nombre") & ""
                End If
                tbtemp.Fields("f3gasto") = TbConsulta.Fields("f3gasto") & ""
                tbtemp.Fields("f3centro") = xcentro
                tbtemp.Fields("f3nomcentro") = xnombre
                tbtemp.Fields("f3nummov") = TbConsulta.Fields("f4nummov")
                tbtemp.Fields("f3fecha") = TbCabRegis.Fields("f4fecha")
                cabrev = ""
                TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc")
                If Not TbDocumento.NoMatch Then
                    cabrev = TbDocumento.Fields("f2abrev") & ""
                End If
                tbtemp.Fields("f3docum") = cabrev & TbCabRegis.Fields("f4serdoc") & "/" & TbCabRegis.Fields("f4numdoc")
                tbtemp.Fields("f3codprov") = TbCabRegis.Fields("f4codprv") & ""
                tbtemp.Fields("f3nomprov") = TbCabRegis.Fields("f4nomprv") & ""
                tbtemp.Fields("f3tipcamb") = TbCabRegis.Fields("f4tipcam")
                If optmoneda(0).Value = True Then
                    If TbCabRegis.Fields("f4moneda") = "S" Then
                        tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                        tbtemp.Fields("f3igv") = Format(TbConsulta.Fields("f3importe") * 0.19, "0.00")
                        nmonto = TbConsulta.Fields("f3importe")
                        nigv = Format(TbConsulta.Fields("f3importe") * 0.19, "0.00")
                    Else
                        tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        tbtemp.Fields("f3igv") = Format((TbConsulta.Fields("f3importe") * 0.19) * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        nmonto = Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        nigv = Format((TbConsulta.Fields("f3importe") * 0.19) * TbCabRegis.Fields("f4tipcam"), "#0.00")
                    End If
                Else
                    If TbCabRegis.Fields("f4moneda") = "D" Then
                        tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                        tbtemp.Fields("f3igv") = Format(TbConsulta.Fields("f3importe") * 0.19, "0.00")
                        nmonto = TbConsulta.Fields("f3importe")
                        nigv = Format(TbConsulta.Fields("f3importe") * 0.19, "0.00")
                    Else
                        tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        tbtemp.Fields("f3igv") = Format((TbConsulta.Fields("f3importe") * 0.19) / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        nmonto = Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        nigv = Format((TbConsulta.Fields("f3importe") * 0.19) / TbCabRegis.Fields("f4tipcam"), "#0.00")
                    End If
                End If
                tbtemp.Fields("f3moneda") = IIf(optmoneda(0).Value = True, "( SOLES )", "( DOLARES )")
                tbtemp.Fields("f3simbolo") = IIf(optmoneda(0).Value = True, "S/.", "US$")
                tbtemp.Fields("empresa") = wnomcia
                tbtemp.Fields("titulo") = IIf(optmoneda(0).Value = True, "( SOLES )", "( DOLARES )")
                tbtemp.Fields("f3detalle") = TbConsulta.Fields("f3concepto") & ""
                tbtemp.Update
                '---------------------------------
                tbtempcab.Seek "=", TbConsulta.Fields("f3gasto") & ""
                If Not tbtempcab.NoMatch Then
                    tbtempcab.Edit
                Else
                    tbtempcab.AddNew
                    tbtempcab.Fields("codgasto") = TbConsulta.Fields("f3gasto") & ""
                    tbtempcab.Fields("nomgasto") = xnomgasto
                    tbtempcab.Fields("empresa") = wnomcia
                    tbtempcab.Fields("titulo") = IIf(optmoneda(0).Value = True, "( SOLES )", "( DOLARES )")
                    tbtempcab.Fields("codcentro") = xcentro
                    tbtempcab.Fields("nomcentro") = xnombre
                End If
                tbtempcab.Fields("importe") = tbtempcab.Fields("importe") + nmonto
                tbtempcab.Fields("igv") = tbtempcab.Fields("igv") + nigv
                tbtempcab.Update
                '---------------------------------
            End If
            TbConsulta.MoveNext
        Loop
    End If

    TbCabRegis.Close
    TbDetRegis.Close
    TbConsulta.Close

    tbtemp.Close
    tbtempcab.Close
    dbtemp.Close

    TbDocumento.Close
    tbgasto.Close
    dbcomtabla.Close
    Me.MousePointer = 1

End Sub

Private Sub cmdaceptar_Click()
Dim xccosto As String

    xccosto = Trim(txtcentro.Text)
    xnombre = pnlcentro.Caption
    PROCESA xccosto
    If optorden(0).Value = True Then
        cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
        cryreporte.ReportFileName = wrutatemp & "\rpt_cent.rpt"
        cryreporte.Action = 1
    Else
        If opttipo(0).Value = True Then
            cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
            cryreporte.ReportFileName = wrutatemp & "\rpt_cen2.rpt"
            cryreporte.Action = 1
        Else
            cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
            cryreporte.ReportFileName = wrutatemp & "\rpt_cen3.rpt"
            cryreporte.Action = 1
        End If
    End If

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Set dbempresa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set Tbproveedor = dbempresa.OpenRecordset("EF2PROVEEDORES")
    Tbproveedor.Index = "IDCODIGO"
    
    Set dbcentros = OpenDatabase(wrutaconta & "\db_costo.mdb")
    Set tbcentros = dbcentros.OpenRecordset("centros")
    sw = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Tbproveedor.Close
    dbempresa.Close
    
    tbcentros.Close
    dbcentros.Close

End Sub

Private Sub optmoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        optorden(1).SetFocus
    End If

End Sub

Private Sub optorden_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        opttipo(0).SetFocus
    End If

End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If

End Sub

Private Sub txtcentro_DblClick()
    txtcentro_KeyDown 113, 0
End Sub

Private Sub txtcentro_GotFocus()
    txtcentro.SelStart = 0: txtcentro.SelLength = Len(txtcentro.Text)
End Sub

Private Sub txtcentro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw = True
        gcodcen = txtcentro.Text
        frmayucen.Show 1
        sw = False
        If Len(Trim(gcodcen)) <> 0 Then           '**********
            txtcentro.Text = gcodcen
            pnlcentro.Caption = Trim(gnomcen)     '**********
        End If
        txtcentro_KeyPress 13
    End If

End Sub

Private Sub txtcentro_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        tbcentros.Index = "centros"
        tbcentros.Seek "=", txtcentro.Text
        If Not tbcentros.NoMatch Then
            pnlcentro.Caption = Trim(tbcentros.Fields("f3descrip") & "")
            txtprov.SetFocus
        Else
            MsgBox "Código del centro de costo no existe. Verifique.", 48, "Compras"
            txtcentro.SetFocus
        End If
    End If

End Sub

Private Sub txtcentro_LostFocus()
    
    If sw = False Then
        If Len(Trim(txtcentro.Text)) > 0 Then
            tbcentros.Index = "centros"
            tbcentros.Seek "=", txtcentro.Text
            If Not tbcentros.NoMatch Then
                pnlcentro.Caption = tbcentros.Fields("f3descrip") & ""
                txtprov.SetFocus
            Else
                MsgBox "Código del centro de costo no existe. Verifique.", 48, "Compras"
                txtcentro.SetFocus
            End If
        End If
    End If

End Sub

Private Sub txtprov_DblClick()

    txtprov_KeyDown 113, 0
    
End Sub

Private Sub txtprov_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        gcodprov = ""
        FrmAyudaProv.Show 1
        txtprov.Text = gcodprov
        txtprov_KeyPress 13
    End If

End Sub

Private Sub txtprov_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtprov.Text)) > 0 Then
            Tbproveedor.Index = "IDCODIGO"
            Tbproveedor.Seek "=", txtprov.Text
            If Not Tbproveedor.NoMatch Then
                pnlprov.Caption = Trim(Tbproveedor.Fields("f2nomprov") & "")
                optmoneda(1).SetFocus
            Else
                MsgBox "El código del proveedor no existe. Verifique.", 48, "Compras"
                txtprov.SetFocus
            End If
        Else
            pnlprov.Caption = "TODOS"
            optmoneda(1).SetFocus
        End If
    End If

End Sub
