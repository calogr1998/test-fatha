VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmhonor 
   Caption         =   "Consulta de Honorarios"
   ClientHeight    =   2460
   ClientLeft      =   2685
   ClientTop       =   2505
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4935
   Begin Threed.SSPanel Panel3D1 
      Height          =   1545
      Left            =   225
      TabIndex        =   2
      Top             =   225
      Width           =   4560
      _Version        =   65536
      _ExtentX        =   8043
      _ExtentY        =   2725
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
         Height          =   735
         Left            =   90
         TabIndex        =   5
         Top             =   630
         Width           =   4380
         _Version        =   65536
         _ExtentX        =   7726
         _ExtentY        =   1296
         _StockProps     =   14
         Caption         =   "Impresión"
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
         Begin VB.OptionButton optimpre 
            Caption         =   "132 columnas"
            Height          =   285
            Index           =   1
            Left            =   2475
            TabIndex        =   7
            Top             =   270
            Width           =   1680
         End
         Begin VB.OptionButton optimpre 
            Caption         =   "80 columnas"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   6
            Top             =   315
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin VB.TextBox txtmes 
         Height          =   285
         Left            =   2205
         TabIndex        =   4
         Top             =   135
         Width           =   420
      End
      Begin Crystal.CrystalReport report1 
         Left            =   135
         Top             =   45
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         WindowState     =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   1710
         TabIndex        =   3
         Top             =   180
         Width           =   300
      End
   End
   Begin Threed.SSCommand CmdSalir 
      Height          =   420
      Left            =   2475
      TabIndex        =   1
      Top             =   1935
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
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
   Begin Threed.SSCommand CmdAceptar 
      Height          =   420
      Left            =   1170
      TabIndex        =   0
      Top             =   1935
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
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
End
Attribute VB_Name = "frmhonor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fnummes     As String
Dim wtipdoc     As String

Sub PROCESA()
Dim TbConsulta  As DAO.Recordset
Dim Msql        As String

    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    
    Set dbtempcomp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set Tbtemp_regis = dbtempcomp.OpenRecordset("temp_regis")
    
    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    
    '--------------------------------------
    Fnummes = Val(txtmes.Text)
    Select Case Fnummes
        Case 1: F4MES = "ENERO"
        Case 2: F4MES = "FEBRERO"
        Case 3: F4MES = "MARZO"
        Case 4: F4MES = "ABRIL"
        Case 5: F4MES = "MAYO"
        Case 6: F4MES = "JUNIO"
        Case 7: F4MES = "JULIO"
        Case 8: F4MES = "AGOSTO"
        Case 9: F4MES = "SETIEMBRE"
        Case 10: F4MES = "OCTUBRE"
        Case 11: F4MES = "NOVIEMBRE"
        Case 12: F4MES = "DICIEMBRE"
    End Select
    F4MES = F4MES & " - " & wanno
    '--------------------------------------

    TbDocumento.MoveFirst
    Do While Not TbDocumento.EOF
        If UCase(TbDocumento.Fields("f2abrev") & "") = "HON" Then
            wtipdoc = TbDocumento.Fields("f2coddoc") & ""
            Exit Do
        End If
        TbDocumento.MoveNext
    Loop

    dbtempcomp.Execute ("Delete From temp_regis")
    
    Msql = "Select * From  REGISOFI WHERE F4MESMOV = '" + txtmes.Text + "' and F4TIPDOC = '" + wtipdoc + "' order by F4NUMMOV"
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
    
    If TbConsulta.RecordCount > 0 Then
        TbConsulta.MoveFirst
        Do While Not TbConsulta.EOF
            '-------------------------------------------------
            Tbtemp_regis.AddNew
            Tbtemp_regis.Fields("F4MESMOV") = TbConsulta.Fields("F4MESMOV") & ""
            Tbtemp_regis.Fields("F4NUMMOV") = TbConsulta.Fields("F4NUMMOV") & ""
            Tbtemp_regis.Fields("F4CODPRV") = TbConsulta.Fields("F4CODPRV") & ""
            Tbtemp_regis.Fields("F4NOMPRV") = TbConsulta.Fields("F4NOMPRV") & ""
            Tbtemp_regis.Fields("F4TIPDOC") = TbConsulta.Fields("F4TIPDOC") & ""
            Tbtemp_regis.Fields("F4SERDOC") = TbConsulta.Fields("F4SERDOC") & ""
            Tbtemp_regis.Fields("F4NUMDOC") = TbConsulta.Fields("F4NUMDOC") & ""
            Tbtemp_regis.Fields("F4FECHA") = TbConsulta.Fields("F4FECHA")
            Tbtemp_regis.Fields("F4MONEDA") = TbConsulta.Fields("F4MONEDA") & ""
            Tbtemp_regis.Fields("F4RUCPRV") = TbConsulta.Fields("F4RUCPRV") & ""
            Tbtemp_regis.Fields("F4REFERE") = TbConsulta.Fields("F4REFERE") & ""
            Tbtemp_regis.Fields("F4BASIMP") = Format(TbConsulta.Fields("F4BASIMP"), "###,##0.00")
            Tbtemp_regis.Fields("F4IGV") = Format(TbConsulta.Fields("F4MONTORET"), "###,##0.00")
            Tbtemp_regis.Fields("F4MONINA") = Format(TbConsulta.Fields("F4MONINA"), "###,##0.00")
            Tbtemp_regis.Fields("F4OTRIMP") = Format(TbConsulta.Fields("F4FONAVI"), "###,##0.00")
            Tbtemp_regis.Fields("F4TOTAL") = Format(TbConsulta.Fields("F4TOTAL"), "###,##0.00")
            Tbtemp_regis.Fields("F4TIPCAM") = Format(TbConsulta.Fields("F4TIPCAM"), "###,##0.000")
            Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
            Tbtemp_regis.Fields("F4MES") = F4MES
            TbDocumento.Seek "=", TbConsulta.Fields("F4TIPDOC") & ""
            If Not TbDocumento.NoMatch = True Then
                Tbtemp_regis.Fields("F4DESDOC") = TbDocumento.Fields("f2desdoc") & ""
            Else
                Tbtemp_regis.Fields("F4DESDOC") = " "
            End If
            TbCabRegis.Seek "=", TbConsulta.Fields("F4MESMOV") & "", TbConsulta.Fields("F4NUMMOV") & ""
            If Not TbCabRegis.NoMatch Then
                If TbCabRegis.Fields("f4moneda") = "D" Then
                    Tbtemp_regis.Fields("f4totald") = TbCabRegis.Fields("f4total")
                    Tbtemp_regis.Fields("f4basimpd") = TbCabRegis.Fields("f4basimp")
                    Tbtemp_regis.Fields("f4igvd") = TbCabRegis.Fields("f4igv")
                    Tbtemp_regis.Fields("f4moninad") = TbCabRegis.Fields("f4monina")
                    Tbtemp_regis.Fields("f4otrimpd") = TbCabRegis.Fields("f4otrimp")
                Else
                    Tbtemp_regis.Fields("f4totald") = Format(TbConsulta.Fields("f4total") / TbConsulta.Fields("f4tipcam"), "#0.00")
                    Tbtemp_regis.Fields("f4basimpd") = Format(TbConsulta.Fields("f4basimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                    Tbtemp_regis.Fields("f4igvd") = Format(TbConsulta.Fields("f4igv") / TbConsulta.Fields("f4tipcam"), "#0.00")
                    Tbtemp_regis.Fields("f4moninad") = Format(TbConsulta.Fields("f4monina") / TbConsulta.Fields("f4tipcam"), "#0.00")
                    Tbtemp_regis.Fields("f4otrimpd") = Format(TbConsulta.Fields("f4otrimp") / TbConsulta.Fields("f4tipcam"), "#0.00")
                End If
            End If
            Tbtemp_regis.Update
            '-------------------------------------------------
            TbConsulta.MoveNext
            If TbConsulta.EOF Then Exit Do
        Loop
    End If
    TbConsulta.Close
    
    TbDocumento.Close
    dbcomtabla.Close
    
    Tbtemp_regis.Close
    dbtempcomp.Close
    
    TbCabRegis.Close
    dbcompras.Close

End Sub

Private Sub cmdaceptar_Click()
    
    Me.MousePointer = 11
    If wf1formato = "9" Then
        PROCESA_RECIBOS_CANC
    Else
        PROCESA
    End If
    Me.MousePointer = 1
    
    If optimpre(1).Value = True Then
        Report1.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
        If wf1formatov = "2" Then
            Report1.ReportFileName = wrutatemp & "\" & "reghon02.rpt"
        Else
            Report1.ReportFileName = wrutatemp & "\" & "regishon.rpt"
        End If
    Else
        Report1.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
        Report1.ReportFileName = wrutatemp & "\" & "regish80.rpt"
    End If
    Report1.Action = 1

End Sub

Private Sub cmdsalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    txtmes.Text = mes
    
End Sub

Private Sub txtmes_GotFocus()
    
    txtmes.SelStart = 0: txtmes.SelLength = Len(txtmes.Text)
    
End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Val(txtmes.Text & "") >= 1 And Val(txtmes.Text & "") <= 12 Then
            cmdaceptar.SetFocus
        Else
            MsgBox "El mes es incorrecto. Verifique. ", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub

Private Sub PROCESA_RECIBOS_CANC()
Dim dbctasp         As DAO.Database
Dim tbctasp         As DAO.Recordset
Dim tbctasm         As DAO.Recordset
Dim tbctaspr        As DAO.Recordset

    Set dbctasp = OpenDatabase(wrutabancos & "\CTASPAG.MDB")
    Set tbctasp = dbctasp.OpenRecordset("SELECT * FROM PAG_DCTO WHERE VIA_INGR='2' AND MONTH(PAG_DCTO.FCH_COMP) = " & Val(txtmes.Text) & "")
    Set tbctasm = dbctasp.OpenRecordset("PAG_MVTO")
    Set tbctaspr = dbctasp.OpenRecordset("PAG_DCTO")
    
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbOfiRegis = dbcompras.OpenRecordset("REGISOFI")
    
    Set dbtempcomp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set Tbtemp_regis = dbtempcomp.OpenRecordset("temp_regis")
    
    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    
    '--------------------------------------
    Fnummes = Val(txtmes.Text)
    Select Case Fnummes
        Case 1: F4MES = "ENERO"
        Case 2: F4MES = "FEBRERO"
        Case 3: F4MES = "MARZO"
        Case 4: F4MES = "ABRIL"
        Case 5: F4MES = "MAYO"
        Case 6: F4MES = "JUNIO"
        Case 7: F4MES = "JULIO"
        Case 8: F4MES = "AGOSTO"
        Case 9: F4MES = "SETIEMBRE"
        Case 10: F4MES = "OCTUBRE"
        Case 11: F4MES = "NOVIEMBRE"
        Case 12: F4MES = "DICIEMBRE"
    End Select
    F4MES = F4MES & " - " & wanno
    '--------------------------------------

    TbDocumento.MoveFirst
    Do While Not TbDocumento.EOF
        If UCase(TbDocumento.Fields("f2abrev") & "") = "HON" Then
            wtipdoc = TbDocumento.Fields("f2coddoc") & ""
            Exit Do
        End If
        TbDocumento.MoveNext
    Loop

    dbtempcomp.Execute ("Delete From temp_regis")
    
    If tbctasp.RecordCount > 0 Then
        tbctasp.MoveFirst
        Do While Not tbctasp.EOF
            tbctasm.Index = "idcorrcomp"
            tbctasm.Seek "=", tbctasp.Fields("correla")
            If Not tbctasm.NoMatch Then
                Do While tbctasm.Fields("corr_comp") = tbctasp.Fields("correla") And Not tbctasm.EOF
                    tbctaspr.Index = "nro_corre"
                    tbctaspr.Seek "=", tbctasm.Fields("corr_dcto")
                    If Not tbctaspr.NoMatch Then
                        If UCase(Mid(tbctaspr.Fields("nro_comp") & "", 1, 3)) = "HON" Then
                            PROCESA_REC_COMPRAS tbctaspr.Fields("REG_COM") & "", tbctasp.Fields("MONEDA") & "", tbctasm.Fields("IMPUTASO"), tbctasm.Fields("IMPUTADO")
                        End If
                    End If
                    tbctasm.MoveNext
                    If tbctasm.EOF Then Exit Do
                    If tbctasm.Fields("corr_comp") <> tbctasp.Fields("correla") Then Exit Do
                Loop
            End If
            tbctasp.MoveNext
            If tbctasp.EOF Then Exit Do
        Loop
                        
    Else
        MsgBox "No existen registros para procesar.", 48, "Compras"
    End If
    
    TbDocumento.Close
    dbcomtabla.Close
    
    Tbtemp_regis.Close
    dbtempcomp.Close
    
    TbOfiRegis.Close
    TbCabRegis.Close
    dbcompras.Close
    
    tbctasp.Close
    tbctasm.Close
    tbctaspr.Close
    dbctasp.Close
    
End Sub

Private Sub PROCESA_REC_COMPRAS(pregcom As String, pmonpago As String, pimputaso As Double, pimputado As Double)
Dim nporc       As Double
Dim nmonto      As Double

    TbOfiRegis.Index = "IDMESNUM"
    TbOfiRegis.Seek "=", Mid(pregcom, 1, 2), Mid(pregcom, 3, 7)
    If Not TbOfiRegis.NoMatch Then
        
        If pmonpago = "S" Then
            nmonto = pimputado
        Else
            nmonto = pimputaso
        End If
        
        nporc = Val(Format(nmonto / TbOfiRegis.Fields("F4TOTAL"), "0.00"))
        
        Tbtemp_regis.Index = "IDMESNUM"
        Tbtemp_regis.Seek "=", Mid(pregcom, 1, 2), Mid(pregcom, 3, 7)
        If Tbtemp_regis.NoMatch Then
            Tbtemp_regis.AddNew
            Tbtemp_regis.Fields("F4MESMOV") = TbOfiRegis.Fields("F4MESMOV") & ""
            Tbtemp_regis.Fields("F4NUMMOV") = TbOfiRegis.Fields("F4NUMMOV") & ""
        Else
            Tbtemp_regis.Edit
        End If
        Tbtemp_regis.Fields("F4CODPRV") = TbOfiRegis.Fields("F4CODPRV") & ""
        Tbtemp_regis.Fields("F4NOMPRV") = TbOfiRegis.Fields("F4NOMPRV") & ""
        Tbtemp_regis.Fields("F4TIPDOC") = TbOfiRegis.Fields("F4TIPDOC") & ""
        Tbtemp_regis.Fields("F4SERDOC") = TbOfiRegis.Fields("F4SERDOC") & ""
        Tbtemp_regis.Fields("F4NUMDOC") = TbOfiRegis.Fields("F4NUMDOC") & ""
        Tbtemp_regis.Fields("F4FECHA") = TbOfiRegis.Fields("F4FECHA")
        Tbtemp_regis.Fields("F4MONEDA") = TbOfiRegis.Fields("F4MONEDA") & ""
        Tbtemp_regis.Fields("F4RUCPRV") = TbOfiRegis.Fields("F4RUCPRV") & ""
        Tbtemp_regis.Fields("F4REFERE") = TbOfiRegis.Fields("F4REFERE") & ""
        Tbtemp_regis.Fields("F4BASIMP") = Format(nporc * TbOfiRegis.Fields("F4BASIMP"), "###,##0.00")
        Tbtemp_regis.Fields("F4IGV") = Format(nporc * TbOfiRegis.Fields("F4MONTORET"), "###,##0.00")
        Tbtemp_regis.Fields("F4MONINA") = Format(nporc * TbOfiRegis.Fields("F4MONINA"), "###,##0.00")
        Tbtemp_regis.Fields("F4OTRIMP") = Format(nporc * TbOfiRegis.Fields("F4FONAVI"), "###,##0.00")
        Tbtemp_regis.Fields("F4TOTAL") = nmonto
        Tbtemp_regis.Fields("F4TIPCAM") = Format(TbOfiRegis.Fields("F4TIPCAM"), "###,##0.000")
        Tbtemp_regis.Fields("F4EMPRESA") = wnomcia
        Tbtemp_regis.Fields("F4MES") = F4MES
        
        TbDocumento.Seek "=", TbOfiRegis.Fields("F4TIPDOC") & ""
        If Not TbDocumento.NoMatch = True Then
            Tbtemp_regis.Fields("F4DESDOC") = TbDocumento.Fields("f2desdoc") & ""
        Else
            Tbtemp_regis.Fields("F4DESDOC") = " "
        End If
        TbCabRegis.Seek "=", TbOfiRegis.Fields("F4MESMOV") & "", "" & TbOfiRegis.Fields("F4NUMMOV")
        If Not TbCabRegis.NoMatch Then
            If TbCabRegis.Fields("f4moneda") = "D" Then
                Tbtemp_regis.Fields("f4totald") = Val(Format(nmonto, "0.00"))
                Tbtemp_regis.Fields("f4basimpd") = Val(Format(nporc * TbCabRegis.Fields("f4basimp"), "0.00"))
                Tbtemp_regis.Fields("f4igvd") = Val(Format(nporc * TbCabRegis.Fields("f4igv"), "0.00"))
                Tbtemp_regis.Fields("f4moninad") = Val(Format(nporc * TbCabRegis.Fields("f4monina"), "0.00"))
                Tbtemp_regis.Fields("f4otrimpd") = Val(Format(nporc * TbCabRegis.Fields("f4otrimp"), "0.00"))
            Else
                Tbtemp_regis.Fields("f4totald") = Format((nmonto / TbOfiRegis.Fields("f4tipcam")), "#0.00")
                Tbtemp_regis.Fields("f4basimpd") = Format((TbOfiRegis.Fields("f4basimp") / TbOfiRegis.Fields("f4tipcam")) * nporc, "#0.00")
                Tbtemp_regis.Fields("f4igvd") = Format((TbOfiRegis.Fields("f4igv") / TbOfiRegis.Fields("f4tipcam")) * nporc, "#0.00")
                Tbtemp_regis.Fields("f4moninad") = Format((TbOfiRegis.Fields("f4monina") / TbOfiRegis.Fields("f4tipcam")) * nporc, "#0.00")
                Tbtemp_regis.Fields("f4otrimpd") = Format((TbOfiRegis.Fields("f4otrimp") / TbOfiRegis.Fields("f4tipcam")) * nporc, "#0.00")
            End If
        End If
        Tbtemp_regis.Update
        
    End If
                
End Sub
