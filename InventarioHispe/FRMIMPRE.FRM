VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmimpre 
   Appearance      =   0  'Flat
   Caption         =   "Impresión"
   ClientHeight    =   2580
   ClientLeft      =   4230
   ClientTop       =   3810
   ClientWidth     =   3960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   2580
   ScaleWidth      =   3960
   Begin Crystal.CrystalReport rptres 
      Left            =   3555
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Crystal.CrystalReport report1 
      Left            =   48
      Top             =   2064
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Threed.SSCommand CmdSalir 
      Height          =   390
      Left            =   2070
      TabIndex        =   5
      Top             =   2025
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand CmdAceptar 
      Height          =   390
      Left            =   675
      TabIndex        =   4
      Top             =   2025
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   1770
      Left            =   135
      ScaleHeight     =   1710
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   135
      Width           =   3750
      Begin Threed.SSCheck chkagrupa 
         Height          =   204
         Left            =   192
         TabIndex        =   3
         Top             =   1392
         Width           =   3324
         _Version        =   65536
         _ExtentX        =   5863
         _ExtentY        =   360
         _StockProps     =   78
         Caption         =   "Agrupar por Tipo de Documento"
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
      Begin Threed.SSOption optimpre 
         Height          =   204
         Index           =   1
         Left            =   1248
         TabIndex        =   2
         Top             =   816
         Width           =   1404
         _Version        =   65536
         _ExtentX        =   2476
         _ExtentY        =   360
         _StockProps     =   78
         Caption         =   "132 columnas"
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
      Begin Threed.SSOption optimpre 
         Height          =   348
         Index           =   0
         Left            =   1248
         TabIndex        =   1
         Top             =   192
         Width           =   1548
         _Version        =   65536
         _ExtentX        =   2730
         _ExtentY        =   614
         _StockProps     =   78
         Caption         =   "80 columnas"
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
End
Attribute VB_Name = "frmimpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ruc_empre       As String
Dim linea           As String
Dim FILA            As Integer
Dim xtipdoc         As String
Dim xf4basimp       As Double
Dim xf4igv          As Double
Dim xf4igv3         As Double
Dim xf4monina       As Double
Dim xf4moninad      As Double
Dim xf4total        As Double
Dim xf4totald       As Double
Dim NUMPAG          As Integer

Private Sub CABECERA()
    linea = Space(230)
    Mid(linea, 1, 50) = wnomcia
    Mid(linea, 122, 7) = "Fecha: "
    Mid(linea, 130, 10) = Format(Now, "dd/mm/yyyy")
    WriteXY_texto linea, FILA, 1, 0
    
    linea = Space(230)
    Mid(linea, 1, 18) = "RUC - " & ruc_empre
    Mid(linea, 122, 18) = "Página: "
    WriteXY_texto linea, FILA, 1, 0

    linea = Space(230)
    Mid(linea, 75, 20) = "REGISTRO DE COMPRAS"
    WriteXY_texto linea, FILA, 1, 0
    
    linea = Space(230)
    Mid(linea, 77, 50) = dev_mes(Val(FrmRepRegistro.txtmes.Text)) & " - " & wanno
    WriteXY_texto linea, FILA, 1, 0

    WriteXY_texto String(157, "-"), 8, 1, 0

    linea = Space(230)
    Mid(linea, 1, 7) = "Nº Reg."
    Mid(linea, 9, 5) = "Fecha"
    Mid(linea, 20, 3) = "T/D"
    Mid(linea, 24, 5) = "Serie"
    Mid(linea, 30, 10) = "Nº Comprob"
    Mid(linea, 41, 3) = "RUC"
    Mid(linea, 50, 9) = "Proveedor"
    Mid(linea, 75, 7) = "Detalle"
    Mid(linea, 97, 8) = "Bas.Imp."
    Mid(linea, 112, 3) = "Igv"
    Mid(linea, 125, 5) = "Exon."
    Mid(linea, 137, 10) = "Otros Imp."
    Mid(linea, 150, 5) = "Total"

    WriteXY_texto linea, FILA, 1, 0

    WriteXY_texto String(157, "-"), 10, 1, 0

End Sub

Private Sub Cabecera1()
    linea = Space(230)
    Mid(linea, 1, 50) = wnomcia
    Mid(linea, 122, 7) = "Fecha: "
    Mid(linea, 130, 10) = Format(Now, "dd/mm/yyyy")
    WriteXY_texto linea, FILA, 1, 0
    
    linea = Space(230)
    Mid(linea, 1, 18) = "RUC - " & ruc_empre
    Mid(linea, 122, 18) = "Página: "
    WriteXY_texto linea, FILA, 1, 0

    linea = Space(230)
    Mid(linea, 75, 20) = "REGISTRO DE COMPRAS"
    WriteXY_texto linea, FILA, 1, 0
    
    linea = Space(230)
    Mid(linea, 77, 50) = dev_mes(Val(FrmRepRegistro.txtmes.Text)) & " - " & wanno
    WriteXY_texto linea, FILA, 1, 0

    WriteXY_texto String(157, "-"), 8, 1, 0

    linea = Space(230)
    Mid(linea, 1, 7) = "Nº Reg."
    Mid(linea, 9, 21) = "Nombre/RazonSocial"
    Mid(linea, 40, 4) = "RUC"
    Mid(linea, 50, 5) = "F.Doc"
    Mid(linea, 60, 3) = "ID"
    Mid(linea, 67, 4) = "Ser"
    Mid(linea, 72, 8) = "Numero"
    Mid(linea, 84, 9) = "ValorAf."
    Mid(linea, 93, 13) = "Importaciones"
    Mid(linea, 111, 14) = "Fact.Exterior"
    Mid(linea, 128, 10) = "Honorarios"
    Mid(linea, 141, 9) = "Inafectos"
    Mid(linea, 152, 11) = "Igv.Afectos"
    Mid(linea, 164, 13) = "Igv.Importac."
    Mid(linea, 180, 7) = "4ta+IES"
    Mid(linea, 190, 5) = "Otros"
    Mid(linea, 199, 11) = "Proveedores"
    Mid(linea, 213, 15) = "Otras Modifica."

    WriteXY_texto linea, FILA, 1, 0
    WriteXY_texto String(157, "-"), 10, 1, 0
End Sub

Private Sub cmdaceptar_Click()
Dim xmes            As String
Dim sw_detalle      As Integer

    '------------------------ pregunta si desea cerrar el mes
    Set dbcontrol = OpenDatabase(App.Path & "\CTRCOM.MDB")
    Set tbparametro1 = dbcontrol.OpenRecordset("PARAM_COM")
    tbparametro1.Index = "IDCODEMP"

    xmes = Format(FrmRepRegistro.txtmes.Text, "00")
    tbparametro1.Seek "=", wempresa
    If Not tbparametro1.NoMatch Then
        If tbparametro1.Fields("f1mes" & xmes) & "" <> "*" Then
            If MsgBox("Desea cerrar el mes ?", 36, "Compras") = 6 Then
                tbparametro1.Edit
                tbparametro1.Fields("f1mes" & xmes) = "*"
                tbparametro1.Update
            End If
        End If
    End If
    tbparametro1.Close
    dbcontrol.Close
    
    '----------------------------------------------------------
    sw_detalle = False
    If wf1formato = "9" Then
        If MsgBox("Desea imprimir el detalle ?", 36, "Compras") = 6 Then
            sw_detalle = True
        End If
    End If
    '----------------------------------------------------------

    If optimpre(0).Value = True Then  '--- 80 columnas
        Me.MousePointer = 11
        Report1.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
        If chkagrupa.Value = True Then
            Select Case wf1formato
                Case "6":   '---- BERUDISA
                    Report1.ReportFileName = wrutatemp & "\" & "reg80_be.rpt"
                Case "8":   '---- CAPITEL
                    Report1.ReportFileName = wrutatemp & "\" & "regis80C.rpt"
                Case Else   '---- TODOS LOS DEMAS : ORDENADO POR FECHA
                    If FrmRepRegistro.optorden(2).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regis80f.rpt"
                    Else    '---- TODOS LOS DEMAS : ORDENADO POR Nº MOV
                        Report1.ReportFileName = wrutatemp & "\" & "regis80.rpt"
                    End If
            End Select
        Else
            Report1.ReportFileName = wrutatemp & "\" & "regis8s.rpt"
        End If
        Report1.Action = 1
        Me.MousePointer = 1
    Else   '--- 132 columnas
        Me.MousePointer = 11
        Report1.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
        Select Case wf1formato
            Case "2"    '--- TECPRO
                Report1.ReportFileName = wrutatemp & "\" & "regcom02.rpt"
                Report1.Action = 1
            Case "3"    '--- CENTRO KOR
                Report1.ReportFileName = wrutatemp & "\" & "regcom03.rpt"
                Report1.Action = 1
            Case "4"    '--- OTRO
                Report1.ReportFileName = wrutatemp & "\" & "regcom04.rpt"
                Report1.Action = 1
            Case "5"   '------ Comercial Alimenticia
                Report1.ReportFileName = wrutatemp & "\" & "regcom05.rpt"
                Report1.Action = 1
            Case "6"        '---- LA ALAMEDA
                reg_texto
            Case "7"        '---- CENTRO KOR - TEXTO
                reg_texto1
            Case "8":   '---- CAPITEL
                Report1.ReportFileName = wrutatemp & "\" & "regis80C.rpt"
            Case "9":   '---- AGROPECUARIA ESMERALDA
                If sw_detalle = True Then
                    If FrmRepRegistro.optorden(0).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regcom09a.rpt"
                    End If
                    If FrmRepRegistro.optorden(1).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regcom09b.rpt"
                    End If
                    If FrmRepRegistro.optorden(2).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regcom09c.rpt"
                    End If
                Else
                    If FrmRepRegistro.optorden(0).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regcom09a2.rpt"
                    End If
                    If FrmRepRegistro.optorden(1).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regcom09b2.rpt"
                    End If
                    If FrmRepRegistro.optorden(2).Value = True Then
                        Report1.ReportFileName = wrutatemp & "\" & "regcom09c2.rpt"
                    End If
                End If
                Report1.Action = 1
            Case Else       '---- TODOS LOS DEMAS
                Report1.ReportFileName = wrutatemp & "\" & "regis132.rpt"
                Report1.Action = 1
        End Select
        
        Me.MousePointer = 1
    End If

    '---------------------------------------------------
    '----------- Imprime el Resumen de Compras
    If wf1formato = "5" Then
        If MsgBox("Desea imprimir el Resumen de Compras ?", 36, "Compras") = 6 Then
            procesa_resumen
        End If
    End If
    '---------------------------------------------------

End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub procesa_resumen()
Dim dbbase As DAO.Database
Dim tbcons As DAO.Recordset
Dim tbres  As DAO.Recordset

    Set dbbase = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set tbcons = dbbase.OpenRecordset("select * from temp_regis order by f4tipdoc")
    dbbase.Execute ("delete * from temp_resumen")
    Set tbres = dbbase.OpenRecordset("temp_resumen")
    If tbcons.RecordCount > 0 Then
        tbcons.MoveFirst
        Do While Not tbcons.EOF
            xtipdoc = tbcons.Fields("f4tipdoc") & ""
            xf4basimp = 0#: xf4igv = 0#: xf4igv3 = 0#: xf4monina = 0#
            xf4moninad = 0#: xf4total = 0#: xf4totald = 0#
            Do While xtipdoc = tbcons.Fields("f4tipdoc") & "" And Not tbcons.EOF
                xf4basimp = xf4basimp + tbcons.Fields("f4basimp")
                xf4igv = xf4igv + tbcons.Fields("f4igv")
                xf4igv3 = xf4igv3 + tbcons.Fields("f4igv3")
                xf4monina = xf4monina + tbcons.Fields("f4monina")
                xf4moninad = xf4moninad + tbcons.Fields("f4moninad")
                xf4total = xf4total + tbcons.Fields("f4total")
                xf4totald = xf4totald + tbcons.Fields("f4totald")
                tbcons.MoveNext
                If tbcons.EOF Then Exit Do
                If xtipdoc <> tbcons.Fields("f4tipdoc") & "" Then Exit Do
            Loop
            If xf4basimp <> 0# Or xf4igv <> 0# Or xf4igv3 <> 0# Or xf4monina <> 0# Or xf4moninad <> 0# Or xf4total <> 0# Or xf4totald <> 0# Then
                TbDocumento.Index = "IDCODDOC"
                TbDocumento.Seek "=", xtipdoc
                If Not TbDocumento.NoMatch Then
                    tbres.AddNew
                    tbres.Fields("tipdoc") = xtipdoc
                    tbres.Fields("descdoc") = "T O T A L " & Space(5) & UCase(Trim(TbDocumento.Fields("f2desdoc") & ""))
                    tbres.Fields("f4basimp") = xf4basimp
                    tbres.Fields("f4igv") = xf4igv
                    tbres.Fields("f4igv3") = xf4igv3
                    tbres.Fields("f4monina") = xf4monina
                    tbres.Fields("f4moninad") = xf4moninad
                    tbres.Fields("f4total") = xf4total
                    tbres.Fields("f4totald") = xf4totald
                    tbres.Fields("EMPRESA") = wnomcia
                    tbres.Fields("F4MES") = F4MES
                    tbres.Update
                End If
            End If
        Loop
        If tbcons.RecordCount = 1 Then
            tbres.AddNew
            tbres.Fields("EMPRESA") = wnomcia
            tbres.Fields("F4MES") = F4MES
            tbres.Update
        End If
    Else
        tbres.AddNew
        tbres.Fields("EMPRESA") = wnomcia
        tbres.Fields("F4MES") = F4MES
        tbres.Update
    End If

    tbcons.Close
    tbres.Close
    dbbase.Close

    rptres.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
    rptres.ReportFileName = wrutatemp & "\" & "rpt_resc.rpt"
    rptres.Action = 1

End Sub

Private Sub reg_texto()
Dim FILA            As Integer
Dim wtipo           As String
Dim dbtempo         As DAO.Database
Dim tbtempo         As DAO.Recordset
Dim ssfactura       As DAO.Recordset
Dim ws1             As Double
Dim ws2             As Double
Dim ws3             As Double
Dim ws4             As Double
Dim ws5             As Double
Dim w1              As Double
Dim w2              As Double
Dim w3              As Double
Dim w4              As Double
Dim w5              As Double

    ws1 = 0#: ws2 = 0#: ws3 = 0#: ws4 = 0#: ws5 = 0#
    Set dbcontrol = OpenDatabase(App.Path & "\CTRCOM.MDB")
    Set tbparametro1 = dbcontrol.OpenRecordset("PARAM_COM")
    tbparametro1.Index = "IDCODEMP"
    tbparametro1.Seek "=", wempresa
    If Not tbparametro1.NoMatch Then
        ruc_empre = tbparametro1.Fields("F1RUCEMP") & ""
    End If
    tbparametro1.Close
    dbcontrol.Close
    
    Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set tbtempo = dbtempo.OpenRecordset("temp_regis")
    Set ssfactura = dbtempo.OpenRecordset("select * from temp_regis order by f4nummov")

    If Not ssfactura.EOF Then

        SeteaImpresora 0
        
        'Printer.ScaleMode = 4
        'Printer.FontSize = 10
        CABECERA
        FILA = 12
        w1 = 0#: w2 = 0#: w3 = 0#: w4 = 0#: w5 = 0#
        NUMPAG = 0
        Do While Not ssfactura.EOF
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 1: PRINTER.Print ssfactura.Fields("f4nummov")
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 9: PRINTER.Print Format(ssfactura.Fields("F4FECHA"), "DD/MM/YYYY")
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 20: PRINTER.Print ssfactura.Fields("f4tipdoc") & ""
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 24: PRINTER.Print Trim(ssfactura.Fields("F4SERDOC") & "")
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 30: PRINTER.Print Trim(ssfactura.Fields("f4numdoc") & "")
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 41: PRINTER.Print Trim(ssfactura.Fields("f4rucprv") & "")
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 50: PRINTER.Print Left(Trim(ssfactura.Fields("f4nomprv") & ""), 25)
            'PRINTER.CurrentY = Fila: PRINTER.CurrentX = 75: PRINTER.Print Left(Trim(ssfactura.Fields("f4refere") & ""), 20)

            linea = Space(230)
            Mid(linea, 1, 7) = Format(Format(ssfactura.Fields("f4nummov") & "", "0000000"), "@@@@@@@")
            Mid(linea, 9, 10) = Format(Format(ssfactura.Fields("f4fecha") & "", "dd/mm/yyyy"), "@@@@@@@@@@")
            Mid(linea, 20, 2) = Format(ssfactura.Fields("f4tipdoc"), "@@")
            Mid(linea, 24, 3) = Format(Format(ssfactura.Fields("f4serdoc") & "", "000"), "@@@")
            Mid(linea, 30, 7) = Format(ssfactura.Fields("f4numdoc") & "", "@@@@@@@")
            Mid(linea, 38, 11) = Format(ssfactura.Fields("f4rucprv") & "", "@@@@@@@@@@@")
            Mid(linea, 50, 24) = Left(ssfactura.Fields("f4nomprv") & "", 25)
            Mid(linea, 75, 20) = Left(ssfactura.Fields("f4refere") & "", 20)
            Mid(linea, 95, 14) = Format(Format(ssfactura.Fields("f4basimp"), "###,###,##0.00"), "@@@@@@@@@@@@@@")
            Mid(linea, 108, 14) = Format(Format(ssfactura.Fields("f4igv"), "#,###,##0.00"), "@@@@@@@@@@@@")
            Mid(linea, 120, 14) = Format(Format(ssfactura.Fields("f4monina"), "###,###,##0.00"), "@@@@@@@@@@@@@@")
            Mid(linea, 134, 14) = Format(Format(ssfactura.Fields("f4otrimp"), "###,##0.00"), "@@@@@@@@@@")
            Mid(linea, 144, 14) = Format(Format(ssfactura.Fields("f4total"), "###,###,##0.00"), "@@@@@@@@@@@@@@")
            
            WriteXY_texto linea, FILA, 1, 0

            'WriteXY ssfactura.Fields("f4nummov") & "", fila, 1, 0
            'WriteXY Format(ssfactura.Fields("F4FECHA"), "DD/MM/YYYY"), fila, 9, 0
            'WriteXY ssfactura.Fields("f4tipdoc") & "", fila, 20, 0
            'WriteXY Trim(ssfactura.Fields("F4SERDOC") & ""), fila, 24, 0
            'WriteXY Trim(ssfactura.Fields("f4numdoc") & ""), fila, 30, 0
            'WriteXY Trim(ssfactura.Fields("f4rucprv") & ""), fila, 41, 0
            'WriteXY Left(Trim(ssfactura.Fields("f4nomprv") & ""), 25), fila, 50, 0
            'WriteXY Left(Trim(ssfactura.Fields("f4refere") & ""), 20), fila, 75, 0

            'WriteXY Val(ssfactura.Fields("f4basimp") & ""), fila, 95, 2
            'WriteXY Val(ssfactura.Fields("f4igv") & ""), fila, 110, 2
            'WriteXY Val(ssfactura.Fields("f4monina") & ""), fila, 122, 2
            'WriteXY Val(ssfactura.Fields("f4otrimp") & ""), fila, 134, 2
            'WriteXY Val(ssfactura.Fields("f4total") & ""), fila, 146, 2
            
            FILA = FILA + 1
            w1 = w1 + Val(ssfactura.Fields("f4basimp") & "")
            ws1 = ws1 + Val(ssfactura.Fields("f4basimp") & "")

            w2 = w2 + Val(ssfactura.Fields("f4igv") & "")
            ws2 = ws2 + Val(ssfactura.Fields("f4igv") & "")

            w3 = w3 + Val(ssfactura.Fields("f4monina") & "")
            ws3 = ws3 + Val(ssfactura.Fields("f4monina") & "")

            w4 = w4 + Val(ssfactura.Fields("f4otrimp") & "")
            ws4 = ws4 + Val(ssfactura.Fields("f4otrimp") & "")

            w5 = w5 + Val(ssfactura.Fields("f4total") & "")
            ws5 = ws5 + Val(ssfactura.Fields("f4total") & "")

            If FILA >= 58 Then
                'Printer.Line (1, Fila)-(155, Fila)
                
                WriteXY_texto String(157, "-"), FILA, 1, 0

                linea = Space(230)
                Mid(linea, 1, 20) = "SUB TOTAL :  S/. -->"
                Mid(linea, 95, 14) = Format(Format(ws1, "###,###,##0.00"), "@@@@@@@@@@@@@@")
                Mid(linea, 108, 14) = Format(Format(ws2, "#,###,##0.00"), "@@@@@@@@@@@@")
                Mid(linea, 120, 14) = Format(Format(ws3, "###,###,##0.00"), "@@@@@@@@@@@@@@")
                Mid(linea, 134, 14) = Format(Format(ws4, "###,##0.00"), "@@@@@@@@@@")
                Mid(linea, 144, 14) = Format(Format(ws5, "###,###,##0.00"), "@@@@@@@@@@@@@@")
                WriteXY_texto linea, FILA, 1, 0

                'WriteXY "SUB TOTAL :  S/. -->", fila + 1, 1, 0
                'WriteXY Format(ws1, "#0.00"), fila + 1, 95, 2
                'WriteXY Format(ws2, "#0.00"), fila + 1, 110, 2
                'WriteXY Format(ws3, "#0.00"), fila + 1, 122, 2
                'WriteXY Format(ws4, "#0.00"), fila + 1, 134, 2
                'WriteXY Format(ws5, "#0.00"), fila + 1, 146, 2
                
                Rem MVV WriteXY Format(w1, "#0.00"), Fila + 1, 95, 2
                Rem MVV WriteXY Format(w2, "#0.00"), Fila + 1, 110, 2
                Rem MVV WriteXY Format(w3, "#0.00"), Fila + 1, 122, 2
                Rem MVV WriteXY Format(w4, "#0.00"), Fila + 1, 134, 2
                Rem MVV WriteXY Format(w5, "#0.00"), Fila + 1, 146, 2
                ws1 = 0#: ws2 = 0#: ws3 = 0#: ws4 = 0#: ws5 = 0#
                'Printer.NewPage
                Print #1, Chr(12)
                CABECERA
                FILA = 12
            End If
            ssfactura.MoveNext
        Loop
        'Printer.Line (1, Fila)-(155, Fila)
        
        WriteXY_texto String(157, "-"), FILA, 1, 0

        linea = Space(230)
        Mid(linea, 1, 20) = "SUB TOTAL :  S/. -->"
        Mid(linea, 95, 14) = Format(Format(ws1, "###,###,##0.00"), "@@@@@@@@@@@@@@")
        Mid(linea, 108, 14) = Format(Format(ws2, "#,###,##0.00"), "@@@@@@@@@@@@")
        Mid(linea, 120, 14) = Format(Format(ws3, "###,###,##0.00"), "@@@@@@@@@@@@@@")
        Mid(linea, 134, 14) = Format(Format(ws4, "###,##0.00"), "@@@@@@@@@@")
        Mid(linea, 144, 14) = Format(Format(ws5, "###,###,##0.00"), "@@@@@@@@@@@@@@")
        WriteXY_texto linea, FILA, 1, 0
        

        'WriteXY "SUB TOTAL :  S/. -->", fila + 1, 1, 0
        'WriteXY Format(ws1, "#0.00"), fila + 1, 95, 2
        'WriteXY Format(ws2, "#0.00"), fila + 1, 110, 2
        'WriteXY Format(ws3, "#0.00"), fila + 1, 122, 2
        'WriteXY Format(ws4, "#0.00"), fila + 1, 134, 2
        'WriteXY Format(ws5, "#0.00"), fila + 1, 146, 2

        
        linea = Space(230)
        Mid(linea, 1, 20) = "TOTAL :  S/. -->"
        Mid(linea, 95, 14) = Format(Format(w1, "###,###,##0.00"), "@@@@@@@@@@@@@@")
        Mid(linea, 108, 14) = Format(Format(w2, "#,###,##0.00"), "@@@@@@@@@@@@")
        Mid(linea, 120, 14) = Format(Format(w3, "###,###,##0.00"), "@@@@@@@@@@@@@@")
        Mid(linea, 134, 14) = Format(Format(w4, "###,##0.00"), "@@@@@@@@@@")
        Mid(linea, 144, 14) = Format(Format(w5, "###,###,##0.00"), "@@@@@@@@@@@@@@")
        WriteXY_texto linea, FILA, 1, 0
        
        
        'WriteXY "TOTAL :  S/. -->", fila + 1, 1, 0
        'WriteXY Format(w1, "#0.00"), fila + 1, 95, 2
        'WriteXY Format(w2, "#0.00"), fila + 1, 110, 2
        'WriteXY Format(w3, "#0.00"), fila + 1, 122, 2
        'WriteXY Format(w4, "#0.00"), fila + 1, 134, 2
        'WriteXY Format(w5, "#0.00"), fila + 1, 146, 2
        'Printer.EndDoc
        
        Print #1, Space(230)
        Close #1
        frmView.Show 1

    End If

    ssfactura.Close
    tbtempo.Close
    dbtempo.Close

End Sub
Private Sub reg_texto1()
Dim FILA            As Integer
Dim wtipo           As String
Dim dbtempo         As DAO.Database
Dim tbtempo         As DAO.Recordset
Dim ssfactura       As DAO.Recordset
Dim ws1             As Double
Dim ws2             As Double
Dim ws3             As Double
Dim ws4             As Double
Dim ws5             As Double
Dim w1              As Double
Dim w2              As Double
Dim w3              As Double
Dim w4              As Double
Dim w5              As Double
Dim s1              As Double
Dim s2              As Double
Dim s3              As Double
Dim s4              As Double
Dim s5              As Double
Dim ss1             As Double
Dim ss2             As Double
Dim ss3             As Double
Dim ss4             As Double
Dim ss5             As Double

    ws1 = 0#: ws2 = 0#: ws3 = 0#: ws4 = 0#: ws5 = 0#
    Set dbcontrol = OpenDatabase(App.Path & "\CTRCOM.MDB")
    Set tbparametro1 = dbcontrol.OpenRecordset("PARAM_COM")
    tbparametro1.Index = "IDCODEMP"
    tbparametro1.Seek "=", wempresa
    If Not tbparametro1.NoMatch Then
        ruc_empre = tbparametro1.Fields("F1RUCEMP") & ""
    End If
    tbparametro1.Close
    dbcontrol.Close
    
    Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set tbtempo = dbtempo.OpenRecordset("temp_regis")
    Set ssfactura = dbtempo.OpenRecordset("select * from temp_regis order by f4nummov")

    If Not ssfactura.EOF Then

        SeteaImpresora 0
       Cabecera1
        FILA = 12
        w1 = 0#: w2 = 0#: w3 = 0#: w4 = 0#: w5 = 0#
        NUMPAG = 0
        Do While Not ssfactura.EOF
            
            linea = Space(230)
         Mid(linea, 1, 7) = Format(Format(ssfactura.Fields("f4nummov") & "", "0000000"), "@@@@@@@")
         Mid(linea, 9, 21) = (ssfactura.Fields("f4NOMPRV"))
         Mid(linea, 35, 11) = Format(ssfactura.Fields("f4rucprv") & "", "@@@@@@@@@@@")
         Mid(linea, 50, 8) = Format(Format(ssfactura.Fields("f4fecha") & "", "dd/mm/yyyy"), "@@@@@@@@@@")
         Mid(linea, 60, 3) = Format(ssfactura.Fields("f4tipdoc"), "@@")
         Mid(linea, 67, 4) = Format(Format(ssfactura.Fields("f4serdoc") & "", "000"), "@@@")
         Mid(linea, 72, 6) = Format(ssfactura.Fields("f4numdoc") & "", "@@@@@@@")
         Mid(linea, 77, 13) = Format(Format(ssfactura.Fields("f4basimp"), "###,###,##0.00"), "@@@@@@@@@@@@@")
         Mid(linea, 90, 13) = Format(Format(ssfactura.Fields("f4importa"), "###,###,##0.00"), "@@@@@@@@@@@@@")
         Mid(linea, 104, 13) = Format(Format(ssfactura.Fields("f4facext"), "###,###,##0.00"), "@@@@@@@@@@@@@")
         Mid(linea, 124, 13) = Format(Format(ssfactura.Fields("f4honorario"), "###,###,##0.00"), "@@@@@@@@@@@@@")
         Mid(linea, 138, 11) = Format(Format(ssfactura.Fields("f4monina"), "###,###,##0.00"), "@@@@@@@@@@@")
                       
         Mid(linea, 178, 9) = Format(Format(ssfactura.Fields("f4_4ta_ies"), "###,###,##0.00"), "@@@@@@@@@")
         'Mid(linea, 209, 11) = Format(Format(ssfactura.Fields("f4proveedor"), "###,###,##0.00"), "@@@@@@@@@@@")
         Mid(linea, 214, 11) = Format(Format(ssfactura.Fields("f4tipcam"), "###,###,##0.00"), "@@@@@@@@@@@")
            
            
            WriteXY_texto linea, FILA, 1, 0

          FILA = FILA + 1
            
            w1 = w1 + Val(ssfactura.Fields("f4basimp") & "")
            ws1 = ws1 + Val(ssfactura.Fields("f4basimp") & "")

            s1 = s1 + Val(ssfactura.Fields("f4importa") & "")
            ss1 = ss1 + Val(ssfactura.Fields("f4importa") & "")
            
            s2 = s2 + Val(ssfactura.Fields("f4facext") & "")
            ss2 = ss2 + Val(ssfactura.Fields("f4facext") & "")
            
            s3 = s3 + Val(ssfactura.Fields("f4honorario") & "")
            ss3 = ss3 + Val(ssfactura.Fields("f4honorario") & "")
            
            w3 = w3 + Val(ssfactura.Fields("f4monina") & "")
            ws3 = ws3 + Val(ssfactura.Fields("f4monina") & "")
            
           '*******falta igv afecto
           
            'w2 = w2 + Val(ssfactura.Fields("f4igv") & "")
           ' ws2 = ws2 + Val(ssfactura.Fields("f4igv") & "")
          
            '******falta igv import
           ' w4 = w4 + Val(ssfactura.Fields("f4otrimp") & "")
           ' ws4 = ws4 + Val(ssfactura.Fields("f4otrimp") & "")

           
            s4 = s4 + Val(ssfactura.Fields("f4_4ta_ies") & "")
            ss4 = ss4 + Val(ssfactura.Fields("f4_4ta_ies") & "")
           
           'falta otros
           's5 = s5 + Val(ssfactura.Fields("         ") & "")
           'ss5 = ss5 + Val(ssfactura.Fields("      ") & "")
           
            's6 = s6 + Val(ssfactura.Fields("f4proveedor") & "")
            'ss6 = ss6 + Val(ssfactura.Fields("f4proveedor") & "")
           
           
          
            If FILA >= 58 Then
          
                
                WriteXY_texto String(157, "-"), FILA, 1, 0

                linea = Space(230)
                Mid(linea, 1, 20) = "SUB TOTAL :  S/. -->"
                Mid(linea, 77, 13) = Format(Format(ws1, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 90, 13) = Format(Format(ss1, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 104, 13) = Format(Format(ss2, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 124, 13) = Format(Format(ss3, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 138, 11) = Format(Format(ws3, "###,###,##0.00"), "@@@@@@@@@@@")
                Mid(linea, 178, 9) = Format(Format(ss3, "###,###,##0.00"), "@@@@@@@@@")
                
                WriteXY_texto linea, FILA, 1, 0
 
                ws1 = 0#: ws2 = 0#: ws3 = 0#: ws4 = 0#: ws5 = 0#
          
                Print #1, Chr(12)
                Cabecera1
                FILA = 12
            End If
            ssfactura.MoveNext
        Loop
        
        
        WriteXY_texto String(157, "-"), FILA, 1, 0

        linea = Space(230)
        Mid(linea, 1, 20) = "SUB TOTAL :  S/. -->"
        
                Mid(linea, 77, 13) = Format(Format(ws1, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 90, 13) = Format(Format(ss1, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 104, 13) = Format(Format(ss2, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 124, 13) = Format(Format(ss3, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 138, 11) = Format(Format(ws3, "###,###,##0.00"), "@@@@@@@@@@@")
                Mid(linea, 178, 9) = Format(Format(ss3, "###,###,##0.00"), "@@@@@@@@@")
        WriteXY_texto linea, FILA, 1, 0
        
        linea = Space(230)
        Mid(linea, 1, 20) = "TOTAL :  S/. -->"
        
                Mid(linea, 77, 13) = Format(Format(w1, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 90, 13) = Format(Format(s1, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 104, 13) = Format(Format(s2, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 124, 13) = Format(Format(s3, "###,###,##0.00"), "@@@@@@@@@@@@@")
                Mid(linea, 138, 11) = Format(Format(w3, "###,###,##0.00"), "@@@@@@@@@@@")
                Mid(linea, 178, 9) = Format(Format(s3, "###,###,##0.00"), "@@@@@@@@@")
      
        WriteXY_texto linea, FILA, 1, 0
        
        Print #1, Space(230)
        Close #1
        frmView.Show 1
    End If
    ssfactura.Close
    tbtempo.Close
    dbtempo.Close

End Sub

Private Sub SeteaImpresora(PImpresion As Integer)
On Error GoTo ErrPrn
    
    If PImpresion = 0 Then
        Open Trim(wusuario) & ".TXT" For Output As #1
    Else
        Printer.ScaleMode = 4
        Printer.FontName = "Courier"
        Printer.FontSize = 12
    End If
    Exit Sub

ErrPrn:
    
    If Err = 380 Then  'Invalid Property Value
        Printer.FontName = "Courier"
        Printer.FontSize = 7.5
    Else
        MsgBox Error$(Err)
    End If
    Exit Sub

End Sub

