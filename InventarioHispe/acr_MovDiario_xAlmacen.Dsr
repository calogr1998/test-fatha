VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_MovDiario_xAlmacen 
   Caption         =   "Proyecto1 - acr_MovDiario_xAlmacen (ActiveReport)"
   ClientHeight    =   10935
   ClientLeft      =   -4845
   ClientTop       =   1800
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19288
   SectionData     =   "acr_MovDiario_xAlmacen.dsx":0000
End
Attribute VB_Name = "acr_MovDiario_xAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ant1 As Double
Dim ant2 As Double
Dim subtotal1(6) As Double
Dim CONT As Long
Dim OpcI As String
Dim OpcS As String
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    Dim strSQL      As String
    Dim rpt         As Object
    Dim prProd As New Recordset
    Dim rs1 As New Recordset
    Set rpt = New acr_MovDiario_V_I
        If prProd.State = 1 Then prProd.Close
        If Link = "I" Then
            strSQL = strSQL + "SELECT DISTINCTROW IF4VALES.F4NUMVAL, SF1ORIGENES.F1NOMORI, left(IF4VALES.F4REFERE,50) AS F4REFERE, IF3VALES.F5CODPRO, left(IF5PLA.F5NOMPRO,50) AS F5NOMPRO, IF5PLA.F7CODMED, IF3VALES.F3CANPRO, IF3VALES.F3VALVTA"
            strSQL = strSQL + " FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)"
            strSQL = strSQL + " WHERE IF4VALES.F2CODALM='" & MovDiario_xAlmacen.Txtcodalm.Text & "' AND LEFT(IF4VALES.F4NUMVAL,1) = 'I' AND IF4VALES.F4FECVAL=CVDate('" & Format(MovDiario_xAlmacen.txtfecha.value, "DD/MM/YYYY") & "')"
            prProd.Open strSQL, cnn_dbbancos, adOpenStatic, adLockReadOnly
            With acr_IDiario_xAlmacen
                .datconexion.ConnectionString = cnn_dbbancos
                .datconexion.Source = strSQL
                .fldempresa.Text = wnomcia
                .fldcodalmacen.Text = MovDiario_xAlmacen.Txtcodalm.Text
                .fldnomalmacen.Text = MovDiario_xAlmacen.Txtnomalm.Caption
                .fldFecha.Text = Format(Date, "dd/mm/yyyy")
                .fldtitulo.Text = "Del  " & Format(MovDiario_xAlmacen.txtfecha.value, "dd/mm/yyyy")
                acr_IDiario_xAlmacen.Show 1
            End With
        ElseIf Link = "S" Then
            strSQL = strSQL + "SELECT DISTINCTROW IF4VALES.F4NUMVAL, SF1ORIGENES.F1NOMORI, left(IF4VALES.F4REFERE,80), IF3VALES.F5CODPRO, LEFT(IF5PLA.F5NOMPRO,120)AS F5NOMPRO,IF4VALES.NUMORDEN, IF5PLA.F7CODMED, IF3VALES.F3CANPRO, IF3VALES.F3VALVTA"
            strSQL = strSQL + " FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF3VALES INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)"
            strSQL = strSQL + " WHERE IF4VALES.F2CODALM='" & MovDiario_xAlmacen.Txtcodalm.Text & "' AND LEFT(IF4VALES.F4NUMVAL,1) = 'S' AND IF4VALES.F4FECVAL=CVDate('" & Format(MovDiario_xAlmacen.txtfecha.value, "DD/MM/YYYY") & "')"
            prProd.Open strSQL, cnn_dbbancos, adOpenStatic, adLockReadOnly
            With acr_SDiario_xAlmacen
                .datconexion.ConnectionString = cnn_dbbancos
                .datconexion.Source = strSQL
                .fldempresa.Text = wnomcia
                .fldcodalmacen.Text = MovDiario_xAlmacen.Txtcodalm.Text
                .fldnomalmacen.Text = MovDiario_xAlmacen.Txtnomalm.Caption
                .fldFecha.Text = Format(Date, "dd/mm/yyyy")
                .fldtitulo.Text = "Del  " & Format(MovDiario_xAlmacen.txtfecha.value, "dd/mm/yyyy")
                acr_SDiario_xAlmacen.Show 1
            End With
        
        Else
            strSQL = "SELECT DISTINCTROW MOV_DIARIO_DIA.F5CODPRO FROM MOV_DIARIO_DIA WHERE INGRESO = " & Format(Link, "0.00") & ""
            prProd.Open strSQL, cnn_form, adOpenStatic, adLockReadOnly
        
            If Not (prProd.EOF Or prProd.Bof) Then
                strSQL = ""
                Link = prProd!f5codpro
                
                If rs1.State = adStateOpen Then rs1.Close
                rs1.Open "DROP TABLE MOV_DIARIO_VALES", cnn_form, adOpenDynamic, adLockOptimistic
                If rs1.State = adStateOpen Then rs1.Close
                sql = "SELECT IF4VALES.F2CODALM, IF3VALES.F5CODPRO, IF5PLA.F5CODFAB,IF3VALES.F3CANPRO, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL,IF4VALES.F4SERDOC,IF4VALES.F4NUMDOC,IF4VALES.F4SERGUIA,IF4VALES.F4NUMGUIA INTO MOV_DIARIO_VALES IN '" & wrutatemp & "\TEMPLUS.MDB' "
                sql = sql + " FROM  (IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO"
                sql = sql + " WHERE (((IF4VALES.F2CODALM)='" & MovDiario_xAlmacen.Txtcodalm.Text & "') and (Left([IF4VALES]![F4NUMVAL],1)='I') AND ((IF4VALES.F4FECVAL)=CVDate('" & Format(MovDiario_xAlmacen.txtfecha.value, "DD/MM/YYYY") & "')))"
                If ctipoadm_bd = "M" Then
                    cnn_form.Execute sql
                     'AlmacenaQuery_sql sql, cnn_form
                Else
                    cnn_dbbancos.Execute sql
                     'AlmacenaQuery_sql sql, cnn_dbbancos
                End If
                    Fecha = MovDiario_xAlmacen.txtfecha.value
                    strSQL = "Select * from MOV_DIARIO_VALES where f5Codpro = '" & Link & "'"
                    'Call ShowReport(rpt, Link, strSQL)
            End If
        End If
End Sub

Private Sub ActiveReport_Initialize()
    With Me.Toolbar.Tools
        .ITEM(0).Visible = False
        .ITEM(2).Caption = "&Imprimir"
        .ITEM(2).Tooltip = "Imprimir"
        .ITEM(4).Tooltip = "Copiar"
        .Insert 5, "&Excel"
        .ITEM(5).AddIcon LoadPicture(App.Path & "\Excel.ico")
        .ITEM(5).Tooltip = "Graba el reporte en un archivo excel(*.xls)"
        .ITEM(5).Enabled = True
        .ITEM(7).Tooltip = "Buscar"
        .ITEM(9).Tooltip = "P�gina �nica"
        .ITEM(10).Tooltip = "P�ginas m�ltiples"
        .ITEM(12).Tooltip = "Zoom (-)"
        .ITEM(13).Tooltip = "Zoom (+)"
        .ITEM(16).Tooltip = "P�gina previa"
        .ITEM(17).Tooltip = "P�gina siguiente"
        .ITEM(20).Caption = "&Anterior"
        .ITEM(21).Caption = "&Siguiente"
        .ITEM(20).Tooltip = ""
        .ITEM(21).Tooltip = ""
        
    End With
End Sub

Private Sub ActiveReport_ReportStart()
sw_Report = False
'datconexion.ConnectionString = cconex_dbbancos
    'inicio = True
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
 Dim oEXL As ActiveReportsExcelExport.ARExportExcel

    If Tool.ID = 4015 Then
        Load frmCommon
        strFilePath = frmCommon.Ruta
        Unload frmCommon
        If Trim(strFilePath) <> "" Then
            Set oEXL = New ActiveReportsExcelExport.ARExportExcel
            oEXL.FileName = strFilePath
            oEXL.Export Me.Pages
            MsgBox "Exportaci�n terminada", vbInformation, "Reporte"
        End If
    End If
End Sub


Private Sub Detail_BeforePrint()
    
'    ant1 = ant1 + CDbl(ENTRADAK.Text) - CDbl(SALIDAK.Text) + CDbl(wcant)
'    fldsaldok.Text = ant1
'    fldsaldok.Text = Format(fldsaldok.Text, "#,###0.00")
'    wcant = 0#
'
'    ant2 = ant2 + CDbl(ENTRADACOS.Text) - CDbl(SALIDACOS.Text) + CDbl(fldpeso_ant.Text)
'    fldsaldo.Text = ant2
'    fldsaldo.Text = Format(fldsaldo.Text, "#,###0.00")
OpcI = "I"
INGRESOS.Hyperlink = INGRESOS.Text
fldcodprod.Hyperlink = fldcodprod.Text
OpcS = "S"
SALIDAS.Hyperlink = SALIDAS.Text

LblIngreso.Hyperlink = left(LblIngreso.Caption, 1)
LblSalida.Hyperlink = left(LblSalida.Caption, 1)
End Sub

Private Sub GroupFooter1_BeforePrint()
    
    flddiferk.Text = Format(Val(Format(ENTRADAK_SUB.Text, "0.00")) - Val(Format(SALIDAK_SUB.Text, "0.00")) + Val(Format(fldcant_ant.Text, "0.00")), "###,###,##0.00")
    'FLDDIFER.Text = Format(Val(Format(ENTRADACOS_SUB.Text, "0.00")) - Val(Format(SALIDACOS_SUB.Text, "0.00")) + Val(Format(fldpeso_ant.Text, "0.00")), "###,###,##0.00")
    FLDDIFER.Text = Format(Val(ant2), "###,###,##0.00")
    
    If Val(Format(flddiferk.Text, "0.00")) < 0 Then
        flddiferk.ForeColor = &HC0&
    Else
        flddiferk.ForeColor = &H0&
    End If
    
    If Val(Format(FLDDIFER.Text, "0.00")) < 0 Then
        FLDDIFER.ForeColor = &HC0&
    Else
        FLDDIFER.ForeColor = &H0&
    End If
    
    If fldcant_ant.Text > 0 Then
        ENTRADAK_SUB.Text = Format(ENTRADAK_SUB + CDbl(fldcant_ant.Text), "#,##0.00")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDAK_SUB.Text = Format(SALIDAK_SUB + Abs(CDbl(fldcant_ant.Text)), "#,##0.00")
    End If
    
    If fldpeso_ant.Text > 0 Then
        ENTRADACOS_SUB.Text = Format(ENTRADACOS_SUB + CDbl(fldpeso_ant.Text), "#,##0.00")
    ElseIf fldcant_ant.Text < 0 Then
        SALIDACOS_SUB.Text = Format(SALIDACOS_SUB + Abs(CDbl(fldpeso_ant.Text)), "#,##0.00")
    End If
    
    subtotal1(1) = subtotal1(1) + CDbl(ENTRADAK_SUB.Text)
    subtotal1(2) = subtotal1(2) + CDbl(SALIDAK_SUB.Text)
    subtotal1(3) = subtotal1(3) + CDbl(flddiferk.Text)
    
    subtotal1(4) = subtotal1(4) + CDbl(ENTRADACOS_SUB.Text)
    subtotal1(5) = subtotal1(5) + CDbl(SALIDACOS_SUB.Text)
    subtotal1(6) = subtotal1(6) + CDbl(FLDDIFER.Text)

End Sub

Private Sub GroupHeader1_Format()

    ant1 = 0
    ant2 = 0
    Saldo_Inicial fldcodprod, kardex.aboDesde.value
    CONT = CONT + 1
    acr_kardex.fldcant_ant = Format(wcant, "###,###,##0.00")
    acr_kardex.fldcant_ant.OutputFormat = "#,##0.00;(#,##0.00)"
    
End Sub


Private Sub ReportFooter_Format()
    
     't1.Text = Format(subtotal1(1), "#,###0.00")
'    t2.Text = Format(subtotal1(2), "#,###0.00")
'    t3.Text = Format(subtotal1(3), "#,###0.00")
'
'    T4.Text = Format(subtotal1(4), "#,###0.00")
'    T5.Text = Format(subtotal1(5), "#,###0.00")
'    T6.Text = Format(subtotal1(6), "#,###0.00")

    'POR INVESTIGAR ANIBAL
''    If Val(Format(t3.Text, "0.00")) < 0 Then
''        t3.ForeColor = &HC0&
''    Else
''        t3.ForeColor = &H0&
''    End If
''
'    If Val(Format(T6.Text, "0.00")) < 0 Then
'        T6.ForeColor = &HC0&
'    Else
'        T6.ForeColor = &H0&
'    End If

End Sub
