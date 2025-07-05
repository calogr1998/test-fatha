Attribute VB_Name = "jackie"

'-----------------------RECORDSET ---------------

Global rsconsulta     As New ADODB.Recordset
Global Temp           As New ADODB.Connection
Global RsMovAlmacen   As ADODB.Recordset
Global RsAlmacenes    As ADODB.Recordset
Global RsStockDet     As ADODB.Recordset
Global RsStockCab     As ADODB.Recordset
Global RsProducto     As ADODB.Recordset
Global RsParametro    As ADODB.Recordset
Global RsCentros      As ADODB.Recordset
Global RsMedida       As ADODB.Recordset
Global RsOrigen       As ADODB.Recordset
Global RsNivel01      As ADODB.Recordset
Global RsNivel02      As ADODB.Recordset
Global RsNivel03      As ADODB.Recordset
Global RsNivel04      As ADODB.Recordset
Global RST            As ADODB.Recordset
Global RsPartida      As ADODB.Recordset
Global RsParain       As ADODB.Recordset

'------------------ VARIABLES -------------------

Public cconex_inventa                 As String
Global F                              As Integer
Global wmes                           As Integer
Global stockact, sal, sald, ing, ingd As Single
Global Cospro, Cosprod, Debm, Habm    As Double
Global StockLog                       As Double
Global FecUlt                         As Date
Global SQL, SQL1, wnumval, Gtipval    As String
Global wcodPrv, wnomPrv, WcodPar      As String
Global wnomcosto, WNomPar             As String
Global wcodori, wnomori, wparact_stock As String
Global wtipcam As Double
Global ctipo As String * 1
Global wstock, wvalvta  As Double
Global sw_ocompra              As Boolean

Public Function VALIDA_FPAGO(pfpago As String)
Dim sw      As Boolean

    sw = False
    If RST.State = adStateOpen Then RST.Close
    RST.Open "Select F2DESPAG from ef2forpag where f2forpag='" & Trim(pfpago) & "'", cnn_dbbancos  'cnn_dbbancos
    If RST.EOF = False Then
        wnompag = RST!F2DESPAG & ""
        sw = True
    Else
        sw = False
    End If
    RST.Close
    VALIDA_FPAGO = sw
        
End Function

Public Function VALIDA_CLIENTE(pcodcli As String)
Dim sw      As Boolean
    
    If RST.State = adStateOpen Then RST.Close
    RST.Open "Select f2CODcli,f2nomcli,F2NEWRUC,f2dircli,f2forpag,f2lista_asig from ef2clientes where f2codcli='" & pcodcli & "' OR f2NEWRUC='" & pcodcli & "'", cnn_dbbancos
    If RST.EOF = False Then
        wcodcli = "" & RST!f2CODcli
        wnomcli = "" & RST!f2nomcli
        wruccli = "" & RST!F2NEWRUC
        wdircli = "" & RST!f2dircli
        wforpag = "" & RST!f2forpag
        nnumlista = Val("" & RST.Fields("f2lista_asig") & "")
        sw = True
    Else
        sw = False
    End If
    RST.Close
    VALIDA_CLIENTE = sw
        
End Function

Public Function VALIDA_CLIENTE_V(pcodcli As String, pcodven As String)
Dim sw      As Boolean
    
    If RST.State = adStateOpen Then RST.Close
    RST.Open "Select f2CODcli,f2nomcli,F2NEWRUC,f2dircli,f2forpag,f2lista_asig from ef2clientes where (f2codcli='" & pcodcli & "' OR f2newruc='" & pcodcli & "') AND (F2CODVEN='" & pcodven & "')", cnn_dbbancos
    If RST.EOF = False Then
        wcodcli = "" & RST!f2CODcli
        wnomcli = "" & RST!f2nomcli
        wruccli = "" & RST!F2NEWRUC
        wdircli = "" & RST!f2dircli
        wforpag = "" & RST!f2forpag
        nnumlista = Val(RST.Fields("f2lista_asig") & "")
        sw = True
    Else
        sw = False
    End If
    RST.Close
    VALIDA_CLIENTE_V = sw
        
End Function

Public Function VALIDA_RESPONSABLE(pcodres As String)
Dim sw      As Boolean

    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from EF2VENDEDORES where F2CODVEN='" & pcodres & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        wnomven = rs.Fields("F2NOMVEN") & ""
        wfacmin_nac = Val(Format("" & rs.Fields("F2FACNAC_MIN"), "0.00"))
        wfacmin_imp = Val(Format("" & rs.Fields("F2FACIMP_MIN"), "0.00"))
        sw = True
    Else
        sw = False
    End If
    rs.Close
    VALIDA_RESPONSABLE = sw
        
End Function

Public Function VALIDA_ORIGEN(PCodOri As String)
Dim sw      As Boolean
Set RST = New ADODB.Recordset
    sw = False
    If RST.State = adStateOpen Then RST.Close
    RST.Open "Select * from Sf1Origenes where F1CodOri='" & Trim(PCodOri) & "'", cnn_dbbancos
    If Not RST.EOF Then
        wcodori = Trim(RST!f1codori & "")
        wnomori = Trim(RST!F1NOMORI & "")
        sw = True
    Else
        sw = False
    End If
    RST.Close
    VALIDA_ORIGEN = sw

End Function
Public Function VALIDA_ALMACENO(pcodalm As String)
Dim sw  As Boolean
Set rs = New ADODB.Recordset

    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select F2NOMALM from ef2almacenes where f2codalm='" & Trim(pcodalm) & "'", cnn_dbbancos
    If rs.EOF = False Then
        WNomPar = rs!F2NOMALM & ""
        sw = True
    Else
        sw = False
    End If
    rs.Close
    VALIDA_ALMACENO = sw
        
End Function

Public Sub GRABAR_ACTUALIZACIONES(Fmes As Integer, Alma As String, cod As String)
Dim amovs_arr(0 To 7) As a_grabacion

amovs_arr(0).campo = "F5DEBM" & Format(Fmes, "00"): amovs_arr(0).valor = Debm: amovs_arr(0).TIPO = "N"
amovs_arr(1).campo = "F5ING" & Format(Fmes, "00"): amovs_arr(1).valor = ing: amovs_arr(1).TIPO = "N"
amovs_arr(2).campo = "F5INGD" & Format(Fmes, "00"): amovs_arr(2).valor = ingd: amovs_arr(2).TIPO = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).TIPO = "N"
amovs_arr(4).campo = "F6STOCKLOG": amovs_arr(4).valor = StockLog: amovs_arr(4).TIPO = "N"
amovs_arr(5).campo = "F6FECULT": amovs_arr(5).valor = FecUlt: amovs_arr(5).TIPO = "F"
amovs_arr(6).campo = "F5COSPRO": amovs_arr(6).valor = Cospro: amovs_arr(6).TIPO = "N"
amovs_arr(7).campo = "F5COSPROD": amovs_arr(7).valor = Cosprod: amovs_arr(7).TIPO = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO amovs_arr(), "IF6ALMA", "M", 7, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub Vales_Detalle(PNUMVAL As String, pcodpro As String, pcanpro As Double, pvalpro As Double, pcodalm As String, PFECVAL As Date, pvaldol As Double)
Dim csql, SSQL As String
Dim WCONT   As Integer
Dim wfecha  As Variant
Dim wprecos As Double
Dim wsaldoc As Double
Dim wsaldom As Double
Dim wsaldod As Double
    
    Set RsMovAlmacen = New ADODB.Recordset
    Set RsStockCab = New ADODB.Recordset
    Set RsStockDet = New ADODB.Recordset
    Set RsProducto = New ADODB.Recordset
    
    csql = "SELECT F5CODPRO FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
    If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
    RsMovAlmacen.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If RsMovAlmacen.EOF Then
        '''''''''''''''''CREA UNO NUEVO ''''''''''''''''''
        SSQL = "INSERT INTO IF6ALMA(F2CODALM,F5CODPRO,F6STOCKMAX,F6STOCKMIN,F6STOCKACT,F6STOCKLOG) VALUES('" & pcodalm & "','" & pcodpro & "','0.00','0.00','0.00','0.00')"
        cnn_dbbancos.Execute SSQL
    End If
    RsMovAlmacen.Close
    
    wmes = Month(CVDate(PFECVAL))
    
    '----------- VUELVE A BUSCAR ------------------
    csql = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
    If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
    RsMovAlmacen.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsMovAlmacen.EOF Then
        RsMovAlmacen.MoveFirst
        If Left(PNUMVAL, 1) = "I" Then
            If swingre = 1 Then 'si es un ingreso al costo promedio
                pvalpro = "" & Format(RsMovAlmacen.Fields("f5cospro"), "#0.000")
                pvaldol = "" & Format(RsMovAlmacen.Fields("f5cosprod"), "#0.000")
            Else
                If (Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro) <> 0 Then
                    If pvalpro <> 0 And pvaldol <> 0 Then
                        If RsMovAlmacen.Fields("f5cospro") <> 0# And RsMovAlmacen.Fields("f5cosprod") <> 0# Then
                            Cospro = Val("" & Format(((Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.000")) * Val(Format(RsMovAlmacen.Fields("f5cospro"), "#0.000"))) + (Format(pcanpro * pvalpro, "#0.000"))) / (Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro), "#0.000"))
                            Cosprod = Val("" & Format(((Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.000")) * Val(Format(RsMovAlmacen.Fields("f5cosprod"), "#0.000"))) + (Format(pcanpro * pvaldol, "#0.000"))) / (Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro), "#0.000"))
                        Else
                            Cospro = Val("" & Format(((Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.000")) * Format(pvalpro, "0.00")) + (Format(pcanpro * pvalpro, "0.000"))) / (Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.000")) + pcanpro), "0.000"))
                            Cosprod = Val("" & Format(((Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.000")) * Format(pvaldol, "0.00")) + (Format(pcanpro * pvaldol, "0.000"))) / (Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.000")) + pcanpro), "0.000"))
                        End If
                    Else
                        Cospro = Val("" & Format(RsMovAlmacen.Fields("f5cospro"), "0.000"))
                        Cosprod = Val("" & Format(RsMovAlmacen.Fields("f5cosprod"), "0.000"))
                    End If
                End If
            End If
            
            Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.000")) + pcanpro
            ing = Val(Format(RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")), "#0.000")) + (pcanpro * pvalpro)
            ingd = Val(Format(RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")), "#0.000")) + (pcanpro * Val("" & pvaldol))
            stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
            StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
            FecUlt = Format(PFECVAL, "dd/mm/yyyy")
            
            GRABAR_ACTUALIZACIONES wmes, pcodalm, pcodpro
            
        Else
            If pcanpro < 0 Then
                pcanpro = (pcanpro * -1)
            End If
            Cospro = Val("" & Format(RsMovAlmacen.Fields("f5cospro"), "0.000"))
            Cosprod = Val("" & Format(RsMovAlmacen.Fields("f5cosprod"), "0.000"))
            Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "0.000")) + pcanpro
            sal = Val(Format(RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")), "0.000")) + (Format(RsMovAlmacen.Fields("f5cospro"), "0.00") * pcanpro)
            sald = Val(Format(RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")), "0.000")) + (Format(RsMovAlmacen.Fields("f5cosprod"), "0.00") * pcanpro)
            stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.000")) - pcanpro
            FecUlt = CVDate(PFECVAL)
            
            GRABAR_ACTUALIZACIONES1 wmes, pcodalm, pcodpro
                
        End If
    End If
    RsMovAlmacen.Close
        
    SSQL = "SELECT F5moneda,F5STOCKACT FROM IF5PLA WHERE F5CODPRO = '" & pcodpro & "' OR F5CODFAB ='" & pcodpro & "'"
    If RsProducto.State = adStateOpen Then RsProducto.Close
    RsProducto.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProducto.EOF Then
        RsProducto.MoveFirst
        If Left(PNUMVAL, 1) = "I" Then
            precos = IIf(RsProducto.Fields("F5moneda") = "S", pvalpro, pvaldol)
            stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) + pcanpro
            SQL = "UPDATE IF5PLA SET F5PRECOS = " & precos & ",F5STOCKACT = " & stockact & " WHERE F5CODPRO = '" & pcodpro & "'"
        Else
            stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) - pcanpro
            SQL = "UPDATE IF5PLA SET F5STOCKACT = " & stockact & " WHERE F5CODPRO = '" & pcodpro & "'"
        End If
        cnn_dbbancos.Execute SQL
    End If
    RsProducto.Close

End Sub

Public Sub GRABAR_ACTUALIZACIONES1(Fmes As Integer, Alma As String, cod As String)

Dim amovs_arr(0 To 6) As a_grabacion

amovs_arr(0).campo = "F5HABM" & Format(Fmes, "00"): amovs_arr(0).valor = Habm: amovs_arr(0).TIPO = "N"
amovs_arr(1).campo = "F5SAL" & Format(Fmes, "00"): amovs_arr(1).valor = sal: amovs_arr(1).TIPO = "N"
amovs_arr(2).campo = "F5SALD" & Format(Fmes, "00"): amovs_arr(2).valor = sald: amovs_arr(2).TIPO = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).TIPO = "N"
amovs_arr(4).campo = "F6FECULT": amovs_arr(4).valor = CVDate(FecUlt): amovs_arr(4).TIPO = "F"
amovs_arr(5).campo = "F5COSPRO": amovs_arr(5).valor = Cospro: amovs_arr(5).TIPO = "N"
amovs_arr(6).campo = "F5COSPROD": amovs_arr(6).valor = Cosprod: amovs_arr(6).TIPO = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO amovs_arr(), "IF6ALMA", "M", 6, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub Actualizar_Almacenes(pcodalm As String, pcodpro As String, pcanpro As Double, PFECMOV As Variant, psoles As Double, pdolares As Double, ptipmov As String, pprecio As Double)
Dim csql, SSQL As String
    
Set RsProducto = New ADODB.Recordset
Set RsMovAlmacen = New ADODB.Recordset
'''''''''''CAMBIOS EN SQL
SSQL = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsMovAlmacen.EOF Then
    wmes = Format(Month(CVDate(PFECMOV)), "00")
    'CSQL = "SELECT * FROM IF5PLA WHERE F5CODPRO = '" & pcodpro & "'"
    'If RsProducto.State = adStateOpen Then RsProducto.Close
    'RsProducto.Open CSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    'If Not RsProducto.EOF Then
     '  If ptipmov = "S" Then
     '     stockact = Format(VAL(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) + pcanpro, "0.00")
     '  Else
     '     stockact = Format(VAL(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) - pcanpro, "0.00")
     '  End If
     '  SQL = "UPDATE IF5PLA SET F5STOCKACT = " & stockact & " WHERE F5CODPRO = '" & pcodpro & "'"
     '  cnn_dbbancos.Execute (SQL)
    'End If
    Debm = 0#: stockact = 0#: ing = 0#: ingd = 0#
    Habm = 0#: StockLog = 0#: sal = 0#: sald = 0#
    If ptipmov = "S" Then
        Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) - pcanpro
        ing = RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")) - psoles
        ingd = RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")) - pdolares
        GRABAR_ACTUALIZAALMA wmes, pcodalm, pcodpro
    Else
    '''''''''CORREGIR MES
        Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
        StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
        sal = RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")) - psoles
        sald = RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")) - pdolares
        GRABAR_ACTUALIZAALMA1 wmes, pcodalm, pcodpro
    End If

End If
End Sub

Sub GRABAR_ACTUALIZAALMA(Fmes As Integer, Alma As String, cod As String)
Dim amovs_arr(0 To 3) As a_grabacion

amovs_arr(0).campo = "F5DEBM" & Format(Fmes, "00"): amovs_arr(0).valor = Debm: amovs_arr(0).TIPO = "N"
amovs_arr(1).campo = "F5ING" & Format(Fmes, "00"): amovs_arr(1).valor = ing: amovs_arr(1).TIPO = "N"
amovs_arr(2).campo = "F5INGD" & Format(Fmes, "00"): amovs_arr(2).valor = ingd: amovs_arr(2).TIPO = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).TIPO = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO amovs_arr(), "IF6ALMA", "M", 3, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub GRABAR_ACTUALIZAALMA1(Fmes As Integer, Alma As String, cod As String)
Dim amovs_arr(0 To 4) As a_grabacion

amovs_arr(0).campo = "F5HABM" & Format(Fmes, "00"): amovs_arr(0).valor = Habm: amovs_arr(0).TIPO = "N"
amovs_arr(1).campo = "F5SAL" & Format(Fmes, "00"): amovs_arr(1).valor = sal: amovs_arr(1).TIPO = "N"
amovs_arr(2).campo = "F5SALD" & Format(Fmes, "00"): amovs_arr(2).valor = sald: amovs_arr(2).TIPO = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).TIPO = "N"
amovs_arr(4).campo = "F6STOCKLOG": amovs_arr(4).valor = StockLog: amovs_arr(4).TIPO = "N"


'------- ACTUALIZAR STOCKS
GRABA_REGISTRO amovs_arr(), "IF6ALMA", "M", 4, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub ImprimendoV(pcodalm As String, PNUMVAL As String, pcosto As Integer)
    
Dim ITEM, Fila As Integer
    
ITEM = 1
Fila = 1

Set RsStockCab = New ADODB.Recordset
Set RsStockDet = New ADODB.Recordset
Set RsProducto = New ADODB.Recordset
Set rsconsulta = New ADODB.Recordset

If RsStockCab.State = adStateOpen Then RsStockCab.Close
RsStockCab.Open "SELECT * FROM IF4VALES WHERE F2CODALM = '" & pcodalm & "' AND F4NUMVAL = '" & PNUMVAL & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockCab.EOF Then
    
    Printer.ScaleMode = 4
    TituloV PNUMVAL

    If RsStockDet.State = adStateOpen Then RsStockDet.Close
    RsStockDet.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & pcodalm & "' AND F4NUMVAL = '" & PNUMVAL & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockDet.EOF Then
        Fila = Fila + 6
        Do While Not RsStockDet.EOF
            If RsProducto.State = adStateOpen Then RsProducto.Close
            RsProducto.Open "SELECT * FROM IF5PLA WHERE F5CODPRO ='" & RsStockDet.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsProducto.EOF Then
                writexy Trim(RsStockDet.Fields("F5codpro")), Fila, 1, 0
                writexy Left(RsProducto.Fields("F5nompro"), 65), Fila, 11, 0
                writexy Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), Fila, 63, 2
                writexy Trim("" & RsProducto.Fields("f7codmed")), Fila, 80, 0
                Fila = Fila + 1
                If RsProducto.Fields("f5series") = "1" Then
                    If rsconsulta.State = adStateOpen Then rsconsulta.Close
                    SQL = "select * from if3series where f2codalm='" & pcodalm & "' and f4numval='" & PNUMVAL & "' and f5codpro='" & Trim(RsProducto.Fields("f5codpro")) & "' order by f3numser"
                    rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsconsulta.EOF Then
                        rsconsulta.MoveFirst
                        Do While Not rsconsulta.EOF
                            writexy ITEM & ".-  E.S.N. ==>" & rsconsulta.Fields("F3numser"), Fila, 16, 0
                            Fila = Fila + 1
                            rsconsulta.MoveNext
                            If Fila >= 60 Then
                                Printer.NewPage
                                TituloV PNUMVAL
                                Fila = 18
                            End If
                            ITEM = ITEM + 1
                        Loop
                    End If
                    Fila = Fila + 1
                End If
                ITEM = 1
                If Fila >= 60 Then
                   Printer.NewPage
                   TituloV PNUMVAL
                   Fila = 18
                End If
                RsProducto.MoveNext
            End If
            RsStockDet.MoveNext
        Loop
    End If

    Printer.Line (60, Fila + 6)-(88, Fila + 6)
    writexy "Almacén", Fila + 6, 68, 0
    
    Printer.Line (5, Fila + 6)-(33, Fila + 6)
    writexy "Solicitado Por", Fila + 6, 12, 0
    Printer.EndDoc
    
End If

End Sub

Sub TituloV(PNUMVAL As String)
Dim Fila As Integer
Set RsAlmacenes = New ADODB.Recordset
Set rsproveedor = New ADODB.Recordset
Set RsOrigen = New ADODB.Recordset
Set RsPartida = New ADODB.Recordset
    

    If RsOrigen.State = adStateOpen Then RsOrigen.Close
    RsOrigen.Open "SELECT * FROM SF1ORIGENES WHERE F1CODORI = '" & RsStockCab.Fields("F1CODORI") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    CabeceraV
    Printer.FontBold = True
    Printer.FontSize = 15
    writexy IIf(RsOrigen.Fields("F1TIPMOV") = "I", "VALE DE INGRESO", "VALE DE SALIDA"), 6, 34, 0
    Printer.FontBold = False
    Printer.FontSize = 9
    
    Fila = 11
    Printer.FontBold = True

    If Trim(nomcentro) <> "" Then
        writexy "C.COSTO:", 10, 5, 0
        writexy Trim(nomcentro), 10, 15, 0
    End If

    writexy "Nº VALE:", Fila, 5, 0
    writexy PNUMVAL, Fila, 15, 0
    
    writexy "Almacén:", Fila, 60, 0
    
    If Not IsNull(RsStockCab.Fields("F2codalm")) Then
        If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
        RsAlmacenes.Open "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & RsStockCab.Fields("F2CODALM") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsAlmacenes.EOF Then
            writexy Trim(RsAlmacenes.Fields("f2codalm") & " - " & Mid(RsAlmacenes.Fields("F2nomalm"), 1, 21)), Fila, 70, 0
        End If
        
    End If

    Fila = Fila + 1
    
    If Not IsNull(RsStockCab.Fields("F2codprov")) Then
        If rsproveedor.State = adStateOpen Then rsproveedor.Close
        rsproveedor.Open "SELECT * FROM EF2PROVEEDORES WHERE F2NEWRUC = '" & RsStockCab.Fields("F2CODPROV") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsproveedor.EOF Then
            writexy "Proveedor:", Fila, 5, 0
            writexy Trim(rsproveedor.Fields("f2codprov") & " - " & rsproveedor.Fields("F2nomprov")), Fila, 15, 0
        End If
        
    End If
    
    If Not IsNull(RsStockCab.Fields("F2codpar")) Then
        If RsStockCab.Fields("f1codori") = "XT0" Or RsStockCab.Fields("f1codori") = "XT1" Then
            If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
            RsAlmacenes.Open "SELECT * FROM EF2ALMACENES", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsAlmacenes.Find "F2codalm = '" & RsStockCab.Fields("F2codpar") & "'"
            If Not RsAlmacenes.EOF Then
                If RsStockCab.Fields("f1codori") = "XT0" Then
                   writexy "ORIGEN  : ", Fila, 60, 0
                Else
                   writexy "DESTINO : ", Fila, 60, 0
                End If
                writexy UCase(RsAlmacenes.Fields("f2codalm") & "-" & RsAlmacenes.Fields("f2nomalm")), Fila, 70, 0
            End If
        Else
            
            If RsPartida.State = adStateOpen Then RsPartida.Close
            RsPartida.Open "SELECT * FROM EF2PARTIDAS", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsPartida.Find "F2codpar = '" & RsStockCab.Fields("F2codpar") & "'"
            If Not RsPartida.EOF Then
                writexy "DESTINO : ", Fila, 60, 0
                writexy UCase(RsPartida.Fields("f2nompar")), Fila, 70, 0
            End If
        End If
    Else
        If Val("" & RsStockCab.Fields("F4numord")) > 0 Then
            writexy "O/Compra : ", Fila, 60, 0
            writexy Format(RsStockCab.Fields("F4numord"), "000000"), Fila, 70, 0
        End If
    End If


    Printer.FontBold = False
    writexy "Concepto:", Fila + 1, 5, 0
    writexy Trim(RsOrigen.Fields("F1nomori")), Fila + 1, 15, 0

    writexy "Documento:", Fila + 1, 60, 0
    writexy Trim(RsStockCab.Fields("F1coddoc") & " (" & RsStockCab.Fields("F4numdoc") & ") "), Fila + 1, 70, 0
    
    writexy "Fecha:", Fila + 2, 5, 0
    writexy Trim(Trim(cmes(Month(CVDate(RsStockCab.Fields("F4fecval"))))) & " " & Format(Day(CVDate(RsStockCab.Fields("F4fecval"))), "##") & ", de " & Year(CVDate(RsStockCab.Fields("F4fecval")))), Fila + 2, 15, 0
    
    'writexy "Moneda:", FILA + 2, 60, 0
    'writexy Trim(IIf(RSstocKcab.Fields("F4moneda") = "S", "Soles   T.C.: ", "Dólares   T.C.: ") & Format(RSstocKcab.Fields("F4tipcam"), "###,##0.00#")), FILA + 2, 70, 0

    Printer.FontBold = True
    Printer.Line (1, Fila + 3)-(90, Fila + 3)
    writexy "Código", Fila + 4, 1, 0
    writexy "Artículo", Fila + 4, 11, 0
    writexy "Cantidad", Fila + 4, 63, 0
    writexy "Uni.", Fila + 4, 80, 0
    'writexy "Costo Uni.", FILA + 4, 77, 0
    'writexy "Total", FILA + 4, 87, 0
    Printer.Line (1, Fila + 5)-(90, Fila + 5)
    Printer.FontBold = False
    
End Sub

Sub CabeceraV()
Dim csql As String
Set RsParain = New ADODB.Recordset

csql = "SELECT * FROM SF1PARAIN WHERE F1CODEMP = '" & wempresa & "'"
If RsParain.State = adStateOpen Then RsParain.Close
RsParain.Open csql, cnn_control, adOpenDynamic, adLockOptimistic
  
Printer.ScaleMode = 4
Printer.FontName = "Courier New" 'Printer.Fonts (RsParain.Fields("F1FONNAM"))
Printer.FontSize = 12
Printer.FontBold = True
writexy Trim("" & RsParain.Fields("f1nomemp")), 2, 1, 0
Printer.FontSize = 8
Printer.FontBold = False

writexy "Fecha: ", 2, 80, 0
writexy Format(Now, "dd/mm/yyyy"), 2, 88, 0
writexy "Intersys - Inventario", 3, 1, 0
writexy "Página: ", 3, 80, 0
writexy Format(Printer.Page, "###00"), 3, 88, 0

End Sub

Sub Reactualiza_Almacenes(pcodalm As String, pcodpro As String, pcanpro As Double, PFECMOV As Variant, psoles As Double, pdolares As Double, ptipmov As String, pprecio As Double)
Dim SSQL, csql As String
Set RsProducto = New ADODB.Recordset
Set RsMovAlmacen = New ADODB.Recordset

stockact = 0#: Debm = 0#: ing = 0#: ingd = 0#
StockLog = 0#: Habm = 0#: sal = 0#: sald = 0#
SQL = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsMovAlmacen.EOF Then
    wmes = Format(Month(CVDate(PFECMOV)), "00")
    SSQL = "SELECT F5STOCKACT FROM IF5PLA WHERE F5CODPRO = '" & pcodpro & "'"
    If RsProducto.State = adStateOpen Then RsProducto.Close
    RsProducto.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProducto.EOF Then
       If ptipmov = "S" Then
          stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) + pcanpro
       Else
          stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) - pcanpro
       End If
       csql = "UPDATE IF5PLA SET F5STOCKACT = " & stockact & " WHERE F5CODPRO = '" & pcodpro & "'"
       cnn_dbbancos.Execute (csql)
    End If
    RsProducto.Close
    
    If ptipmov = "S" Then
        Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) - pcanpro
        ing = Val(Format(RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")), "0.00")) - psoles
        ingd = Val(Format(RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")), "0.00")) - pdolares
        GRABAR_ACTUALIZAALMA wmes, pcodalm, pcodpro
    Else
        Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
        StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
        sal = Val(Format(RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")), "0.00")) - psoles
        sald = Val(Format(RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")), "0.00")) - pdolares
        GRABAR_ACTUALIZAALMA1 wmes, pcodalm, pcodpro

    End If
    
End If
End Sub

Public Sub BASE_TEMPORAL(Base As String)
Dim CON As String
Set Temp = New ADODB.Connection

CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & Base & "; Persist Security Info=False"
Temp.Open CON

End Sub

