Attribute VB_Name = "mod_transcont"
Option Explicit

Public wdg1        As Integer
Public wdg2        As Integer
Public wdg3        As Integer
Public wdg4        As Integer
Public wdg5        As Integer

Public Sub SUM_ANALISIS(pref As Variant, pcuenta As Variant, pdh As Variant, psoles As Variant, pdolar As Variant, pdetalle As Variant, pitem As Variant, pcheque As Variant, ptipdoc As Variant, ptc As Variant, porigen As String, pfecha As Variant, pcompro As String, pmes As String, pconexion As ADODB.Connection)
Dim rsplan          As New ADODB.Recordset
Dim rscf9saldo      As New ADODB.Recordset
Dim rscf9cta        As New ADODB.Recordset
Dim amovs_cab()     As a_grabacion
Dim cmes9           As String
Dim nsaldo          As Double
Dim nsaldod         As Double

    If Len(Trim(pref)) <> 0 Then
        If rsplan.State = adStateOpen Then rsplan.Close
        rsplan.Open "SELECT * FROM CF5PLA WHERE F5CODCTA='" & pcuenta & "'", cnn_dbtabla, adOpenDynamic, adLockOptimistic
        If Not rsplan.EOF Then
            If rscf9saldo.State = adStateOpen Then rscf9saldo.Close
            rscf9saldo.Open "SELECT * FROM CF9SALDO WHERE F9CODCTA='" & pcuenta & "' AND F9NROREF='" & pref & "'", pconexion, adOpenDynamic, adLockOptimistic
            If Not rscf9saldo.EOF Then
                If rsplan.Fields("F5MONEDA") & "" = "D" Then
                    If Val(Format(rscf9saldo.Fields("F9SALDOD"), "0.00")) = 0# And Len(Trim("" & rscf9saldo.Fields("F9MESD"))) > 0 And Val(rscf9saldo.Fields("F9MESD") & "") < Val(pmes) Then
                        pconexion.Execute ("UPDATE CF9SALDO SET F9MESD=' ' WHERE F9CODCTA='" & pcuenta & "' AND F9NROREF='" & pref & "'")
                    End If
                Else
                    If Val(Format(rscf9saldo.Fields("F9SALDO"), "0.00")) = 0 And Len(Trim("" & rscf9saldo.Fields("F9MES"))) > 0 And Val(rscf9saldo.Fields("F9MES") & "") < Val(pmes) Then
                        pconexion.Execute ("UPDATE CF9SALDO SET F9MES=' ' WHERE F9CODCTA='" & pcuenta & "' AND F9NROREF='" & pref & "'")
                    End If
                End If
            Else
                ReDim amovs_cab(0 To 2) As a_grabacion
                amovs_cab(0).campo = "F9CODCTA": amovs_cab(0).valor = pcuenta: amovs_cab(0).TIPO = "T"
                amovs_cab(1).campo = "F9NROREF": amovs_cab(1).valor = pref: amovs_cab(1).TIPO = "T"
                amovs_cab(2).campo = "F9MESI": amovs_cab(2).valor = pmes: amovs_cab(2).TIPO = "T"
                GRABA_REGISTRO amovs_cab(), "CF9SALDO", "A", 2, pconexion, ""
            End If
            rscf9saldo.Close
            
            If rscf9saldo.State = adStateOpen Then rscf9saldo.Close
            rscf9saldo.Open "SELECT * FROM CF9SALDO WHERE F9CODCTA='" & pcuenta & "' AND F9NROREF='" & pref & "'", pconexion, adOpenDynamic, adLockOptimistic
            If Not rscf9saldo.EOF Then
                
                ReDim amovs_cab(0 To 3) As a_grabacion
                
                cmes9 = "": nsaldo = 0#: nsaldod = 0#
                If rsplan.Fields("f5moneda") & "" = "D" Then
                    If Val(Format(rscf9saldo.Fields("f9saldod"), "0.00")) = 0 And Len(Trim(rscf9saldo.Fields("f9mesd") & "")) > 0 Then
                        amovs_cab(0).campo = "F9MESD": amovs_cab(0).valor = " ": amovs_cab(0).TIPO = "T"
                        cmes9 = " "
                    End If
                Else
                    If Val(Format(rscf9saldo.Fields("f9saldo"), "0.00")) = 0# And Len(Trim("" & rscf9saldo.Fields("f9mes"))) > 0 Then
                        amovs_cab(0).campo = "F9MES": amovs_cab(0).valor = " ": amovs_cab(0).TIPO = "T"
                        cmes9 = " "
                    End If
                End If
                                
                If pdh = "D" Then
                    amovs_cab(1).campo = "F9SALDO": amovs_cab(1).valor = Val(Format(Val("" & rscf9saldo.Fields("F9SALDO")) + psoles, "0.00")): amovs_cab(1).TIPO = "N"
                    amovs_cab(2).campo = "F9SALDOD": amovs_cab(2).valor = Val(Format(Val("" & rscf9saldo.Fields("F9SALDOD")) + pdolar, "0.00")): amovs_cab(2).TIPO = "N"
                    nsaldo = Val(Format(Val("" & rscf9saldo.Fields("F9SALDO")) + psoles, "0.00"))
                    nsaldod = Val(Format(Val("" & rscf9saldo.Fields("F9SALDOD")) + pdolar, "0.00"))
                Else
                    amovs_cab(1).campo = "F9SALDO": amovs_cab(1).valor = Val(Format(Val("" & rscf9saldo.Fields("F9SALDO")) - psoles, "0.00")): amovs_cab(1).TIPO = "N"
                    amovs_cab(2).campo = "F9SALDOD": amovs_cab(2).valor = Val(Format(Val("" & rscf9saldo.Fields("F9SALDOD")) - pdolar, "0.00")): amovs_cab(2).TIPO = "N"
                    nsaldo = Val(Format(Val("" & rscf9saldo.Fields("F9SALDO")) - psoles, "0.00"))
                    nsaldod = Val(Format(Val("" & rscf9saldo.Fields("F9SALDOD")) - pdolar, "0.00"))
                End If
                
                If rsplan.Fields("f5moneda") & "" = "D" Then
                    If Val(Format(nsaldod, "0.00")) = 0# And Len(Trim(cmes9)) = 0 Then
                        amovs_cab(3).campo = "F9MESD": amovs_cab(3).valor = pmes: amovs_cab(3).TIPO = "T"
                    End If
                Else
                    If Val(Format(nsaldo, "0.00")) = 0# And Len(Trim(cmes9)) = 0 Then
                        amovs_cab(3).campo = "F9MES": amovs_cab(3).valor = pmes: amovs_cab(3).TIPO = "T"
                    End If
                End If
                
                GRABA_REGISTRO amovs_cab(), "CF9SALDO", "M", 3, pconexion, "F9CODCTA='" & pcuenta & "' AND F9NROREF='" & pref & "'"
            
            End If
            rscf9saldo.Close
            
            ReDim amovs_cab(0 To 12) As a_grabacion
            amovs_cab(0).campo = "F9ORIGEN": amovs_cab(0).valor = porigen: amovs_cab(0).TIPO = "T"
            amovs_cab(1).campo = "F9FECHA": amovs_cab(1).valor = pfecha: amovs_cab(1).TIPO = "F"
            amovs_cab(2).campo = "F9NROREF": amovs_cab(2).valor = pref: amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F9COMPRO": amovs_cab(3).valor = pcompro: amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F9DEBHAB": amovs_cab(4).valor = pdh: amovs_cab(4).TIPO = "T"
            amovs_cab(5).campo = "F9DETALL": amovs_cab(5).valor = pdetalle: amovs_cab(5).TIPO = "T"
            amovs_cab(6).campo = "F9ELEMEN": amovs_cab(6).valor = pitem: amovs_cab(6).TIPO = "T"
            amovs_cab(7).campo = "F9CHEQUE": amovs_cab(7).valor = pcheque: amovs_cab(7).TIPO = "N"
            amovs_cab(8).campo = "F9CODCTA": amovs_cab(8).valor = pcuenta: amovs_cab(8).TIPO = "T"
            amovs_cab(9).campo = "F9IMPORTE": amovs_cab(9).valor = psoles: amovs_cab(9).TIPO = "N"
            amovs_cab(10).campo = "F9IMPORTED": amovs_cab(10).valor = pdolar: amovs_cab(10).TIPO = "N"
            amovs_cab(11).campo = "F9TIPDOC": amovs_cab(11).valor = ptipdoc: amovs_cab(11).TIPO = "T"
            amovs_cab(12).campo = "F9TIPCAMBD": amovs_cab(12).valor = ptc: amovs_cab(12).TIPO = "N"
            GRABA_REGISTRO amovs_cab(), "CF9CTA", "A", 12, pconexion, ""
            
        End If
        rsplan.Close
    End If

End Sub

Public Sub REG_MAYOR(pconexion As ADODB.Connection, pmes As String, pconexion_temp As ADODB.Connection, pnomtabla As String)
Dim ccompro             As String
Dim ctipograba          As String
Dim cwhere              As String
Dim rstemporal          As New ADODB.Recordset
Dim rscabecera          As New ADODB.Recordset
Dim amovs_cab()         As a_grabacion

    If rstemporal.State = adStateOpen Then rstemporal.Close
    rstemporal.Open "SELECT * FROM " & pnomtabla & "", pconexion_temp, adOpenDynamic, adLockOptimistic
    If Not rstemporal.EOF Then
        rstemporal.MoveFirst
        Do While Not rstemporal.EOF
            cwhere = ""
            ccompro = rstemporal.Fields("f3compro") & ""
            If rscabecera.State = adStateOpen Then rscabecera.Close
            rscabecera.Open "SELECT * FROM CF4TCO" & pmes & " WHERE F4COMPRO='" & ccompro & "'", pconexion, adOpenDynamic, adLockOptimistic
            If Not rscabecera.EOF Then
                ctipograba = "M"
                cwhere = "F4COMPRO='" & ccompro & "'"
            Else
                ctipograba = "A"
            End If
            
            ReDim amovs_cab(0 To 8) As a_grabacion
            amovs_cab(0).campo = "F4COMPRO": amovs_cab(0).valor = rstemporal.Fields("F3COMPRO") & "": amovs_cab(0).TIPO = "T"
            amovs_cab(1).campo = "F4ORIGEN": amovs_cab(1).valor = "" & rstemporal.Fields("F3ORIGEN"): amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F4OBRA": amovs_cab(2).valor = "" & rstemporal.Fields("F3OBRA"): amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F4MONEDA": amovs_cab(3).valor = "" & rstemporal.Fields("F3MONEDA"): amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F4FECHA": amovs_cab(4).valor = rstemporal.Fields("F3FCHOPR"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F4TIPCAMBD": amovs_cab(5).valor = Val("" & rstemporal.Fields("F3TIPCAMBD")): amovs_cab(5).TIPO = "N"
            
            If Not rscabecera.EOF Then
                If rstemporal.Fields("f3debhab") = "D" Then
                    amovs_cab(6).campo = "F4TOTDEB": amovs_cab(6).valor = Val("" & rscabecera.Fields("F4TOTDEB")) + Val("" & rstemporal.Fields("F3IMPORTE")): amovs_cab(6).TIPO = "N"
                    amovs_cab(7).campo = "F4TOTDEBD": amovs_cab(7).valor = Val("" & rscabecera.Fields("F4TOTDEBD")) + Val("" & rstemporal.Fields("F3IMPORTED")): amovs_cab(7).TIPO = "N"
                Else
                    amovs_cab(6).campo = "F4TOTHAB": amovs_cab(6).valor = Val("" & rscabecera.Fields("F4TOTHAB")) + Val("" & rstemporal.Fields("F3IMPORTE")): amovs_cab(6).TIPO = "N"
                    amovs_cab(7).campo = "F4TOTHABD": amovs_cab(7).valor = Val("" & rscabecera.Fields("F4TOTHABD")) + Val("" & rstemporal.Fields("F3IMPORTED")): amovs_cab(7).TIPO = "N"
                End If
                amovs_cab(8).campo = "F4NUMELE": amovs_cab(8).valor = Format(Val("" & rscabecera.Fields("F4NUMELE")) + 1, "####"): amovs_cab(8).TIPO = "T"
            Else
                If rstemporal.Fields("f3debhab") = "D" Then
                    amovs_cab(6).campo = "F4TOTDEB": amovs_cab(6).valor = Val("" & rstemporal.Fields("F3IMPORTE")): amovs_cab(6).TIPO = "N"
                    amovs_cab(7).campo = "F4TOTDEBD": amovs_cab(7).valor = Val("" & rstemporal.Fields("F3IMPORTED")): amovs_cab(7).TIPO = "N"
                Else
                    amovs_cab(6).campo = "F4TOTHAB": amovs_cab(6).valor = Val("" & rstemporal.Fields("F3IMPORTE")): amovs_cab(6).TIPO = "N"
                    amovs_cab(7).campo = "F4TOTHABD": amovs_cab(7).valor = Val("" & rstemporal.Fields("F3IMPORTED")): amovs_cab(7).TIPO = "N"
                End If
                amovs_cab(8).campo = "F4NUMELE": amovs_cab(8).valor = Format(1, "####"): amovs_cab(8).TIPO = "T"
            End If
            
            GRABA_REGISTRO amovs_cab(), "CF4TCO" & pmes, ctipograba, 8, pconexion, cwhere
            
            rstemporal.MoveNext
            If rstemporal.EOF Then Exit Do
        Loop
        
        pconexion.Execute ("UPDATE CF4TCO" & pmes & " SET F4CUADRES='1' WHERE F4TOTDEB<>F4TOTHAB")
        pconexion.Execute ("UPDATE CF4TCO" & pmes & " SET F4CUADRES='0' WHERE F4TOTDEB=F4TOTHAB")
        
        pconexion.Execute ("UPDATE CF4TCO" & pmes & " SET F4CUADRED='1' WHERE F4TOTDEBD<>F4TOTHABD")
        pconexion.Execute ("UPDATE CF4TCO" & pmes & " SET F4CUADRES='0' WHERE F4TOTDEBD=F4TOTHABD")
        
    End If
    rstemporal.Close
   
End Sub

Private Sub REG_SALDOS(pconexion As ADODB.Connection, pmes As String, pconexion_temp As ADODB.Connection, pnomtabla As String)
Dim gr                  As Integer
Dim grado               As Integer
Dim ccuenta             As String
Dim ccta                As String
Dim rsplan              As New ADODB.Recordset
Dim rsplan2             As New ADODB.Recordset
Dim amovs_cab(0 To 3)   As a_grabacion
Dim cwhere              As String
Dim rstemporal          As New ADODB.Recordset
Dim nsoles              As Double
Dim ndolar              As Double

    If rstemporal.State = adStateOpen Then rstemporal.Close
    rstemporal.Open "SELECT * FROM " & pnomtabla & "", pconexion_temp, adOpenDynamic, adLockOptimistic
    If Not rstemporal.EOF Then
        rstemporal.MoveFirst
        Do While Not rstemporal.EOF
            ccuenta = rstemporal.Fields("F5CODCTA") & ""
            nsoles = Val(rstemporal.Fields("F3IMPORTE") & "")
            ndolar = Val(rstemporal.Fields("F3IMPORTED") & "")
            If rsplan.State = adStateOpen Then rsplan.Close
            rsplan.Open "SELECT * FROM CF5PLA WHERE F5CODCTA='" & ccuenta & "'", pconexion, adOpenDynamic, adLockOptimistic
            If Not rsplan.EOF Then
                grado = rsplan.Fields("f5grdcta") + 1
                gr = 1
                Do While grado > gr
                    Select Case gr
                        Case 1: ccta = Mid(ccuenta, 1, wdg1)
                        Case 2: ccta = Mid(ccuenta, 1, wdg2)
                        Case 3: ccta = Mid(ccuenta, 1, wdg3)
                        Case 4: ccta = Mid(ccuenta, 1, wdg4)
                        Case 5: ccta = Mid(ccuenta, 1, wdg5)
                    End Select
                    ccta = Trim(ccta)
                    If rsplan2.State = adStateOpen Then rsplan2.Close
                    rsplan2.Open "SELECT * FROM CF5PLA WHERE F5CODCTA='" & ccta & "'", pconexion, adOpenDynamic, adLockOptimistic
                    If Not rsplan2.EOF Then
                        cwhere = "F5CODCTA='" & ccta & "'"
                        If rstemporal.Fields("F3DEBHAB") & "" = "D" Then
                            amovs_cab(0).campo = "F5DEBM" & pmes: amovs_cab(0).valor = Val("" & rsplan2.Fields("F5DEBM" & pmes)) + nsoles: amovs_cab(0).TIPO = "N"
                            amovs_cab(1).campo = "F5DEBM99": amovs_cab(1).valor = Val("" & rsplan2.Fields("F5DEBM99")) + nsoles: amovs_cab(1).TIPO = "N"
                            amovs_cab(2).campo = "F5DEBDM" & pmes: amovs_cab(2).valor = Val("" & rsplan2.Fields("F5DEBDM" & pmes)) + ndolar: amovs_cab(2).TIPO = "N"
                            amovs_cab(3).campo = "F5DEBDM99": amovs_cab(3).valor = Val("" & rsplan2.Fields("F5DEBDM99")) + ndolar: amovs_cab(3).TIPO = "N"
                        Else
                            amovs_cab(0).campo = "F5HABM" & pmes: amovs_cab(0).valor = Val("" & rsplan2.Fields("F5HABM" & pmes)) + nsoles: amovs_cab(0).TIPO = "N"
                            amovs_cab(1).campo = "F5HABM99": amovs_cab(1).valor = Val("" & rsplan2.Fields("F5HABM99")) + nsoles: amovs_cab(1).TIPO = "N"
                            amovs_cab(2).campo = "F5HABDM" & pmes: amovs_cab(2).valor = Val("" & rsplan2.Fields("F5HABDM" & pmes)) + ndolar: amovs_cab(2).TIPO = "N"
                            amovs_cab(3).campo = "F5HABDM99": amovs_cab(3).valor = Val("" & rsplan2.Fields("F5HABDM99")) + ndolar: amovs_cab(3).TIPO = "N"
                        End If
                        GRABA_REGISTRO amovs_cab(), "CF5PLA", "M", 3, pconexion, cwhere
                    End If
                    gr = gr + 1
                    rsplan2.Close
                Loop
            End If
            rsplan.Close
            rstemporal.MoveNext
        Loop
    End If
    rstemporal.Close

End Sub

Public Sub TRANSFIERE_ASIENTOS(pconexion As ADODB.Connection, pconexion_temp As ADODB.Connection, pmes As String, pconexion_sistema As ADODB.Connection, psistema As String, ptablacab As String)
Dim rstemporal          As New ADODB.Recordset
Dim amovs_cab()         As a_grabacion

    rstemporal.Open "SELECT * FROM CONTABLE", pconexion_temp, adOpenDynamic, adLockOptimistic
    If Not rstemporal.EOF Then
        rstemporal.MoveFirst
        Do While Not rstemporal.EOF
            ReDim amovs_cab(0 To 21) As a_grabacion
            amovs_cab(0).campo = "F3COMPRO": amovs_cab(0).valor = rstemporal.Fields("F3COMPRO") & "": amovs_cab(0).TIPO = "T"
            amovs_cab(1).campo = "F3PROAME": amovs_cab(1).valor = rstemporal.Fields("F3PROAME") & "": amovs_cab(1).TIPO = "T"
            amovs_cab(2).campo = "F3ELEMEN": amovs_cab(2).valor = rstemporal.Fields("F3ELEMEN") & "": amovs_cab(2).TIPO = "T"
            amovs_cab(3).campo = "F3ORIGEN": amovs_cab(3).valor = rstemporal.Fields("F3ORIGEN") & "": amovs_cab(3).TIPO = "T"
            amovs_cab(4).campo = "F3FCHOPR": amovs_cab(4).valor = rstemporal.Fields("F3FCHOPR"): amovs_cab(4).TIPO = "F"
            amovs_cab(5).campo = "F3DETALL": amovs_cab(5).valor = rstemporal.Fields("F3DETALL") & "": amovs_cab(5).TIPO = "T"
            amovs_cab(6).campo = "F5CODCTA": amovs_cab(6).valor = rstemporal.Fields("F5CODCTA") & "": amovs_cab(6).TIPO = "T"
            amovs_cab(7).campo = "F3CODGAS": amovs_cab(7).valor = rstemporal.Fields("F3CODGAS") & "": amovs_cab(7).TIPO = "N"
            amovs_cab(8).campo = "F3CHEQUE": amovs_cab(8).valor = rstemporal.Fields("F3CHEQUE") & "": amovs_cab(8).TIPO = "T"
            amovs_cab(9).campo = "F3NROREF": amovs_cab(9).valor = rstemporal.Fields("F3NROREF") & "": amovs_cab(9).TIPO = "T"
            amovs_cab(10).campo = "F3IMPORTE": amovs_cab(10).valor = Val(rstemporal.Fields("F3IMPORTE") & ""): amovs_cab(10).TIPO = "N"
            amovs_cab(11).campo = "F3IMPORTED": amovs_cab(11).valor = Val(rstemporal.Fields("F3IMPORTED") & ""): amovs_cab(11).TIPO = "N"
            amovs_cab(12).campo = "F3MONEDA": amovs_cab(12).valor = rstemporal.Fields("F3MONEDA") & "": amovs_cab(12).TIPO = "T"
            amovs_cab(13).campo = "F3TIPCAMBD": amovs_cab(13).valor = Val(rstemporal.Fields("F3TIPCAMBD") & ""): amovs_cab(13).TIPO = "N"
            amovs_cab(14).campo = "F3TIPDOC": amovs_cab(14).valor = rstemporal.Fields("F3TIPDOC") & "": amovs_cab(14).TIPO = "T"
            amovs_cab(15).campo = "F3DEBHAB": amovs_cab(15).valor = rstemporal.Fields("F3DEBHAB") & "": amovs_cab(15).TIPO = "T"
            amovs_cab(16).campo = "F3COSTO": amovs_cab(16).valor = rstemporal.Fields("F3COSTO") & "": amovs_cab(16).TIPO = "T"
            amovs_cab(17).campo = "F3CTABANC": amovs_cab(17).valor = rstemporal.Fields("F3CTABANC") & "": amovs_cab(17).TIPO = "N"
            amovs_cab(18).campo = "F3MOVBANC": amovs_cab(18).valor = rstemporal.Fields("F3MOVBANC") & "": amovs_cab(18).TIPO = "N"
            amovs_cab(19).campo = "F3DESTINO": amovs_cab(19).valor = rstemporal.Fields("F3DESTINO") & "": amovs_cab(19).TIPO = "T"
            amovs_cab(20).campo = "F3ANNOCOMP": amovs_cab(20).valor = rstemporal.Fields("F3ANNOCOMP") & "": amovs_cab(20).TIPO = "T"
            amovs_cab(21).campo = "F3REGCOMP": amovs_cab(21).valor = rstemporal.Fields("F3REGCOMP") & "": amovs_cab(21).TIPO = "T"
            
            GRABA_REGISTRO amovs_cab(), "CF3MOV" & pmes, "A", 21, pconexion, ""
            
            SUM_ANALISIS rstemporal.Fields("F3NROREF") & "", rstemporal.Fields("F5CODCTA"), rstemporal.Fields("F3DEBHAB"), rstemporal.Fields("F3IMPORTE"), rstemporal.Fields("F3IMPORTED"), rstemporal.Fields("F3DETALL"), rstemporal.Fields("F3ELEMEN"), rstemporal.Fields("F3CHEQUE"), rstemporal.Fields("F3TIPDOC"), rstemporal.Fields("F3TIPCAMBD"), rstemporal.Fields("F3ORIGEN") & "", rstemporal.Fields("F3FCHOPR"), rstemporal.Fields("F3COMPRO") & "", pmes, cnn_analisis
            
            If psistema = "B" Then  '---- Bancos
                pconexion_sistema.Execute ("UPDATE " & ptablacab & " SET F4CONTABLE='*' WHERE F4CODCTA= " & rstemporal.Fields("F3CTABANC") & " AND F4NUMMOV = " & rstemporal.Fields("F3MOVBANC") & "")
            End If
            If psistema = "C" Then  '---- Compras
                pconexion_sistema.Execute ("UPDATE " & ptablacab & " SET F4CONTABLE='*' WHERE F4ANNO= '" & rstemporal.Fields("F3ANNOCOMP") & "' AND F4MESMOV = '" & Mid(rstemporal.Fields("F3REGCOMP") & "", 1, 2) & "' AND F4NUMMOV = '" & Mid(rstemporal.Fields("F3REGCOMP") & "", 3, 7) & "'")
            End If
            If psistema = "R" Then  '---- Comprobantes de Retención
                pconexion_sistema.Execute ("UPDATE " & ptablacab & " SET TRANSFERIDO='*' WHERE SERIE= '" & Mid(rstemporal.Fields("F3COMP_RETENCION") & "", 1, 3) & "' AND NUM_DOCUMENTO = '" & Mid(rstemporal.Fields("F3COMP_RETENCION") & "", 5, 7) & "'")
            End If
            rstemporal.MoveNext
        Loop
        REG_MAYOR pconexion, pmes, pconexion_temp, "CONTABLE"
        REG_SALDOS cnn_dbtabla, pmes, pconexion_temp, "CONTABLE"
        
    End If
    
End Sub
