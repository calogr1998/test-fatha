Attribute VB_Name = "mod_regcompras"
Option Explicit

Public wcontacnt        As String
Public wexcel           As String
Public dbcntcont        As DAO.Database
Public tbcntcont        As DAO.Recordset
Public mes              As String
Public wanno            As String
Public gcodppp          As String
Public dbbanco          As DAO.Database
Public tbtalon          As DAO.Recordset
Public dbempresa        As DAO.Database
Public Tbproveedor      As DAO.Recordset
Public dbcompras        As DAO.Database
Public TbCabRegis_new   As New ADODB.Recordset
Public TbCabRegis       As DAO.Recordset
'Public TbOfiRegis       As DAO.Recordset
Public TbOfiRegis       As New ADODB.Recordset
Public TbDetRegis       As New ADODB.Recordset
Public tbmes            As DAO.Recordset
Public dbcomtabla       As DAO.Database
Public tbcomtab         As DAO.Recordset
Public tbdepa           As DAO.Recordset
Public tbcambios        As DAO.Recordset
Public tbgastos         As DAO.Recordset
Public TbDocumento      As DAO.Recordset
Public dbconta          As DAO.Database
Public tbcf5pla         As DAO.Recordset
Public tbcf5costo       As DAO.Recordset
Public dbcentros        As DAO.Database
Public tbcentros        As DAO.Recordset
Public dbcontrol        As DAO.Database
Public tbparametro1     As DAO.Recordset
Public Wnuevo           As Integer
Public wf1viscod        As String * 1
Public wocompra         As String
Public wf1tipdoc_asoc   As String
Public wf1inggasto      As String
Public wbancos          As String * 1
Public wf1traslado      As String
Public gctapag          As String * 1
Public wf1formatov      As String
Public wf1numera        As String * 1
Public wf1trasoc        As String
Public DbCtaPag         As DAO.Database
Public gorden_cs        As String
Public temp_deta        As DAO.Recordset
Public tbcabcta         As DAO.Recordset
Public tbdetcta         As DAO.Recordset
Public gcorre           As Double
Public gretenc          As Double
Public gfonavi          As Double
Public gnummov          As String
Public wingobra         As String
Public GCODGAS          As String
Public gcodcon          As String
Public gnomgas          As String
Public gcodcen          As String
Public wcentro          As String
Public worigen          As String
Public wctaigv          As String
Public wctaotros        As String
Public wredsuma         As String
Public wvisualiza_cod   As String  'visualizar la columna cod.,fab,ambos
Public wredresta        As String
Public wctaret          As String
Public wctafon          As String
Public llampro          As Integer
Public gcueppp          As String
Public gsegppp          As String
Public gmoneda          As String
Public gcodord          As Long
Public gcodprov         As String
Public gcoddepa         As String
Public gcodcta          As String * 3
Public wayudamov        As Integer
Public rucprv           As String
Public XITEM            As String
Public xgasto           As String
Public dbtempcomp       As DAO.Database
Public Tbtemp_regis     As DAO.Recordset
Public F4MES            As String
Public wf1formato       As String
Public xmodo            As String
Public xmes             As String
Public xmoneda          As String
Public xpro             As String
Public xtipo            As String
Public dbcomtab         As DAO.Database
Public grucprov         As String
Public tbregisdoc       As DAO.Recordset
Public tbregismov       As DAO.Recordset
Public gmonedacta       As String
Public gmondes          As String
Public dbplancta        As DAO.Database
Public tbplancta        As DAO.Recordset
Public dbtemp           As DAO.Database
Public NCTA             As Integer
Public gnomcen          As String
Public dbbancos         As DAO.Database
Public tbcta            As DAO.Recordset
Public TBBANCO          As DAO.Recordset
Public tbcheques        As DAO.Recordset
Public TBBF4TCO         As DAO.Recordset
Public dbmovis          As DAO.Database
Public nmovbanco        As String
Public xcaja            As Integer
Public xmovcaja         As Integer
Public xmescaja         As Integer
Public xctacaja         As Integer
Public gcodtip          As String
Public xtalon           As Integer
Public DbInventa        As DAO.Database
Public tbcabtmp         As DAO.Recordset
Public TbDetImport      As DAO.Recordset
Public TBPRODUCTO       As DAO.Recordset
Public TbCabImport      As DAO.Recordset
Public dvinvtemp        As DAO.Database
Public tbcf9saldo       As DAO.Recordset
Public tbcf9cta         As DAO.Recordset
Public tbmovcf4         As DAO.Recordset
Public tbcontrol        As DAO.Recordset
Public dbtemconta       As DAO.Database
Public dbmovconta       As DAO.Database
Public tbmovcf3         As DAO.Recordset
Public dbtabla          As DAO.Database
Public dbanalisis       As DAO.Database
Public wdcto            As String
Public gtipo            As String
Public TbDetOrdenes     As New ADODB.Recordset
Public TbCabOrdenes     As New ADODB.Recordset
Public gnomcon          As String
Public gcodnom          As String
Public wgastos          As String
Public wcodigo          As String
Public tbconta          As DAO.Recordset
Public wctacont         As String
Public wnomctacont      As String
Public LLAMADA          As String
Public des_grupo        As String
Public cod_grupo        As String
Public wf1renovacion    As String
Public wdestino         As String
Public wf1cnting        As String
Public wload_usuario    As String
Public wcodigos         As String
Public tbgrupos         As DAO.Recordset
Public tbmovim          As DAO.Recordset
Public tbcompra         As DAO.Recordset
Public TBCOMPraDOC      As DAO.Recordset
Public dbcompra         As DAO.Database
Public dbmovim          As DAO.Database
Public dbgrupos         As DAO.Database
Public WF2CODGAS        As String
Public wf1agente        As String
Public Sw_Graba_Registro As Boolean
Public ntc              As Double

Public Function dev_mes(mes)
Dim nmes    As Integer
Dim cmes    As String
   
   nmes = Val(mes)
   Select Case mes
      Case 0: cmes = "Apertura"
      Case 1: cmes = "Enero"
      Case 2: cmes = "Febrero"
      Case 3: cmes = "Marzo"
      Case 4: cmes = "Abril"
      Case 5: cmes = "Mayo"
      Case 6: cmes = "Junio"
      Case 7: cmes = "Julio"
      Case 8: cmes = "Agosto"
      Case 9: cmes = "Setiembre"
      Case 10: cmes = "Octubre"
      Case 11: cmes = "Noviembre"
      Case 12: cmes = "Diciembre"
      Case 13: cmes = "Cierre 1"
      Case 14: cmes = "Cierre 2"
   End Select
   dev_mes = cmes

End Function

Public Function dev_mes_ingles(mes)
Dim nmes    As Integer
Dim cmes    As String
   
   nmes = Val(mes)
   Select Case mes
      Case 0: cmes = "Apertura"
      Case 1: cmes = "January"
      Case 2: cmes = "February"
      Case 3: cmes = "March"
      Case 4: cmes = "April"
      Case 5: cmes = "May"
      Case 6: cmes = "June"
      Case 7: cmes = "July"
      Case 8: cmes = "August"
      Case 9: cmes = "September"
      Case 10: cmes = "October"
      Case 11: cmes = "November"
      Case 12: cmes = "December"
      Case 13: cmes = "Cierre 1"
      Case 14: cmes = "Cierre 2"
   End Select
   dev_mes_ingles = cmes

End Function


Public Sub WriteXY_texto(pdata As Variant, PFILA As Variant, pcolu As Variant, ptipo As Integer)

    Print #1, pdata
    
End Sub

Public Sub resta_bf5(mes As String, nmonto As Double, ccta As Integer, cdestino As String, wrutabancos As String)
Dim dbmovis     As DAO.Database
Dim tbcuenta    As DAO.Recordset

On Error GoTo error_grababf5

Act_bf5:
    Set dbmovis = OpenDatabase(wrutabancos & "\db_tabla.MDB")
    Set tbcuenta = dbmovis.OpenRecordset("BF5PLA")
    tbcuenta.Index = "IDCODIGO"
    
    tbcuenta.Seek "=", ccta
    If Not tbcuenta.NoMatch Then
        tbcuenta.Edit
        If cdestino = "I" Then
            tbcuenta.Fields("f5debm" & mes) = tbcuenta.Fields("f5debm" & mes) - nmonto
            tbcuenta.Fields("f5debm99") = tbcuenta.Fields("f5debm99") - nmonto
        Else
            tbcuenta.Fields("f5habm" & mes) = tbcuenta.Fields("f5habm" & mes) - nmonto
            tbcuenta.Fields("f5habm99") = tbcuenta.Fields("f5habm99") - nmonto
        End If
        tbcuenta.Fields("f5saldo99") = tbcuenta.Fields("f5debm99") - tbcuenta.Fields("f5habm99")
        tbcuenta.Update
    End If
    tbcuenta.Close
    dbmovis.Close
    
    Exit Sub

error_grababf5:
    If Err = 3186 Then
        MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos"
        tbcuenta.Close
        dbmovis.Close
        Resume Act_bf5
    End If
    If Err = 3163 Then
        MsgBox Error(Err) & " Verifique la base de datos."
        Resume Next
    End If

End Sub

Public Sub actualizar_bf3(mes As String, wrutabancos As String, bf3() As String)
Dim dbmovis     As DAO.Database
Dim tbdetalle   As DAO.Recordset
 
On Error GoTo error_grababf3

Act_bf3:
    Set dbmovis = OpenDatabase(wrutabancos & "db_tabla.MDB")
    Set tbdetalle = dbmovis.OpenRecordset("BF3MOV" & Format(mes, "00"))
    tbdetalle.Index = "BF3CTA" & Format(mes, "00")
    
    tbdetalle.Seek "=", Val(bf3(5)), Val(bf3(1)), Val(bf3(2))
    If Not tbdetalle.NoMatch Then
        tbdetalle.Edit
    Else
        tbdetalle.AddNew
    End If
    tbdetalle.Fields("nummov") = Val(bf3(1))
    tbdetalle.Fields("elemen") = Val(bf3(2))
    tbdetalle.Fields("tipmov") = bf3(3)
    tbdetalle.Fields("proame") = bf3(4)
    tbdetalle.Fields("codcta") = bf3(5)
    tbdetalle.Fields("codgto") = left(bf3(6), 12)
    tbdetalle.Fields("orden") = bf3(7)
    tbdetalle.Fields("fecdis") = Format(bf3(8), "dd/mm/yyyy")
    tbdetalle.Fields("concepto") = bf3(9)
    tbdetalle.Fields("docum") = Val(bf3(10))
    tbdetalle.Fields("parcial") = Val(bf3(11))
    tbdetalle.Fields("tcambio") = Val(bf3(12))
    tbdetalle.Fields("f3moneda") = bf3(13)
    tbdetalle.Fields("f3debhab") = bf3(14)
    tbdetalle.Fields("tipdocu") = bf3(15)
    tbdetalle.Fields("codigto") = bf3(16)
    tbdetalle.Fields("fechdet") = Format(bf3(17), "dd/mm/yyyy")
    tbdetalle.Fields("ser_doc") = Val(bf3(18))
    tbdetalle.Fields("tcrendi") = Val(bf3(19))
    tbdetalle.Fields("f3costo") = bf3(20)
    tbdetalle.Fields("reg_com") = bf3(21)
    tbdetalle.Update
    
    tbdetalle.Close
    dbmovis.Close

    Exit Sub

error_grababf3:
    If Err = 3186 Then
        MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos"
        tbdetalle.Close
        dbmovis.Close
        Resume Act_bf3
    End If
    If Err = 3163 Then
        MsgBox Error(Err) & " Verifique la base de datos."
        Resume Next
    End If

End Sub

Public Sub actualizar_bf4(mes As String, wrutabancos As String, bf4() As Variant)
Dim dbmovis     As DAO.Database
Dim tbcabecera  As DAO.Recordset
 
On Error GoTo ERROR_GRABA

Act_bf4:
    Set dbmovis = OpenDatabase(wrutabancos & "db_tabla.mdb")
    Set tbcabecera = dbmovis.OpenRecordset("BF4TCO" & Format(mes, "00"))
    tbcabecera.Index = "BF4TCO" & Format(mes, "00")
    
    tbcabecera.Seek "=", Val(bf4(1)), Val(bf4(2))
    If Not tbcabecera.NoMatch Then
        tbcabecera.Edit
    Else
        tbcabecera.AddNew
    End If
    tbcabecera.Fields("f4codcta") = Val(bf4(1))
    tbcabecera.Fields("f4nummov") = Val(bf4(2))
    tbcabecera.Fields("f4numele") = Val(bf4(3))
    tbcabecera.Fields("f4tipmov") = bf4(4)
    tbcabecera.Fields("f4numdoc") = Val(bf4(5))
    tbcabecera.Fields("f4detal") = bf4(6)
    tbcabecera.Fields("f4feccam") = Format(bf4(7), "dd/mm/yyyy")
    tbcabecera.Fields("f4fecgir") = Format(bf4(8), "dd/mm/yyyy")
    tbcabecera.Fields("f4voucher") = Val(bf4(9))
    tbcabecera.Fields("f4numtal") = Val(bf4(10))
    tbcabecera.Fields("f4destino") = bf4(11)
    tbcabecera.Fields("f4total") = Format(bf4(12), "0.00")
    tbcabecera.Fields("f4fecadd") = Format(bf4(13), "dd/mm/yyyy")
    tbcabecera.Fields("f4giradoa") = bf4(14)
    tbcabecera.Fields("f4docpend") = bf4(15)
    tbcabecera.Fields("f4tipcamb") = Val(bf4(16))
    tbcabecera.Fields("f4observa1") = bf4(17)
    tbcabecera.Fields("f4observa2") = bf4(18)
    tbcabecera.Fields("f4inciden") = bf4(19)
    tbcabecera.Fields("f4codprov") = bf4(20)
    tbcabecera.Fields("f4orden") = bf4(21)
    tbcabecera.Update
    
    tbcabecera.Close
    dbmovis.Close

    Exit Sub

ERROR_GRABA:
    If Err = 3186 Then
        MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos"
        tbcabecera.Close
        dbmovis.Close
        Resume Act_bf4
    End If
    If Err = 3163 Then
        MsgBox Error(Err) & " Verifique la base de datos."
        Resume Next
    End If

End Sub

Public Sub actualizar_bf6(bf6() As Variant, wrutabancos As String)
Dim dbmovis     As DAO.Database
Dim tbcheques   As DAO.Recordset
Dim tbtalon     As DAO.Recordset
Dim flag        As Integer
 
On Error GoTo error_grababf6

Act_bf6:
    Set dbmovis = OpenDatabase(wrutabancos & "db_tabla.MDB")
    Set tbcheques = dbmovis.OpenRecordset("BF6CHQ")
    tbcheques.Index = "CTATALCHQ"
    
    Set tbtalon = dbmovis.OpenRecordset("BF7TALON")
    tbtalon.Index = "BF7TALON"
    
    tbcheques.Seek "=", Val(bf6(1)), Val(bf6(2)), Val(bf6(3))
    If Not tbcheques.NoMatch Then
        flag = False
        tbcheques.Edit
    Else
        flag = True
        tbcheques.AddNew
    End If
    tbcheques.Fields("codcta") = Val(bf6(1))
    tbcheques.Fields("codtal") = Val(bf6(2))
    tbcheques.Fields("numdoc") = Val(bf6(3))
    tbcheques.Fields("tipmov") = bf6(4)
    tbcheques.Fields("nummov") = Val(bf6(5))
    tbcheques.Fields("total") = bf6(6)
    tbcheques.Fields("fecgir") = Format(bf6(7), "dd/mm/yyyy")
    tbcheques.Fields("codprov") = bf6(8)
    tbcheques.Fields("giradoa") = bf6(9)
    tbcheques.Update
    
    If flag = True Then
        tbtalon.Seek "=", Val(bf6(1)), Val(bf6(2))
        If Not tbtalon.NoMatch Then
            tbtalon.Edit
            tbtalon.Fields("chegir") = tbtalon.Fields("chegir") + 1
            tbtalon.Fields("nultcheque") = Val(bf6(3))
            tbtalon.Update
            
        End If
    End If
    
    tbtalon.Close
    tbcheques.Close
    dbmovis.Close

    Exit Sub

error_grababf6:
    If Err = 3186 Then
        MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos"
        tbtalon.Close
        tbcheques.Close
        dbmovis.Close
        Resume Act_bf6
    End If
    If Err = 3163 Then
        MsgBox Error(Err) & " Verifique la base de datos."
        Resume Next
    End If

End Sub

Public Sub suma_bf5(mes As String, nmonto As Double, ccta As Integer, cdestino As Variant, wrutabancos As String)
Dim dbmovis     As DAO.Database
Dim tbcuenta    As DAO.Recordset
 
On Error GoTo error_grabasbf5

Act_sumbf5:
    Set dbmovis = OpenDatabase(wrutabancos & "db_tabla.MDB")
    Set tbcuenta = dbmovis.OpenRecordset("BF5PLA")
    tbcuenta.Index = "IDCODIGO"
    
    tbcuenta.Seek "=", ccta
    If Not tbcuenta.NoMatch Then
        tbcuenta.Edit
        If cdestino = "I" Then
            tbcuenta.Fields("f5debm" & Format(mes, "00")) = tbcuenta.Fields("f5debm" & Format(mes, "00")) + nmonto
            tbcuenta.Fields("f5debm99") = tbcuenta.Fields("f5debm99") + nmonto
        Else
            tbcuenta.Fields("f5habm" & Format(mes, "00")) = tbcuenta.Fields("f5habm" & Format(mes, "00")) + nmonto
            tbcuenta.Fields("f5habm99") = tbcuenta.Fields("f5habm99") + nmonto
        End If
        tbcuenta.Fields("f5saldo99") = tbcuenta.Fields("f5debm99") - tbcuenta.Fields("f5habm99")
        tbcuenta.Update
    End If
    tbcuenta.Close
    dbmovis.Close
    
    Exit Sub

error_grabasbf5:
    If Err = 3186 Then
        MsgBox "La base de datos está bloqueada por otro usuario. Espere unos segundos"
        tbcuenta.Close
        dbmovis.Close
        Resume Act_sumbf5
    End If
    If Err = 3163 Then
        MsgBox Error(Err) & " Verifique la base de datos."
        Resume Next
    End If

End Sub

Public Function genera_mov(NCTA As Variant, mes As String, wrutabancos As String)
Dim nmov    As Integer
Dim dbbaseb As DAO.Database
Dim tbtab   As DAO.Recordset
 
    Set dbbaseb = OpenDatabase(wrutabancos & "\db_tabla.mdb")
    Set tbtab = dbbaseb.OpenRecordset("SELECT * FROM BF4TCO" & Format(mes, "00") & " WHERE F4CODCTA =" & NCTA & " ORDER BY F4NUMMOV")
   
    If tbtab.RecordCount > 0 Then
        tbtab.MoveLast
        nmov = tbtab.Fields("F4NUMMOV") + 1
    Else
        nmov = 1
    End If
    tbtab.Close
    dbbaseb.Close
    genera_mov = nmov

End Function
Public Sub CargaOrdenDeCompra(NumeroDeOrden As String)

Dim Af As New ADOFunctions
Dim rs As New ADODB.Recordset
Dim RsCTR_COM As New ADODB.Recordset
Dim RsPago As New ADODB.Recordset
Dim RSCONSULTA As New ADODB.Recordset
Dim nAnchoHoja As Double
Dim X As Object

Set X = New Acr_OrdenCompra
With X
'With Acr_OrdenCompra
        'jcg13 Set cImgInfo = New cImageInfo
   ' MsgBox "Acr_OrdenC_Otros"
        CargaVariables
        .flddirec1.Text = wf1direc1
        .FldTelf.Text = "Teléfono: " & wtelefono & " // Fax: " & wfax
        '.flddirec2.Text = wf1direc2
        nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
        .fldruc.Text = "R.U.C. " & wrucempresa
'        If FileExists(App.Path & "\" & wrucempresa & ".jpg") = True Then
'            .fldempresa.Visible = False
'            .ImageLogo.Visible = True
'            .ImageLogo.Picture = LoadPicture(App.Path & "\" & wrucempresa & ".jpg")
'            With cImgInfo
'                .ReadImageInfo App.Path & "\" & wrucempresa & ".jpg"
'
'                X.ImageLogo.Height = 850
'                X.ImageLogo.Width = 850 * .Width / .Height
'            End With
'            .ImageLogo.top = 0
'            .ImageLogo.left = 0
'        Else
            .fldempresa.Visible = True
            .ImageLogo.Visible = False
            .fldempresa.Text = wnomcia
'        End If
        
        '.IGV.Caption = wigv
        GOC = NumeroDeOrden
        If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
        csql = "SELECT A.F4NUMORD,A.F4ESTNUL, A.F4CODSOLICITUD, B.F2NOMPROV,  A.F4CODSOL, A.F4CONTACTO,  B.F2TELPROV,  B.F2FAXPROV, " & _
              "A.F4FECEMI,A.F4FECVEN, A.F4MONTO, B.F2DIRPROV, A.F4FORPAG,A.F4IGV,A.F4BASIMP,A.F4RND,A.F4OBSERVA, " & _
              "A.F4PLAZO_ENTREGA,A.F4LUGAR_ENTREGA,A.F4FECENT,A.F4OBSERVA,A.F4CODPRV,A.F4TIPMON,A.F4REFERE,A.F4TIPCAM,A.F4FECGRA,A.F4USEGRA,A.F4FECMOD,A.F4USEMOD " & _
              " FROM IF4ORDEN AS A, EF2PROVEEDORES AS B WHERE A.F4CODPRV=B.F2NEWRUC AND A.F4NUMORD = '" & GOC & _
              "' AND A.F4LOCAL='1' ORDER BY A.F4NUMORD DESC"
    
        Set RSCONSULTA = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
        If Not RSCONSULTA.EOF Then
            If RSCONSULTA!F4TIPMON = "S" Then
               ' .LblSubF.Caption = "Sub.Total " & "S/"
                '.LblIgvF.Caption = "I.G.V. " & "S/"
                .LblTotF.Caption = "Total " & "S/"
            Else
                '.LblSubF.Caption = "Sub.Total " & "US$"
                '.LblIgvF.Caption = "I.G.V. " & "US$"
                .LblTotF.Caption = "Total " & "US$"
    
            End If
            .fldEstado.Visible = True
            .fldEstado.Alignment = ddTXRight
            Select Case RSCONSULTA!F4ESTNUL & ""
            Case "N"
                .fldEstado.Text = "SIN APROBACIÓN"
                .fldEstado.ForeColor = vbBlue
            Case "S"
                .fldEstado.Text = "ANULADO"
                .fldEstado.ForeColor = vbRed
            Case "A"
                .fldEstado.Text = "APROBADO"
                .fldEstado.ForeColor = vbBlack
            Case "P"
                .fldEstado.Text = "PENDIENTE DE APROBACIÓN"
                .fldEstado.ForeColor = vbGreen
            Case "A"
                .fldEstado.Text = "APROBADO"
                .fldEstado.ForeColor = vbBlack
            Case "R"
                .fldEstado.Text = "RECHAZADO"
                .fldEstado.ForeColor = vbMagenta
            Case Else
                .fldEstado.Text = ""
            End Select
            '.F4NUMORD.Text = ("" & rsconsulta.Fields("F4NUMORD"))
            .LblTitle.Caption = "ORDEN DE COMPRA"
            .LblNroOC.Caption = "N° " & RSCONSULTA.Fields("F4NUMORD")
            '.F4CODSOLICITUD.Text = "" & RSCONSULTA.Fields("F4CODSOLICITUD")
            .LblCrea.Caption = "Creación Usuario: " & RSCONSULTA.Fields("F4USEGRA") & " (" & RSCONSULTA.Fields("F4fecGRA") & ")"
            .LblModifica.Caption = "Último Usuario: " & RSCONSULTA.Fields("F4USEmod") & " (" & RSCONSULTA.Fields("F4fecmod") & ")"
            .F2NOMPROV.Text = "" & RSCONSULTA.Fields("F2NOMPROV")
            .F2DIRPROV.Text = "" & RSCONSULTA.Fields("F2DIRPROV")
            .F2CONTACTO.Text = "" & RSCONSULTA.Fields("F4CONTACTO")
            .F2CONTACTO.Visible = True
            .F2TELPROV.Text = "" & RSCONSULTA.Fields("F2TELPROV")
            .F2FAXPROV.Text = "" & RSCONSULTA.Fields("F2FAXPROV")
            .F4FECEMI.Text = "" & RSCONSULTA.Fields("F4FECEMI")
            .FldFchEntrega = "" & Format(RSCONSULTA.Fields("F4FECent"), "dd/mm/yyyy")
            .FldTipCam.Text = Format(Val("" & RSCONSULTA.Fields("F4tipcam")), "0.000")
'            .F4IGV.Text = Format("" & RSCONSULTA.Fields("F4IGV"), "###,###,##0.00")
'
'            .F4BASIMP.Text = Format("" & RSCONSULTA.Fields("F4BASIMP"), "###,###,###,##0.00")
'            '.Field4.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            .FldObservaAll.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
'            .F4MONTO.Text = Format("" & RSCONSULTA.Fields("F4MONTO"), "###,###,###,##0.00")
            '.F4PLAZO.Text = rsconsulta.Fields("F4PLAZO_ENTREGA") & ""
            Rem NSE .F3FECEN.Text = DateAdd("d", Val(rsconsulta.Fields("F4PLAZO_ENTREGA") & ""), .F4FECEMI.Text)
            '.F3FECEN.Text = Format(RSCONSULTA.Fields("F4FECENT"), "DD/MM/YYYY")
            '.F4NOTA.Text = "" & RSCONSULTA.Fields("F4OBSERVA")
            .F4CODPRV.Text = "" & RSCONSULTA.Fields("F4CODPRV")
            .FldSon.Text = CADENANUM(Val("" & RSCONSULTA.Fields("F4MONTO")), "" & RSCONSULTA.Fields("F4TIPMON"), "")
            '.referencia.Text = "" & Txt_Referencia.Text
            '.solicitado.Text = "" & pnlnomsoli.Caption
        
            csql = "SELECT F2DESPAG FROM EF2FORPAG WHERE F2FORPAG = '" & RSCONSULTA.Fields("F4FORPAG") & "'"
            Set RsPago = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
            If Not RsPago.EOF Then
                .F2DESPAG.Text = "" & RsPago.Fields("F2DESPAG")
            End If
            RsPago.Close
            'If RsCTR_COM.State = adStateOpen Then RsCTR_COM.Close
            'RsCTR_COM.Open "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'", cnn_ctrcom, adOpenDynamic, adLockOptimistic
            .LblFirma1.Caption = ""
            .LblCargo1.Caption = ""
            .LblFirma2.Caption = ""
            .LblCargo2.Caption = ""
            csql = "SELECT * FROM PARAM_COM  WHERE F1CODEMP= '" & wempresa & "'"
            Set RsCTR_COM = Af.OpenSQLForwardOnly(csql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CTRCOM.MDB" & ";Persist Security Info=False")
            If Not RsCTR_COM.EOF Then
                .LblFirma1.Caption = "" & RsCTR_COM.Fields("F1EMITIDO_OC")
                .LblCargo1.Caption = "" & RsCTR_COM.Fields("F1OBSGEN_OC")
                .LblFirma2.Caption = "" & RsCTR_COM.Fields("F1EMITIDO_OCI")
                .LblCargo2.Caption = "" & RsCTR_COM.Fields("F1OBSGEN_OCi")
                .Refresh
            End If
            RsCTR_COM.Close
            If Len(Trim(.LblFirma1.Caption)) > 0 And Len(Trim(.LblFirma2.Caption)) > 0 Then
                .LblFirma1.Visible = True
                .LblFirma2.Visible = True
                .LblFirma1.Width = nAnchoHoja / 2
                .LblCargo1.Width = nAnchoHoja / 2
                .LblFirma1.left = 0
                .LblCargo1.left = 0
                .LblFirma2.Width = nAnchoHoja / 2
                .LblCargo2.Width = nAnchoHoja / 2
                .LblFirma2.left = .LblFirma1.Width
                .LblCargo2.left = .LblCargo1.Width
            ElseIf Len(Trim(.LblFirma1.Caption)) > 0 And Len(Trim(.LblFirma2.Caption)) = 0 Then
                .LblFirma1.Visible = True
                .LblFirma2.Visible = False
                .LblFirma1.Width = nAnchoHoja
                .LblCargo1.Width = nAnchoHoja
                .LblFirma1.left = 0
                .LblCargo1.left = 0
                
            End If
            
            .F4OBSFECHA.Text = "" & RSCONSULTA.Fields("F4PLAZO_ENTREGA")
            .REMITIR.Text = RSCONSULTA.Fields("F4LUGAR_ENTREGA") & ""
            '.F4EMITIR.Text = "" & wnomcia
            '.f4emitir2.Text = "" & wdireccion
'            .f4emitir3.Text = "Ph: " & wtelefono & "  Fax: " & wfax
            
            If rs.State = 1 Then rs.Close
            
            Set rs = Nothing
            .Refresh
            '.LBLFIRMA.Caption = rs(0)  ' Trim("" & pnlnomsoli.Caption)
            '.lblcargo.Caption = traerCampo("EF2USERS", "F2CARGO", "F2CODUSER", wusuario & "")
            '.lblempresa.Caption = wnomcia
        End If
        .DataControl1.ConnectionString = cnn_dbbancos
        
        csql = "SELECT EF7MEDIDAS.F7SIGMED AS F3MEDIDA, CENTROS.F3CODCLI, CENTROS.PO, IF3ORDEN.* "
        csql = csql & "FROM (IF3ORDEN LEFT JOIN CENTROS ON IF3ORDEN.F3CENCOS = CENTROS.F3COSTO) "
        csql = csql & "LEFT JOIN EF7MEDIDAS ON IF3ORDEN.UNIDAD = EF7MEDIDAS.F7CODMED "
        csql = csql & "WHERE (((IF3ORDEN.F4NUMORD)='" & NumeroDeOrden & "') AND ((IF3ORDEN.F4LOCAL)='1')) "
        csql = csql & "order by IF3ORDEN.item"

        .DataControl1.Source = csql

        .fldEstado.Visible = False
        .Caption = "ORDEN DE COMPRA NACIONAL"
        RSCONSULTA.Close
    End With
'        If Not X Is Nothing Then
'            Load ReporteChildFalse
'
'            ReporteChildFalse.Caption = "ORDEN DE COMPRA " & NumeroDeOrden
'            Set ReporteChildFalse.arvPreview.object = X
'
'            ReporteChildFalse.Show 1
'            Unload ReporteChildFalse
'            Set ReporteChildFalse = Nothing
'
'        End If
End Sub


Public Sub Actualiza_Log(CadSql As String, conexion As String)
    On Error Resume Next
    
    Dim NomFile As String, StrLine As String
    
    Rem SK ADD:
    Dim NomFileUsuario As String, StrLineUsuario As String, strCadSqlUsuario As String
    Dim intNumSlot As Integer
    
    NomFile = wrutabancos & "\Control_Plus_Logistica_" & dev_mes(Month(Date)) & "_" & Year(Date) & ".log"
    
    Rem SK ADD:
    NomFileUsuario = wrutatemp & "\Control_Plus_Logistica_" & ComputerName & "_" & dev_mes(Month(Date)) & "_" & Year(Date) & ".log"
    strCadSqlUsuario = CadSql
    
    intNumSlot = FreeFile
    
    If (InStr(UCase(conexion), "DB_TABLA") > 0 Or InStr(UCase(conexion), "DB_BANCOS") > 0) Then
        '------------------------------------------------------------------------------------------------------------
        Close #intNumSlot
    
        Open Trim(NomFile) For Append As #intNumSlot
        
        '***genera sql
        CadSql = UCase(Replace(CadSql, "'", "|"))
        
        StrLine = "<Fecha Hora:" & Format(Now, "MM/DD/YYYY HH:MM:SS AM/PM") & ">" & "<Usuario:" & wusuario & ">" & "<Pc:" & ComputerName & ">" & "<Sentencia:" & CadSql & ">"
        
        Print #intNumSlot, StrLine
        
        '------------------------------------------------------------------------------------------------------------
        
        Rem SK ADD:
        If Dir(NomFileUsuario, vbArchive) = vbNullString Then
            Close #intNumSlot
            
            Open Trim(NomFileUsuario) For Append As #intNumSlot
            
            StrLineUsuario = "FECHA|USUARIO|NOMBREPC|SENTENCIAEJECUTADA"
            
            Print #intNumSlot, StrLineUsuario
        End If
        
        Close #intNumSlot
        
        Open Trim(NomFileUsuario) For Append As #intNumSlot
        
        StrLineUsuario = Format(Now, "MM/DD/YYYY HH:MM:SS AM/PM") & "|" & wusuario & "|" & ComputerName & "|" & strCadSqlUsuario
        
        Print #intNumSlot, StrLineUsuario
        
        Close #intNumSlot
    End If
    
    Exit Sub
End Sub




Public Function SeleccionaEnImageComboLeft(ByVal DatoBuscado As String, ByVal NombreDelCombo As ImageCombo)
Dim ComboImage As ComboItem
Dim I As Integer
For I = 1 To NombreDelCombo.ComboItems.Count
    Set ComboImage = NombreDelCombo.ComboItems(I)
    If UCase(left(ComboImage.Text, Len(Trim(DatoBuscado)))) = UCase(DatoBuscado) Then
        ComboImage.Selected = True
        Exit For
    Else
    End If
Next
End Function


Public Function SeleccionaEnImageComboRight(ByVal DatoBuscado As String, ByVal NombreDelCombo As ImageCombo)
Dim ComboImage As ComboItem
Dim I As Integer
For I = 1 To NombreDelCombo.ComboItems.Count
    Set ComboImage = NombreDelCombo.ComboItems(I)
    If UCase(right(ComboImage.Text, Len(Trim(DatoBuscado)))) = UCase(DatoBuscado) Then
        ComboImage.Selected = True
        Exit For
    Else
    End If
Next
End Function
Public Function SeleccionaEnComboTipDoc(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
Dim I As Integer
For I = 0 To NombreDelCombo.ListCount - 1
    If left(right(NombreDelCombo.List(I), 5), 2) = DatoBuscado Then
        NombreDelCombo.ListIndex = I
        Exit For
    Else
    End If
Next
End Function

Public Function SeleccionaEnComboForPag(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
Dim I As Integer
For I = 0 To NombreDelCombo.ListCount - 1
    If left(right(NombreDelCombo.List(I), 4), 3) = DatoBuscado Then
        NombreDelCombo.ListIndex = I
        Exit For
    Else
    End If
Next
End Function
Public Sub ExportaRecordsetToExcel(pRs As ADODB.Recordset, pFileName As String)
                         

Dim iRowIndex As Integer, avRows As Variant, ErrorOccured As Boolean
Dim iFieldCount As Integer, objExcel As Excel.Application, objTemp As Object
Dim iColIndex As Integer, iRecordCount As Integer



On Error GoTo errSub
    If Len(Trim(pFileName)) = 0 Then Exit Sub
    With pRs
        .MoveFirst
        'avRows = .GetRows()
        iRecordCount = .RecordCount
        iFieldCount = .Fields.Count
        'Set objExcel = GetObject(, "excel.application")
        Set objExcel = CreateObject("Excel.Application")
        'objExcel.Visible = True
        objExcel.Workbooks.Add

        Set objTemp = objExcel           'Ensure excel remains visible

        If Val(objExcel.Application.Version) >= 8 Then
        '    Set objExcel = objExcel.ActiveSheet
        End If

        iRowIndex = 1

        'Place Name of the fields
        For iColIndex = 0 To iFieldCount - 1
            With objExcel.Cells(iRowIndex, iColIndex + 1)
                .Value = pRs.Fields(iColIndex).Name
                With .Font
                    
                    .Bold = True
                End With
            End With
        Next iColIndex

    End With



    With objExcel
        pRs.MoveFirst
        iRowIndex = 2
        Do While Not pRs.EOF
            
            For iColIndex = 0 To iFieldCount - 1
                If IsNumeric(pRs.Fields(iColIndex).Value & "") Then
                    If Val(pRs.Fields(iColIndex).Value & "") = 0 Then
                        .Cells(iRowIndex, iColIndex + 1) = ""
                    Else
                        If pRs.Fields(iColIndex).Type = 202 Then
                            .Cells(iRowIndex, iColIndex + 1) = "'" & pRs.Fields(iColIndex).Value
                        Else
                            .Cells(iRowIndex, iColIndex + 1) = Val(pRs.Fields(iColIndex).Value)
                        End If
                    End If
                Else
                    If pRs.Fields(iColIndex).Type = 7 Then
                        .Cells(iRowIndex, iColIndex + 1) = Format(Day(Format(pRs.Fields(iColIndex).Value & "", "dd/mm/yyyy")), "00") & "/" & Format(Month(Format(pRs.Fields(iColIndex).Value & "", "dd/mm/yyyy")), "00") & "/" & Year(Format(pRs.Fields(iColIndex).Value & "", "dd/mm/yyyy"))
                    Else
                        .Cells(iRowIndex, iColIndex + 1) = pRs.Fields(iColIndex).Value & ""
                    End If
                End If
                
            Next iColIndex
            iRowIndex = iRowIndex + 1
            pRs.MoveNext
        Loop
        
             
        
        .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
        

        
        
    End With
    'wFileName = "C:\" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "_" & Mid(Time, 1, 2) & Mid(Time, 4, 2) & ".xls"
    If Len(Trim(pFileName)) > 0 Then
        If FileExists(pFileName) = True Then
            Kill pFileName
        End If
    End If
    objExcel.ActiveWorkbook.SaveAs (pFileName)
    
    objExcel.Application.Quit
    Set objTemp = Nothing
    Set objExcel = Nothing
    
    MsgBox "Se generó el archivo: " & vbCrLf & pFileName, vbInformation, wnomcia

Exit Sub

errSub:
    
    If Err.Number = 0 Then
        Resume Next
        ErrorOccured = True
    End If

End Sub

Public Sub CargaVariables()
Dim Af As New ADOFunctions
Dim rscontrol As New ADODB.Recordset
    csql = "SELECT * FROM SF1PARAM WHERE F1CODEMP ='" & wempresa & "'"
    Set rscontrol = Af.OpenSQLForwardOnly(csql, StrConexControl)
    If rscontrol.RecordCount > 0 Then
        wdireccion = Trim("" & rscontrol.Fields("F1DIREMP"))
        wDistrito = Trim("" & rscontrol.Fields("F1DISTRITO"))
        wtelefono = Trim("" & rscontrol.Fields("F1LOGEMP"))
        wfax = Trim("" & rscontrol.Fields("F1FAX"))
        wPais = "Perú"
        wwigv = Val("" & rscontrol.Fields("f1igv"))
        wFob = Val("" & rscontrol.Fields("f1fob"))
        wDesaduana = Val("" & rscontrol.Fields("f1desaduanaje"))
        wAdela = Val("" & rscontrol.Fields("f1adelanto"))
        
        wLocal = "" & rscontrol.Fields("F4LOCAL")
    End If
    rscontrol.Close
End Sub

Public Function FileExists(strPath As String) As Boolean ' verifica si un archivo existe
    On Error Resume Next
 
    If Len(strPath) < 4 Then
        FileExists = False
        Exit Function
    End If
 
    FileExists = IIf(Dir(strPath, _
    vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "", True, False)
End Function


Public Function ObtenerCampoWhere(tabla As String, campo As String, campoCom As String, valor As String, TipoDeComparacion As String, ConexionDeBaseDeDatos As ADODB.Connection, condicion As String) As String
On Error GoTo CapturaError
Dim cad As String
Dim rst As New Recordset

If ConexionDeBaseDeDatos.State = 0 Then ConexionDeBaseDeDatos.Open

If TipoDeComparacion = "F" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = #" & Format(valor, "mm/dd/yyyy") & "# " & condicion
ElseIf TipoDeComparacion = "T" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = '" & valor & "' " & condicion
ElseIf TipoDeComparacion = "N" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = " & valor & " " & condicion
End If
If rst.State = 1 Then rst.Close
rst.Open cad, ConexionDeBaseDeDatos, 3, 1
ObtenerCampoWhere = ""
If Not rst.EOF And Not IsNull(rst.Fields(0)) Then ObtenerCampoWhere = rst.Fields(0)
Exit Function
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    If MsgBox("¿Desea intentar conectarse nuevamente?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
        Resume
    Else
        Exit Function
    End If
End Function
Public Sub GRABA_REGISTRO_DET(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As String, pwhere As String, parr_det() As Variant, pnumfilas As Integer, pvalores As String, pmes As String, pgraba_saldo As String)
On Error GoTo Error_Graba_Registro_Det
Dim CnSave As New ADODB.Connection
Dim I           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
Dim nfila       As Integer
    
        
    Sw_Graba_Registro = False
    
    CnSave.Open pconexion
    
    ccampos = "": cvalores = ""
    For nfila = 0 To pnumfilas
        For I = 0 To pcantidad
            If Mid(pvalores, I + 1, 1) = "1" Then '--- EL CAMPO ESTA ACTIVO PARA GRABARLO
                If ptipo = "A" Then
                    If Len(ccampos) = 0 Then
                        ccampos = ccampos & parreglo(I).campo
                        If parreglo(I).Tipo = "T" Then
                            cvalores = "'" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            'cvalores = parr_det(I, nfila)
                            If Not IsNumeric(parr_det(I, nfila)) Then
                                cvalores = "Null"
                            Else
                                cvalores = parr_det(I, nfila)
                            End If
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = "'" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = "'" & Format(parr_det(I, nfila), "yyyy-mm-dd") & "'"
                            Else
                                'cvalores = "CVDATE('" & parr_det(I, nfila) & "')"
                                If Not IsDate(parr_det(I, nfila)) Then
                                    cvalores = "Null"
                                Else
                                    cvalores = "CVDATE('" & parr_det(I, nfila) & "')"
                                End If
                            End If
                        End If
                    Else
                        ccampos = ccampos & "," & parreglo(I).campo
                        If parreglo(I).Tipo = "T" Then
                            cvalores = cvalores & ",'" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            'cvalores = cvalores & "," & parr_det(I, nfila) & ""
                            If Not IsNumeric(parr_det(I, nfila)) Then
                                cvalores = cvalores & ",Null"
                            Else
                                cvalores = cvalores & "," & 0 & parr_det(I, nfila) & ""
                            End If
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = cvalores & ",'" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & ",'" & Format(parr_det(I, nfila), "yyyy-mm-dd") & "'"
                            Else
                                'cvalores = cvalores & ",CVDATE('" & parr_det(I, nfila) & "')"
                                If Not IsDate(parr_det(I, nfila)) Then
                                    cvalores = cvalores & ",Null"
                                Else
                                    cvalores = cvalores & ",CVDATE('" & parr_det(I, nfila) & "')"
                                End If
                            End If
                        End If
                    End If
                End If
                If ptipo = "M" Then
                    If Len(Trim(cvalores)) = 0 Then
                        If parreglo(I).Tipo = "T" Then
                            cvalores = cvalores & parreglo(I).campo & "='" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            'cvalores = cvalores & parreglo(I).Campo & "=" & parr_det(I, nfila) & ""
                            If Not IsNumeric(parr_det(I, nfila)) Then
                                cvalores = cvalores & parreglo(I).campo & "=Null"
                            Else
                                cvalores = cvalores & parreglo(I).campo & "=" & 0 & parr_det(I, nfila) & ""
                            End If
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = cvalores & parreglo(I).campo & "='" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & parreglo(I).campo & "=" & Format(parr_det(I, nfila), "yyyy-mm-dd")
                            Else
                                'cvalores = cvalores & parreglo(I).Campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                                If Not IsDate(parr_det(I, nfila)) Then
                                    cvalores = cvalores & parreglo(I).campo & "=Null"
                                Else
                                    cvalores = cvalores & parreglo(I).campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                                End If
                            End If
                        End If
                    Else
                        If parreglo(I).Tipo = "T" Then
                            cvalores = cvalores & "," & parreglo(I).campo & "='" & parr_det(I, nfila) & "'"
                        End If
                        If parreglo(I).Tipo = "N" Then
                            'cvalores = cvalores & "," & parreglo(I).Campo & "=" & parr_det(I, nfila) & ""
                            If Not IsNumeric(parr_det(I, nfila)) Then
                                cvalores = cvalores & "," & parreglo(I).campo & "=null"
                            Else
                                cvalores = cvalores & "," & parreglo(I).campo & "=" & parr_det(I, nfila) & ""
                            End If
                        End If
                        If parreglo(I).Tipo = "H" Then
                            cvalores = cvalores & "," & parreglo(I).campo & "='" & Format(parr_det(I, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(I).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & "," & parreglo(I).campo & "=" & Format(parr_det(I, nfila), "yyyy-mm-dd")
                            Else
'                                cvalores = cvalores & "," & parreglo(I).Campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                                If Not IsDate(parr_det(I, nfila)) Then
                                    'cvalores = cvalores & parreglo(I).Campo & "=Null"
                                    cvalores = cvalores & "," & parreglo(I).campo & "=null"
                                Else
                                    'cvalores = cvalores & parreglo(I).Campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                                    cvalores = cvalores & "," & parreglo(I).campo & "=CVDATE('" & parr_det(I, nfila) & "')"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        If ptipo = "A" Then
            csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
            csql = csql & cvalores & ")"
            If CnSave.State = 1 Then CnSave.Close
            CnSave.Open pconexion
            CnSave.Execute csql
            If pgraba_saldo = "*" Then '---- ACTUALIZA CANTIDAD Y PESO
                GRABA_SALDO parr_det(1, nfila), parr_det(4, nfila), parr_det(5, nfila), pmes, "I", CnSave
            End If
            If pgraba_saldo = "A" Then '---- ACTUALIZA CANTIDAD Y COSTO X ALMACEN
                GRABA_SALDO_ALM parr_det(1, nfila), parr_det(2, nfila), parr_det(5, nfila), pmes, wtipoguia, CnSave, parr_det(6, nfila), parr_det(10, nfila), "S"
            End If
        End If
    
        
        If ptipo = "M" Then
            csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
            CnSave.Execute csql
        End If
        ccampos = "": cvalores = ""
    Next
    Actualiza_Log csql, pconexion
    Sw_Graba_Registro = True
    Exit Sub
    
Error_Graba_Registro_Det:
    Sw_Graba_Registro = False
    MsgBox "Se ha producido el sgte. error " & Err.Description, vbCritical, "CONTROL Plus!"
    Resume Next
    Exit Sub
    
End Sub

Public Sub DELETEREC_N(ptabla As String, pconexion As String, pwhere As String)
On Error GoTo ErrorDelete

Dim Cn_Tmp As New ADODB.Connection
Dim CadSql As String
    Cn_Tmp.Open pconexion
    
    
    CadSql = "DELETE * FROM " & ptabla
    
    If Len(Trim(pwhere)) > 0 Then
        CadSql = CadSql & " where " & pwhere
    End If
    
    Cn_Tmp.Execute CadSql
    
    Actualiza_Log CadSql, pconexion
    
    
    If Cn_Tmp.State = 1 Then Cn_Tmp.Close
    Set Cn_Tmp = Nothing
    
    
    Exit Sub

ErrorDelete:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Exit Sub
End Sub

