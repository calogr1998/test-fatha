VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form valida_provi_reg 
   Caption         =   "Validar Provisionales"
   ClientHeight    =   2355
   ClientLeft      =   4110
   ClientTop       =   3615
   ClientWidth     =   9030
   Icon            =   "valida_provi_reg.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "valida_provi_reg.frx":000C
      ToolBars        =   "valida_provi_reg.frx":58A6
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   1935
      Left            =   0
      OleObjectBlob   =   "valida_provi_reg.frx":59C2
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "valida_provi_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TbDocumDet   As New ADODB.Recordset
Dim tbregisdoc      As New ADODB.Recordset
Public wcoddoc      As String, wtipo As String
Dim RsQ As New ADODB.Recordset
Function busca_ctasxpagar() As Integer
ncorre = 0
    sql = "Select * from regisdoc where f4nummov='" & Registro_Compras.TxtNumMov.Text & "' AND F4MESMOV='" & Registro_Compras.txtmesmov.Text & "'"
    If tbregisdoc.State = adStateOpen Then tbregisdoc.Close
    tbregisdoc.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not tbregisdoc.EOF Then
'        TRANS_CTASXPAGAR_NEW "P", "1", tbregisdoc.Fields("F4CORRELA"), UCase(Right(Trim(CmbTipDoc.Text), 3)), tbregisdoc.Fields("F4SERDOC"), tbregisdoc.Fields("F4NUMDOC"), tbregisdoc.Fields("F4FECHA"), tbregisdoc.Fields("F4RUCPRV"), tbregisdoc.Fields("F4CODPRV"), tbregisdoc.Fields("F4MONEDA"), tbregisdoc.Fields("F4TIPCAM"), tbregisdoc.Fields("F4TOTAL"), IIf(UCase(Right(Trim(CmbTipDoc.Text), 3)) = "CRE", "D", "H"), tbregisdoc.Fields("F4REFERE"), tbregisdoc.Fields("F4FECVEN"), tbregisdoc.Fields("F4OBRA") & "", tbregisdoc.Fields("F4NOMPRV"), cnn_dbbancos, tbregisdoc.Fields("F4MESMOV") & tbregisdoc.Fields("F4NUMMOV"), wanno, tbregisdoc.Fields("F4OCOMPRA")
        sql = "Select * from pag_dcto where correla=" & tbregisdoc.Fields("F4CORRELA")
        If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
        TbDocumDet.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not tbregisdoc.EOF Then
        ncorre = Val("" & TbDocumDet.Fields("CORRELA"))
        End If
        TbDocumDet.Close
    End If
    tbregisdoc.Close
    
    busca_ctasxpagar = ncorre
End Function


Public Sub procesa2()

Dim rscta   As New ADODB.Recordset
Dim fvenc1  As String, dfechavenc As String
Dim cprorrateo As String, cconcepto As String

Me.MousePointer = vbHourglass



'llena tabla
'---

Dim RsP As New ADODB.Recordset
Dim K As Integer
csql = "select * from temp_det order by f3item"
If Rs.State = 1 Then Rs.Close
Rs.Open csql, cnn_form, 3, 1
If Rs.RecordCount > 0 Then
    K = 1
    csql = "delete * from valida_compras"
    cnn_form.Execute csql
    'AlmacenaQuery_sql csql, cnn_form
    
    Rs.MoveFirst
    Do While Not Rs.EOF
    'carga provisionales
        csql = "SELECT CENTROS.F3ABREV, BF9GIN.NOMBRE, * FROM (PROVISIONALES LEFT JOIN BF9GIN "
        csql = csql & "ON PROVISIONALES.CONCEPTO = BF9GIN.CODIGO) LEFT JOIN CENTROS ON "
        csql = csql & "PROVISIONALES.CCOSTO = CENTROS.F3COSTO WHERE (((PROVISIONALES.TIPO) "
        csql = csql & "Like 'P999') AND ((PROVISIONALES.CONCEPTO)='" & Rs!F3GASTO & "') AND "
        csql = csql & "((PROVISIONALES.CCOSTO)='" & Rs!F3CENCOS & "'))"
        If RsP.State = 1 Then RsP.Close
        RsP.Open csql, cnn_dbbancos, 3, 1
        If RsP.RecordCount > 0 Then
            
            csql = "select * from valida_compras where ccosto='" & Rs!F3CENCOS & "' and cgasto='" & Rs!F3GASTO & "'"
            If RsQ.State = 1 Then RsQ.Close
            RsQ.Open csql, cnn_form, 3, 1
            If RsQ.RecordCount = 0 Then
                csql = "insert into valida_compras (item,correla,ccosto,proyecto,cgasto,desgasto,provisionado) values "
                csql = csql & "(" & K & "," & RsP!Correla & ",'" & RsP!ccosto & "','" & RsP!f3abrev & "',"
                csql = csql & "'" & RsP!concepto & "','" & RsP!nombre & "'," & RsP!MONTO & ")"
            Else
                'csql = "update valida_compras set provisionado=provisionado+" & RsP!MONTO & " "
                'csql = csql & "where ccosto='" & rs!F3CENCOS & "' and cgasto='" & rs!F3GASTO & "'"
            End If
            cnn_form.Execute csql
            'AlmacenaQuery_sql csql, cnn_form
        End If
        K = K + 1
    'carga compras
        csql = "SELECT Sum(REGISMOV.F3IMPORTE) AS SumaDeF3IMPORTE From REGISMOV "
        csql = csql & "GROUP BY REGISMOV.F3GASTO, REGISMOV.F3CENCOS HAVING "
        csql = csql & "(((REGISMOV.F3GASTO)='" & Rs!F3GASTO & "') AND "
        csql = csql & "((REGISMOV.F3CENCOS)='" & Rs!F3CENCOS & "'));"
        
        csql = "SELECT Sum(IIf(Trim(REGISMOV.F3AFECTO)='*',REGISMOV.F3IMPORTe*" & 1 + wIgv & ",REGISMOV.F3IMPORTe)) AS SumaDeF3IMPORTE "
        csql = csql & "From REGISMOV "
        csql = csql & "GROUP BY REGISMOV.F3GASTO, REGISMOV.F3CENCOS "
        'csql = csql & "HAVING f4nummov='" & registro_compras.TxtNumMov.Text & "' AND F4MESMOV='" & registro_compras.txtmesmov.Text & "'"
        csql = csql & "HAVING (((REGISMOV.F3GASTO)='" & Rs!F3GASTO & "') AND "
        csql = csql & "((REGISMOV.F3CENCOS)='" & Rs!F3CENCOS & "'));"
        
        If RsP.State = 1 Then RsP.Close
        RsP.Open csql, cnn_dbbancos, 3, 1
        If RsP.RecordCount > 0 Then
            
            csql = "update valida_compras set comprado=" & RsP!SumaDeF3IMPORTE & " where cgasto='" & Rs!F3GASTO & "' and ccosto ='" & Rs!F3CENCOS & "'"
            cnn_form.Execute csql
            'AlmacenaQuery_sql csql, cnn_form
            
        End If
    '-------------
        Rs.MoveNext
    Loop
    csql = "update valida_compras set saldo=val(str(provisionado)+'')-val(str(comprado)+'')"
    cnn_form.Execute csql
      'AlmacenaQuery_sql csql, cnn_form
      
    csql = "update valida_compras set valida=iif(saldo<=0,-1,0)"
    cnn_form.Execute csql
      'AlmacenaQuery_sql csql, cnn_form
      
End If


'---
csql = "select * from valida_compras"
    
    

dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
dxDBGrid1.Dataset.Active = False
dxDBGrid1.Dataset.ADODataset.CommandText = csql
dxDBGrid1.Dataset.Active = True
dxDBGrid1.KeyField = "item"


 

Me.MousePointer = 0
End Sub

Private Sub Conf_Grid()

    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoAutoSearch)
'        .Set (egoCanDelete)
'        .Set (egoCanAppend)
'        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoExactScrollBar)
'        .Set (egoHorzThrough)
        .Set (egoVertThrough)
'        '.Set (egoAutoWidth)
        .Set (egoShowBorder)
        .Set (egoEnterShowEditor)
'        .Set (egoEnterThrough)
        .Set (egoShowVertScrollTip)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
'        .Set (egoCanNavigation)
'        .Set (egoCancelOnExit)
'        .Set (egoLoadAllRecords)
'        .Set (egoShowHourGlass)
'        .Set (egoUseBookmarks)
'        .Set (egoUseLocate)
'        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
'        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
'        .Set (egoShowCellTip)
        .Set (egoShowBands)
        .Set (egoHeaderButtonClicking)
    End With

End Sub



Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
'If dxDBGrid1.Dataset.State = 2 Then dxDBGrid1.Dataset.Post
If State = cbsChecked Then
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Columns.ColumnByFieldName("VALIDA").value = True
    dxDBGrid1.Dataset.Post
Else
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Columns.ColumnByFieldName("VALIDA").value = False
    dxDBGrid1.Dataset.Post
End If
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
'If dxDBGrid1.Columns.ColumnByFieldName("CHECK").Value = False Then
'    dxDBGrid1.Dataset.Edit
'    If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value, "0.00")) = 0# Then
'        dxDBGrid1.Columns.ColumnByFieldName("SALDO").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").Value, "###,###,##0.00")
'    Else
'        If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value, "0.00")) <= Val(Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").Value, "0.00")) Then
'            dxDBGrid1.Columns.ColumnByFieldName("SALDO").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").Value - dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value, "###,###,##0.00")
'            dxDBGrid1.Columns.ColumnByFieldName("SALDO").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value, "###,###,##0.00")
'        Else
'            MsgBox "El Monto excede al Saldo", vbInformation, "Atención"
'            dxDBGrid1.Columns.ColumnByFieldName("SALDO").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").Value, "###,###,##0.00")
'            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = "0.00"
'        End If
'   End If
'    dxDBGrid1.Dataset.Post
'End If
End Sub

Private Sub dxDBGrid1_OnEditing(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, Allow As Boolean)
 '   Select Case dxDBGrid1.Columns.FocusedIndex
 '       Case 16, 18:
 '           If dxDBGrid1.Columns.ColumnByFieldName("CHECK").Value = False Then
 '               dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ReadOnly = False
 '           Else
 '               dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ReadOnly = True
 '           End If
 '   End Select
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    'cnombase = "TEMP_RESUMEN.MDB"
    'If cnn_form.State = adStateOpen Then cnn_form.Close
    'cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ";Persist Security Info=False"
    
    cnombase = "TEMPLUS.MDB" '"TEMP_COM.MDB"
    cconex_form = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & cnombase & "; Persist Security Info=False"
    
    If cnn_form.State = 1 Then cnn_form.Close
    
    cnn_form.Open cconex_form
    
    Conf_Grid
    
    procesa2

    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
Dim H
 On Error Resume Next
  H = 0
  dxDBGrid1.Move 0, H, ScaleWidth, ScaleHeight - H
  dxDBGrid1.Bands.ITEM(0).Width = Me.Width / 16.1
End Sub

Private Sub OrdenaMontos(ncorrela As Integer)
Dim i As Integer, J As Integer
Dim sDatos(1 To 8)  As a_MontosFechas
Dim Mov_Upd(0 To 15) As a_grabacion
Dim RsM As New ADODB.Recordset
'*************limpia arreglo
For i = 1 To 8
    sDatos(i).MONTO = "": sDatos(i).Fecha = ""
Next
'*************
csql = "select * from provisionales where correla=" & ncorrela
If RsM.State = 1 Then RsM.Close
RsM.Open csql, cnn_dbbancos, 3, 1
If RsM.RecordCount > 0 Then
    J = 0
    RsM.MoveFirst
    Do While Not RsM.EOF
        For i = 1 To 8
            If Val(RsM.Fields("monto" & i) & "") <> 0 Then
                J = J + 1
                sDatos(J).MONTO = Val(RsM.Fields("monto" & i) & ""): sDatos(J).Fecha = RsM.Fields("fecha" & i) & ""
            End If
        Next
        'arma arreglo para actualizar
        For i = 1 To 8
            Mov_Upd(i - 1).campo = "MONTO" & i: Mov_Upd(i - 1).valor = sDatos(i).MONTO: Mov_Upd(i - 1).TIPO = "N"
            Mov_Upd(i + 7).campo = "FECHA" & i: Mov_Upd(i + 7).valor = sDatos(i).Fecha: Mov_Upd(i + 7).TIPO = "F"
        Next
        GRABA_REGISTRO_logistica Mov_Upd, "PROVISIONALES", "M", 15, cnn_dbbancos, "CORRELA=" & RsM!Correla
        
        RsM.MoveNext
    Loop
End If
End Sub

Private Sub RestaMonto(ncorrela As Integer, nMontoComprado As Double)
Dim i As Integer, J As Integer
Dim nAplicado As Double
Dim sDatos(1 To 8)  As a_MontosFechas
Dim Mov_Upd(0 To 15) As a_grabacion
Dim RsM As New ADODB.Recordset
'*************limpia arreglo
For i = 1 To 8
    sDatos(i).MONTO = "": sDatos(i).Fecha = ""
Next
'*************
csql = "select * from provisionales where correla=" & ncorrela
If RsM.State = 1 Then RsM.Close
RsM.Open csql, cnn_dbbancos, 3, 1
If RsM.RecordCount > 0 Then
    J = 0
    RsM.MoveFirst
    Do While Not RsM.EOF
        For i = 1 To 8
            If Val(RsM.Fields("monto" & i) & "") <> 0 Then
                J = J + 1
                sDatos(J).MONTO = Val(RsM.Fields("monto" & i) & ""): sDatos(J).Fecha = RsM.Fields("fecha" & i) & ""
            End If
        Next
        '****************
        J = 1
        Do While nMontoComprado > 0
            If J = 8 Then
                nMontoComprado = 0
            End If
            If Val(sDatos(J).MONTO) < nMontoComprado Then
                nMontoComprado = nMontoComprado - Val(sDatos(J).MONTO)
                sDatos(J).MONTO = 0
            ElseIf Val(sDatos(J).MONTO) = nMontoComprado Then
                nMontoComprado = 0
                sDatos(J).MONTO = 0
            Else
                sDatos(J).MONTO = Val(sDatos(J).MONTO) - nMontoComprado
                nMontoComprado = 0
            End If
            J = J + 1
        Loop
        'arma arreglo
        For i = 1 To 8
            Mov_Upd(i - 1).campo = "MONTO" & i: Mov_Upd(i - 1).valor = sDatos(i).MONTO: Mov_Upd(i - 1).TIPO = "N"
            Mov_Upd(i + 7).campo = "FECHA" & i: Mov_Upd(i + 7).valor = sDatos(i).Fecha: Mov_Upd(i + 7).TIPO = "F"
        Next
        GRABA_REGISTRO_logistica Mov_Upd, "PROVISIONALES", "M", 15, cnn_dbbancos, "CORRELA=" & RsM!Correla
        
        OrdenaMontos ncorrela
        
        RsM.MoveNext
    Loop
End If
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim rscta   As New ADODB.Recordset
Dim ncorrela As Integer
Select Case Tool.Id
    Case "ID_Procesar":
        ncorrela = traerCampo("REGISDOC", "F4CORRELA", "F4MESMOV", Registro_Compras.txtmesmov.Text, " and F4NUMMOV = '" & Registro_Compras.TxtNumMov.Text & "'")
        csql = "SELECT * FROM valida_compras order by item"
        If Rs.State = 1 Then Rs.Close
        Rs.Open csql, cnn_form, 3, 1
        If Rs.RecordCount > 0 Then
            Rs.MoveFirst
            Do While Not Rs.EOF
                wcodcli = traerCampo("centros", "f3codcli", "f3costo", Rs!ccosto & "", "")
                csql = "select * from provi_mvto where corr_comp=" & ncorrela & " and corr_provi=" & Rs!Correla
                If rscta.State = 1 Then rscta.Close
                rscta.Open csql, cnn_dbbancos, 3, 1
                If rscta.RecordCount > 0 Then
                    'actualiza provi_mvto
                    csql = "update provi_mvto set codcliprov='" & wcodcli & "', importe=" & Rs!comprado
                    csql = csql & ",fch_mvto='" & Registro_Compras.TxtFecVen.value
                    csql = csql & "',referencia='" & UCase(left(Registro_Compras.CmbTipDoc.Text, 3)) & Registro_Compras.TxtSerDoc.Text & "/" & Registro_Compras.TxtNumDoc.Text & "' "
                    csql = csql & "where corr_comp=" & ncorrela & " and corr_provi="
                    csql = csql & Rs!Correla
                Else
                    'inserta provi_mvto
                    csql = "Insert into provi_mvto(codcliprov,corr_comp,corr_provi,importe,"
                    csql = csql & "tcambio,moneda,fch_mvto,referencia,tipo_doc) values "
                    csql = csql & "('" & wcodcli & "'," & ncorrela & "," & Rs!Correla
                    csql = csql & "," & Rs!comprado & "," & Registro_Compras.TxtTipCam.Text
                    csql = csql & ",'" & IIf(Registro_Compras.Mon.Caption = "D", "D", "S")
                    csql = csql & "','" & Registro_Compras.TxtFecVen.value
                    csql = csql & "','" & UCase(left(Registro_Compras.CmbTipDoc.Text, 3)) & Registro_Compras.TxtSerDoc.Text & "/" & Registro_Compras.TxtNumDoc.Text & "',"
                    csql = csql & "'P')"
                End If
                cnn_dbbancos.Execute csql
                  'AlmacenaQuery_sql csql, cnn_dbbancos
                'ORGANIZA ORDEN DE LOS MONTOS
                OrdenaMontos Rs!Correla & ""
                'actualiza provisional
                If Rs!valida = True Then
                    cest_anul = "S"
                    nsaldo = 0
                Else
                    cest_anul = "N"
                    nsaldo = Rs!Saldo
                    RestaMonto Rs!Correla, Val(Rs!comprado & "")
                End If
                csql = "update provisionales set est_anul='" & cest_anul & "',saldo=" & nsaldo & ", "
                csql = csql & " fecmod='" & Date & "', usemod='" & wusuario & "' where correla=" & Rs!Correla & " and tipo='P999'"
                
                cnn_dbbancos.Execute csql
                'AlmacenaQuery_sql csql, cnn_dbbancos
                
                Rs.MoveNext
            Loop
        End If
        Unload Me
    Case "ID_Salir":
        Unload Me
End Select
End Sub


Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "DESCOSTO LIKE '*" & txtBusqueda.Text & "*' OR " & " NRODCTOREAL LIKE '*" & txtBusqueda.Text & "*' OR " & " RAZONREAL LIKE '*" & txtBusqueda.Text & "*' OR " & " NRODCTOPROV LIKE '*" & txtBusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    dxDBGrid1.SetFocus
    End If
End Sub

