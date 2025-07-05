VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form valida_provi 
   Caption         =   "Validar Provisionales"
   ClientHeight    =   4125
   ClientLeft      =   105
   ClientTop       =   1200
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   11880
   Begin VB.TextBox txtbusqueda 
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   4020
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "valida_provi.frx":0000
      ToolBars        =   "valida_provi.frx":5899
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3240
      Left            =   120
      OleObjectBlob   =   "valida_provi.frx":59B4
      TabIndex        =   0
      Top             =   360
      Width           =   13125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Búsqueda"
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "valida_provi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TbDocumDet   As New ADODB.Recordset
Dim tbregisdoc      As New ADODB.Recordset
Public wcoddoc      As String, wtipo As String
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

ncorrela = busca_ctasxpagar()

If ncorrela = 0 Then
    MsgBox "No hay documentos por validar", vbInformation, "Sistema Logistica"
    Exit Sub
End If

dxDBGrid1.Dataset.Active = False
    
DELETEREC_LOG "DATOSPROV", cnn_form
DELETEREC_LOG "DATOSPROV", cnn_form

dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
dxDBGrid1.Dataset.Active = True

dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open
dxDBGrid1.Dataset.Refresh

dxDBGrid1.OptionEnabled = False
dxDBGrid1.Dataset.DisableControls


csql = "SELECT PAG_DCTO.CORRELA, PAG_DCTO.nro_comp, PAG_DCTO.fch_comp, PAG_DCTO.fch_vcto, PAG_DCTO.total, PAG_DCTO.F4MONTO1, PAG_DCTO.PROVEEDOR, PAG_DCTO.NOMPROV, First(REGISMOV.F3GASTO) AS CODGASTO, " & _
        "PAG_DCTO.moneda, IIf(PAG_DCTO.MONEDA = 'S', 'S/.', 'US$') AS SIMBOLO, PAG_DCTO.REFERENCIA, IIf(IsNull(PAG_DCTO.F4CENTRO) Or Len(PAG_DCTO.F4CENTRO)=0,REGISMOV.F3CENCOS,PAG_DCTO.F4CENTRO) AS CENTRO, PAG_DCTO.F4FECHA1, " & _
        "PAG_DCTO.SALDO, EF2PROVEEDORES.F2NOMPROV " & _
        "FROM ((PAG_DCTO LEFT JOIN EF2PROVEEDORES ON PAG_DCTO.proveedor = EF2PROVEEDORES.F2CODPROV) LEFT JOIN REGISDOC ON PAG_DCTO.correla = REGISDOC.F4CORRELA) LEFT JOIN REGISMOV ON (REGISDOC.F4MESMOV = REGISMOV.F4MESMOV) AND (REGISDOC.F4NUMMOV = REGISMOV.F4NUMMOV) " & _
        "GROUP BY PAG_DCTO.CORRELA, PAG_DCTO.nro_comp, PAG_DCTO.fch_comp, PAG_DCTO.fch_vcto, PAG_DCTO.total, PAG_DCTO.F4MONTO1, PAG_DCTO.PROVEEDOR, PAG_DCTO.NOMPROV, " & _
        "PAG_DCTO.moneda, IIf(PAG_DCTO.MONEDA = 'S', 'S/.', 'US$'), PAG_DCTO.REFERENCIA, IIf(IsNull(PAG_DCTO.F4CENTRO) Or Len(PAG_DCTO.F4CENTRO)=0,REGISMOV.F3CENCOS,PAG_DCTO.F4CENTRO), PAG_DCTO.F4FECHA1, " & _
        "PAG_DCTO.SALDO, EF2PROVEEDORES.F2NOMPROV,PAG_DCTO.deb_hab " & _
        "HAVING PAG_DCTO.deb_hab='H' and PAG_DCTO.correla=" & ncorrela & " and PAG_DCTO.correla not in (select corr_comp from provi_mvto where tipo_doc='P')"

'    dxDBGrid1.Columns.ColumnByFieldName("DESCOSTO").Visible = True
   
If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
TbDocumDet.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not TbDocumDet.EOF Then
    TbDocumDet.MoveFirst
    Do While Not TbDocumDet.EOF
        cgruporeal = traerCampo("BF9GIN", "GRUPCOMP", "CODIGO", "" & TbDocumDet.Fields("CODGASTO"), "and base='G'")
        
        csql = "SELECT PROVISIONALES.*, BF9GIN.GRUPCOMP AS GRUPO,BF9GIN.PRORRATEO FROM PROVISIONALES,BF9GIN WHERE (BF9GIN.CODIGO=PROVISIONALES.CONCEPTO AND BF9GIN.BASE='G') AND (ISNULL(BF9GIN.PRORRATEO) OR BF9GIN.PRORRATEO='') " & _
        "AND (ISNULL(PROVISIONALES.REFERENCIA) OR PROVISIONALES.REFERENCIA='') AND PROVISIONALES.SALDO>0 AND PROVISIONALES.DEB_HAB='H' AND PROVISIONALES.EST_ANUL = 'N' AND PROVISIONALES.PRESUPUESTO = 'N' " & _
        "AND PROVISIONALES.CCOSTO='" & "" & TbDocumDet.Fields("CENTRO") & "' AND (PROVISIONALES.CONCEPTO='" & "" & TbDocumDet.Fields("CODGASTO") & "' OR BF9GIN.GRUPCOMP='" & cgruporeal & "') ORDER BY CCOSTO"
    
        If rscta.State = adStateOpen Then rscta.Close
        rscta.Open csql, cnn_dbbancos, adOpenStatic, adLockOptimistic
        With dxDBGrid1.Dataset
        If Not rscta.EOF Then
            rscta.MoveFirst
            Do While Not rscta.EOF
                    .Append
                    .FieldValues("FECHAREAL") = "" & TbDocumDet.Fields("fch_comp")
                    .FieldValues("FECVCTOREAL") = "" & TbDocumDet.Fields("fch_vcto")
                    .FieldValues("NRODCTOREAL") = "" & TbDocumDet.Fields("nro_comp")
                    .FieldValues("MONEDAREAL") = "" & TbDocumDet.Fields("SIMBOLO")
                    .FieldValues("TOTALREAL") = Format(Val("" & TbDocumDet.Fields("total")), "0.00")
                    .FieldValues("SALDOREAL") = Format(Val("" & TbDocumDet.Fields("saldo")), "0.00")
                    .FieldValues("CORRELAREAL") = Val("" & TbDocumDet.Fields("correla"))
                    .FieldValues("CLIPROVREAL") = "" & TbDocumDet.Fields("proveedor")
                    .FieldValues("RAZONREAL") = "" & TbDocumDet.Fields("NOMPROV")
                    .FieldValues("CONCEPTOREAL") = "" & TbDocumDet.Fields("CODGASTO")
                    .FieldValues("DESCONCEPTO") = "" & traerCampo("BF9GIN", "NOMBRE", "CODIGO", "" & TbDocumDet.Fields("CODGASTO"), " AND BASE= 'G'")
                    .FieldValues("COSTOREAL") = "" & TbDocumDet.Fields("CENTRO")
                    .FieldValues("DESCOSTO") = "" & traerCampo("CENTROS", "F3ABREV", "F3COSTO", "" & TbDocumDet.Fields("CENTRO"))
                    .FieldValues("DETALLEREAL") = "" & TbDocumDet.Fields("REFERENCIA")
                    
                    .FieldValues("CONCEPTOPROV") = "" & rscta.Fields("CONCEPTO")
                    .FieldValues("DESGASTOPROV") = "" & traerCampo("BF9GIN", "NOMBRE", "CODIGO", "" & rscta.Fields("CONCEPTO"), " AND BASE= 'G'")
                    .FieldValues("COSTOPROV") = "" & rscta.Fields("CCOSTO")
                    .FieldValues("RAZONPROV") = "" & rscta.Fields("RAZON")
                    .FieldValues("DETALLEPROV") = "" & rscta.Fields("DETALLE")
                    .FieldValues("FECHAPROV") = "" & rscta.Fields("FECHAEMI")
                    .FieldValues("FECVCTOPROV") = rscta.Fields("FECHAVENC")
                    .FieldValues("NRODCTOPROV") = "" & rscta.Fields("NUMERO")
                    .FieldValues("MONEDAPROV") = "" & rscta.Fields("SIMBOLO")
                    .FieldValues("TOTALPROV") = Format(Val("" & rscta.Fields("MONTO")), "0.00") '--monto original
                    .FieldValues("SALDO") = Format(Val("" & rscta.Fields("SALDO")), "0.00")
                    .FieldValues("SALDOTEMP") = Format(Val("" & rscta.Fields("SALDO")), "0.00")
                    .FieldValues("CLIPROVPROVI") = "" & rscta.Fields("CODCLIPROV")
                    .FieldValues("CORRELAPROV") = Val("" & rscta.Fields("CORRELA"))
                    
                rscta.MoveNext
            Loop
            .Post
        End If
        rscta.Close
        End With
    TbDocumDet.MoveNext
    Loop
End If
TbDocumDet.Close
    
dxDBGrid1.Dataset.EnableControls
dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open
dxDBGrid1.OptionEnabled = True
dxDBGrid1.Dataset.Refresh
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
'        .Set (egoTabThrough)
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
'        .Set (egoNameCaseInsensitive)
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
    dxDBGrid1.Columns.ColumnByFieldName("CHECK").value = True
    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value, "###,###,##0.00")
    dxDBGrid1.Columns.ColumnByFieldName("SALDO").value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value - dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value, "###,###,##0.00")
    dxDBGrid1.Dataset.Post
    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ReadOnly = True
Else
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Columns.ColumnByFieldName("SALDO").value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value, "###,###,##0.00")
    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = "0.00"
    dxDBGrid1.Columns.ColumnByFieldName("CHECK").value = False
    dxDBGrid1.Dataset.Post
    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ReadOnly = False
'        dxDBGrid1.Columns.FocusedIndex = 6
    End If
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
If dxDBGrid1.Columns.ColumnByFieldName("CHECK").value = False Then
    dxDBGrid1.Dataset.Edit
    If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value, "0.00")) = 0# Then
        dxDBGrid1.Columns.ColumnByFieldName("SALDO").value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value, "###,###,##0.00")
    Else
        If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value, "0.00")) <= Val(Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value, "0.00")) Then
            dxDBGrid1.Columns.ColumnByFieldName("SALDO").value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value - dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value, "###,###,##0.00")
'            dxDBGrid1.Columns.ColumnByFieldName("SALDO").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value, "###,###,##0.00")
        Else
            MsgBox "El Monto excede al Saldo", vbInformation, "Atención"
            dxDBGrid1.Columns.ColumnByFieldName("SALDO").value = Format(dxDBGrid1.Columns.ColumnByFieldName("SALDOTEMP").value, "###,###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("TOTAL").value = "0.00"
        End If
    End If
    dxDBGrid1.Dataset.Post
End If
End Sub

Private Sub dxDBGrid1_OnEditing(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, Allow As Boolean)
    Select Case dxDBGrid1.Columns.FocusedIndex
        Case 16, 18:
            If dxDBGrid1.Columns.ColumnByFieldName("CHECK").value = False Then
                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ReadOnly = False
            Else
                dxDBGrid1.Columns.ColumnByFieldName("TOTAL").ReadOnly = True
            End If
    End Select
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    cnombase = "templus.mdb" '"TEMP_RESUMEN.MDB"
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & cnombase & ";Persist Security Info=False"
    
    Conf_Grid
    
    procesa2

    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
Dim H
 On Error Resume Next
  H = 350
  dxDBGrid1.Move 0, H, ScaleWidth, ScaleHeight - H
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim rscta   As New ADODB.Recordset
Dim Amov_Upd(0 To 20) As a_grabacion
Select Case Tool.Id
    Case "ID_Procesar":
        Csql2 = "SELECT * FROM DATOSPROV WHERE CHECK=TRUE or TOTAL<>0"
        If rscta.State = adStateOpen Then rscta.Close
        rscta.Open Csql2, cnn_form, adOpenStatic, adLockOptimistic
        If Not rscta.EOF Then
        
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                dxDBGrid1.Dataset.Post
            End If
            If MsgBox("¿Está Seguro de Validar los Doc. Provisionales Seleccionados?", vbYesNo + vbQuestion, "Atención") = vbYes Then
            rscta.MoveFirst
            Do While Not rscta.EOF
                ccodcliprov = Trim(rscta.Fields("CLIPROVREAL") & "")
                snumerodoc = rscta.Fields("NRODCTOREAL") & ""
                ncorrela = rscta.Fields("correlaprov") & ""
                ncorrelareal = rscta.Fields("CORRELAREAL") & ""
                ntotal = rscta.Fields("total") & ""
                nsaldo = rscta.Fields("saldo") & ""
                
                stipo = IIf(wcoddoc = "D", "C", "P")
                
                Csql2 = "INSERT INTO PROVI_MVTO (CODCLIPROV,CORR_COMP,CORR_PROVI,IMPORTE,fch_mvto,REFERENCIA,tipo_doc) VALUES " _
                            & "('" & ccodcliprov & "'," & ncorrelareal & "," & ncorrela & "," & ntotal & ",'" & CVDate(Date) & "','" & snumerodoc & "','" & stipo & "')"
                cnn_dbbancos.Execute (Csql2)
                'AlmacenaQuery_sql Csql2, cnn_dbbancos
              
                
                If rscta.Fields("CHECK") And nsaldo = 0 Then
                    Csql2 = "UPDATE PROVISIONALES SET REFERENCIA='" & snumerodoc & "', SALDO=" & nsaldo & ",EST_ANUL='S' " & _
                            "WHERE CORRELA=" & Val(ncorrela & "")
                    cnn_dbbancos.Execute (Csql2)
                    'AlmacenaQuery_sql Csql2, cnn_dbbancos
                    
                Else
                    Csql2 = "UPDATE PROVISIONALES SET SALDO=SALDO-" & ntotal & " WHERE CORRELA=" & Val(ncorrela & "")
                    cnn_dbbancos.Execute (Csql2)
                    'AlmacenaQuery_sql Csql2, cnn_dbbancos
'                        ---------Restaura los valores de monto1 y monto2
                    csql = "SELECT * FROM PROVISIONALES WHERE CORRELA=" & Val(ncorrela & "") & " AND (NOT ISNULL(MONTO1) and MONTO1>0)"
                    If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
                    TbDocumDet.Open csql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                    If Not TbDocumDet.EOF Then
                        Csql2 = "UPDATE PROVISIONALES SET MONTO1=SALDO,MONTO2=0,MONTO3=0,FECHA1=IIF(ISDATE(FECHA2),FECHA2,FECHA1),FECHA2=null,FECHA3=null WHERE CORRELA=" & Val(ncorrela & "")

                        cnn_dbbancos.Execute (Csql2)
                        'AlmacenaQuery_sql Csql2, cnn_dbbancos
                    End If
                End If
                    
                csql = "SELECT * FROM PROVISIONALES WHERE CORRELA=" & Val(ncorrela & "") & " AND SALDO <= 0"
                If TbDocumDet.State = adStateOpen Then TbDocumDet.Close
                TbDocumDet.Open csql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                If Not TbDocumDet.EOF Then
                    Csql2 = "UPDATE PROVISIONALES SET REFERENCIA='" & snumerodoc & "',EST_ANUL='S' WHERE CORRELA=" & Val(ncorrela & "")
                    cnn_dbbancos.Execute (Csql2)
                      'AlmacenaQuery_sql Csql2, cnn_dbbancos
                End If
                    
                rscta.MoveNext
            Loop
            MsgBox "Proceso Terminado", vbInformation, "Sistema Gerencial"

            procesa2
            
'            Unload Me
            End If
        Else
            MsgBox "No ha seleccionado ningun registro", vbInformation, "Atención"
        End If
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

