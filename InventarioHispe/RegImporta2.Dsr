VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RegImporta2 
   Caption         =   "Proyecto1 - RegImporta2 (ActiveReport)"
   ClientHeight    =   10620
   ClientLeft      =   -2985
   ClientTop       =   3210
   ClientWidth     =   15360
   Icon            =   "RegImporta2.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   18733
   SectionData     =   "RegImporta2.dsx":058A
End
Attribute VB_Name = "RegImporta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wopcion As Byte

Private Sub ActiveReport_Initialize()
With Me.Toolbar.Tools
    .ITEM(0).Visible = False
    .ITEM(2).Caption = "&Imprimir"
    .ITEM(2).Tooltip = "Imprimir"
    .ITEM(3).Visible = False
    .ITEM(4).Visible = False
    .Insert 6, "&Excel"
    .ITEM(6).AddIcon LoadPicture(App.Path & "\Excel.ico")
    .ITEM(6).Tooltip = "Genera archivo en Excel (*.xls)"
    .ITEM(6).Enabled = True
    .ITEM(7).Style = 1
    .ITEM(7).Caption = "&Buscar"
    .ITEM(7).Tooltip = "Buscar"
    .ITEM(9).Tooltip = "Una página"
    .ITEM(10).Tooltip = "Varias páginas"
    .ITEM(12).Tooltip = "Zoom (-)"
    .ITEM(13).Tooltip = "Zoom (+)"
    .ITEM(16).Tooltip = "Página anterior"
    .ITEM(17).Tooltip = "Página siguiente"
    .ITEM(20).Tooltip = "Atrás"
    .ITEM(20).Caption = "Atrás"
    .ITEM(21).Tooltip = "Adelante"
    .ITEM(21).Caption = "Adelante"
End With
End Sub

Private Sub ActiveReport_ReportStart()
'If wopcion = 1 Then
'    txtproducto.DataField = "f5nompro"
'Else
'    txtproducto.DataField = "f5nompro_ing"
'End If

m_sDBName = wrutabancos & "\DB_BANCOS.MDB"
'CARGA TABLA MERCADERIAS
sql = "SELECT IMPORT_DET1.F3MERCA AS F5MODELO, IMPORT_DET1.F3Cantidad, IMPORT_DET1.F3FobUni, "
sql = sql & "IMPORT_DET1.F3FobTot, IMPORT_DET1.F3CosUni, IMPORT_DET1.F3VtaUni, "
sql = sql & "IMPORT_DET1.F3Margen FROM IMPORT_DET1 INNER JOIN IF5PLA "
sql = sql & "ON IMPORT_DET1.F5CodPro = IF5PLA.F5CODPRO "
sql = sql & "WHERE (((IMPORT_DET1.F4NumImp)='" & New_Importaciones.txtnumero.Text & "'))"
'If rs.State = 1 Then rs.Close
'rs.Open sql, cnn_dbbancos, 3, 1
'RptPart1.Height = rs.RecordCount * 210
Set RptPart1.object = New Imp_Part1
RptPart1.object.dcRptData.DatabaseName = m_sDBName
RptPart1.object.dcRptData.RecordSource = sql
'CARGA TOTALES DE MERCADERIAS
sql = "SELECT Sum(IMPORT_DET1.F3FobUni) AS F3FobUni,Sum(IMPORT_DET1.F3FobTot) AS F3FobTot, "
sql = sql & "Sum(IMPORT_DET1.F3CosUni*IMPORT_DET1.F3Cantidad) AS F3CosUni, "
sql = sql & "Sum(IMPORT_DET1.F3VtaUni*IMPORT_DET1.F3Cantidad) AS F3VtaUni "
sql = sql & "FROM IMPORT_DET1 INNER JOIN IF5PLA ON IMPORT_DET1.F5CodPro = IF5PLA.F5CODPRO "
sql = sql & "GROUP BY IMPORT_DET1.F4NumImp "
sql = sql & "HAVING (((IMPORT_DET1.F4NumImp)='" & New_Importaciones.txtnumero.Text & "'))"

Set RptSuma1.object = New Imp_Suma1
RptSuma1.object.dcRptData.DatabaseName = m_sDBName
RptSuma1.object.dcRptData.RecordSource = sql

'CARGA DETALLE DE RESUMEN
sql = "SELECT BF9GIN.NOMBRE AS NOMGASTO, IMP_EMBxPLANT.F1CODGASTO, Sum(IMPORT_DET2.F4Total) AS "
sql = sql & "Total FROM BF9GIN INNER JOIN ((IMP_EMBxPLANT INNER JOIN IMP_PLANT_DET "
sql = sql & "ON (IMP_EMBxPLANT.F1SUBGRUPO = IMP_PLANT_DET.SUBGRUPO) AND "
sql = sql & "(IMP_EMBxPLANT.F1GRUPO = IMP_PLANT_DET.GRUPO)) INNER JOIN IMPORT_DET2 "
sql = sql & "ON (IMP_PLANT_DET.SUBGRUPO = IMPORT_DET2.F4SubGrupo) "
sql = sql & "AND (IMP_PLANT_DET.GRUPO = IMPORT_DET2.F4Grupo)) ON BF9GIN.CODIGO = "
sql = sql & "IMP_EMBxPLANT.F1CODGASTO "
sql = sql & "GROUP BY IMP_EMBxPLANT.F1EMBARCA, BF9GIN.NOMBRE, IMP_EMBxPLANT.F1CODGASTO, "
sql = sql & "IMPORT_DET2.F4NumImp Having (((IMP_EMBxPLANT.F1EMBARCA) = '" & right(New_Importaciones.CboImporta.Text, 4) & "') "
sql = sql & "And ((IMPORT_DET2.F4NumImp) = '" & New_Importaciones.txtnumero.Text & "')) ORDER BY Sum(IMPORT_DET2.F4Total) DESC"
Set RptResDet.object = New Imp_Res1
RptResDet.object.dcRptData.DatabaseName = m_sDBName
RptResDet.object.dcRptData.RecordSource = sql

'CARGA SUMA DE RESUMEN
sql = "SELECT Sum(IMPORT_DET2.F4Total) AS Total FROM BF9GIN INNER JOIN ((IMP_EMBxPLANT "
sql = sql & "INNER JOIN IMP_PLANT_DET ON (IMP_EMBxPLANT.F1GRUPO = IMP_PLANT_DET.GRUPO) "
sql = sql & "AND (IMP_EMBxPLANT.F1SUBGRUPO = IMP_PLANT_DET.SUBGRUPO)) INNER JOIN IMPORT_DET2 "
sql = sql & "ON (IMP_PLANT_DET.GRUPO = IMPORT_DET2.F4Grupo) AND (IMP_PLANT_DET.SUBGRUPO = "
sql = sql & "IMPORT_DET2.F4SubGrupo)) ON BF9GIN.CODIGO = IMP_EMBxPLANT.F1CODGASTO "
sql = sql & "GROUP BY IMP_EMBxPLANT.F1EMBARCA, IMPORT_DET2.F4NumImp "
sql = sql & "Having (((IMP_EMBxPLANT.F1EMBARCA) = '" & right(New_Importaciones.CboImporta.Text, 4) & "') "
sql = sql & "And ((IMPORT_DET2.F4NumImp) = '" & New_Importaciones.txtnumero.Text & "')) "
sql = sql & "ORDER BY Sum(IMPORT_DET2.F4Total) DESC"
Set RptResSum.object = New Imp_Res2
RptResSum.object.dcRptData.DatabaseName = m_sDBName
RptResSum.object.dcRptData.RecordSource = sql

'CARGA GASTOS
sql = "SELECT T3.T3Nro1, T3.T3Nro2, T3.T3Descripcion, T3.T3Detalle, "
sql = sql & "T3.T3Meses, T3.T3Interes, T3.T3Costo, T3.T3Igv, T3.T3Total, "
sql = sql & "T3.T3Inciden FROM TMP_IMP_DET3 AS T3 "
sql = sql & "ORDER BY T3.T3Nro1, T3.T3Nro2"
Set RptGastos.object = New Imp_Gastos
m_sDBName = wrutatemp & "\TMP_IMP.MDB"
RptGastos.object.dcRptData.DatabaseName = m_sDBName
RptGastos.object.dcRptData.RecordSource = sql


End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim oEXL As ActiveReportsExcelExport.ARExportExcel

    If Tool.Id = 4015 Then
        Load frmCommon
        strFilePath = frmCommon.Ruta
        Unload frmCommon
        If Trim(strFilePath) <> "" Then
            Set oEXL = New ActiveReportsExcelExport.ARExportExcel
            oEXL.FileName = strFilePath
            oEXL.Export Me.Pages
            MsgBox "Exportación terminada", vbInformation, "Reporte"
        End If
    End If

End Sub


Private Sub Detail_BeforePrint()
L1.Y2 = 0
L2.Y2 = 0
L3.Y2 = 0
L4.Y2 = 0
L5.Y2 = 0
L6.Y2 = 0
'********************
L1.Y2 = FldDet.Height
L2.Y2 = FldDet.Height
L3.Y2 = FldDet.Height
L4.Y2 = FldDet.Height
L5.Y2 = FldDet.Height
L6.Y2 = FldDet.Height
End Sub

