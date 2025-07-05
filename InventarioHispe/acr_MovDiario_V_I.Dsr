VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} acr_MovDiario_V_I 
   Caption         =   "Proyecto1 - acr_MovDiario_V_I (ActiveReport)"
   ClientHeight    =   5460
   ClientLeft      =   -2235
   ClientTop       =   1905
   ClientWidth     =   12000
   WindowState     =   2  'Maximized
   _ExtentX        =   21167
   _ExtentY        =   9631
   SectionData     =   "acr_MovDiario_V_I.dsx":0000
End
Attribute VB_Name = "acr_MovDiario_V_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumVale As String
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    Dim strSQL      As String
    Dim rpt         As Object
    'Set rpt = New acr_MovDiario_D_V_I
    
   strSQL = "select a.f5codpro,b.f5nompro,b.f7codmed,a.f3canpro,a.f3punit,a.f3valvta "
   strSQL = strSQL & "from if3vales a inner join if5pla b on a.f5codpro=b.f5codpro "
'''''   strSQL = strSQL & "where a.f4numval='" & Link & "' AND a.f2codalm='" & MovDiario_xAlmacen.Txtcodalm.Text & "'"
'strSQL = "Select * from MOV_DIARIO_VALES where f5Codpro = '" & Link & "'"
'Call ShowReport(rpt, Link, strSQL)
NumVale = Link
 IMPRIMIR_VALES '''(rpt, Link, strSQL)
End Sub
Private Sub IMPRIMIR_VALES()
Dim csql            As String
Dim CSQL1           As String
Dim Csql2           As String
Dim Csql3           As String
Dim prov            As String
Dim costo           As String
Dim RegEmpresa      As New ADODB.Recordset
Dim RegCosto        As New ADODB.Recordset
Dim RegMone        As New ADODB.Recordset
Dim ccod_almacen    As String
Dim cnum_vale       As String
Dim ctipo_vale      As String * 1
Dim WMONEDA As String * 1
    Me.MousePointer = vbHourglass
    
    CSQL1 = "Select* from IF4Vales where F4NumVal='" & NumVale & "'"
    RegMone.Open CSQL1, cnn_dbbancos, adOpenStatic, adLockReadOnly
    If Not (RegMone.EOF Or RegMone.Bof) Then
    WMONEDA = RegMone!F4MONEDA
'''''    ccod_almacen = MovDiario_xAlmacen.Txtcodalm.Text
    'cnum_vale = Trim(txtnumero.Text)
    'costo = Trim(txtccosto.Text)
    End If
     '   WMONEDA = IIf(cmbmoneda.ListIndex = 0, "S", "D")
        If WMONEDA = "D" Then
            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALDOL AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valdol] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
                " WHERE ((IF4VALES.F2CODALM)='" & RegMone!f2codalm & "') AND ((IF4VALES.F4NUMVAL)='" & NumVale & "') ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
        Else
            csql = "SELECT DISTINCTROW IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F1CODORI, SF1ORIGENES.F1NOMORI, IF4VALES.F4SERGUIA, IF4VALES.F4NUMGUIA, IF4VALES.F4TIPDOC, IF4VALES.F4SERDOC, IF4VALES.F4NUMDOC, EF2PROVEEDORES.F2CODPROV, IF4VALES.F2CODPROV AS RUC, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F4OBSERVA, IF4VALES.F4CENTRO, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO AS CODIGO, IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, IF5PLA.F7CODMED, IF5PLA.F5CODPRO, IF3VALES.F3VALVTA AS f5prevta, [if3vales].[F3CANPRO]*[if3vales].[f3valvta] AS PRECIO, EF2ALMACENES.F2NOMALM " & _
                " FROM EF2PROVEEDORES RIGHT JOIN (EF2ALMACENES INNER JOIN ((IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) ON EF2ALMACENES.F2CODALM = IF4VALES.F2CODALM) ON EF2PROVEEDORES.F2NEWRUC = IF4VALES.F2CODPROV " & _
                " WHERE ((IF4VALES.F2CODALM)='" & RegMone!f2codalm & "') AND ((IF4VALES.F4NUMVAL)='" & NumVale & "') ORDER BY IF4VALES.F4NUMVAL, IF3VALES.F5CODPRO;"
        End If
      With acr_MovDiario_D_V_I
        .DataControl1.ConnectionString = cnn_dbbancos
        'If Left(txtnumero.Text, 1) = "I" Then
            'prov = Trim(txtproveedor.Text)
            ctipo_vale = "I"
            .Lbl_vale.Caption = " VALE DE INGRESO "
            .lblprov.Visible = True
            '.lblpunto.Visible = True
            .fldprov.Visible = True
             .ToolbarVisible = False
'            .lblpie2.Caption = "Entregado por"
       'End If
        .DataControl1.Source = csql
'        .fldnomprov.Text = pnlproveedor.Caption
'        .fldalmacen.Text = Mid(cmbalmacen.Text, 200)
'        .fldnomcosto.Text = pnlccosto.Caption
          .fldempresa.Text = wnomcia
        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
        .fldvale.Text = NumVale
        .fldalma.Text = ccod_almacen
       If WMONEDA = "D" Then
            .LblMon.Caption = "US$."
       Else
            .LblMon.Caption = "S/"
       End If
        .Show vbModal
    End With

    
   ' End If
Me.MousePointer = vbDefault
End Sub

Private Sub ActiveReport_ReportStart()
Dim cnombase As String
'If sw_Report = False Then
cnombase = "TEMPLUS.MDB"
DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
'sw_Report = True
'Else
'DataControl1.ConnectionString = cconex_dbbancos
'sw_Report = False
'End If
End Sub

Private Sub Detail_Format()
TxtValeI.Hyperlink = TxtValeI.Text
End Sub

