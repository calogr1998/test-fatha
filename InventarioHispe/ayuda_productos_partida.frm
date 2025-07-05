VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_productos_partida 
   Caption         =   "Ayuda de Productos"
   ClientHeight    =   8535
   ClientLeft      =   585
   ClientTop       =   3960
   ClientWidth     =   16935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_productos_partida.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   16935
   StartUpPosition =   2  'CenterScreen
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "ayuda_productos_partida.frx":000C
      ToolBars        =   "ayuda_productos_partida.frx":975E
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6465
      Left            =   135
      OleObjectBlob   =   "ayuda_productos_partida.frx":98E6
      TabIndex        =   1
      Top             =   1350
      Width           =   16305
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   750
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   11505
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   10965
      End
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox1 
      Height          =   270
      Left            =   12480
      TabIndex        =   2
      Top             =   360
      Width           =   2760
      _Version        =   65536
      _cx             =   4868
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Mostrar productos seleccionados"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
End
Attribute VB_Name = "ayuda_productos_partida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim cnn_mov     As New ADODB.Connection
Dim csql        As String
Dim Estado      As Boolean


Private Sub Checkagrupar_Click()
    If Checkagrupar.Value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.Value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub dxCheckBox1_Click()
    If dxCheckBox1.Checked = 1 Then
        dxDBGrid1.Dataset.Filtered = True
        dxDBGrid1.Dataset.Filter = "F4PERINT = -1"
    Else
        dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
    Case "F5VALVTA", "F5VTANET", "AUTXINSUMO", "CANTIDAD_COMP"
        Text = Format(Text, "###,###,##0.00")
    End Select
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").Value
    wcodfab = dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value
    wmarca = dxDBGrid1.Columns.ColumnByFieldName("f2desmar").Value
    wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value
    wmedida = dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value
    wstockact = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f6stockact").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("f6stockact").Value)
    wprecos = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").Value)
    wprecosdol = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").Value)
    wafecto = dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value
    wtipocc = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").Value)
    Me.Hide
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With
    If dxDBGrid1.Columns.FocusedColumn.FieldName = "f5fob" Then
        dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("F4PERINT").Value = True
        dxDBGrid1.Dataset.Post
        
    End If
    

End Sub

Private Sub FILL()
Dim cat As New ADOX.Catalog
Dim Tbl As New Table
Dim tabla As String

cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\Templus.mdb;Persist Security Info=False"
tabla = "tmpProductos" 'Nombre de la tabla que quero verificar

For Each Tbl In cat.Tables
    If Tbl.Name = UCase(tabla) Or Tbl.Name = tabla Then
        cnn_form.Execute ("Drop Table tmpProductos")
    End If
Next Tbl

    If Len(Trim(wcodpartida)) > 0 Then
        csql = "SELECT DISTINCT IF5PLA.F4PERINT, IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO, EF7MEDIDAS.F7SIGMED, Format(dbo_PartidaDetalle.Cantidad*dbo_Partida.MATERIAL1,'0.00') AS F5VALVTA, IF5PLA.F5AFECTO, dbo_PartidaDetalle.Precio AS f5vtanet, Val(Format(Consulta2.SumaDeMETRADO*dbo_PartidaDetalle.Cantidad,'0.00')) AS AUTXINSUMO, consumido.cantidad_comp "
        csql = csql & " INTO tmpProductos IN '" & wrutatemp & "\Templus.mdb'"
        csql = csql & " FROM ([SELECT PRESUPUESTO, PARTIDA, Sum(METRADO) AS SumaDeMETRADO FROM CRONOGRAMA GROUP BY PRESUPUESTO, PARTIDA]. AS Consulta2 INNER JOIN dbo_Partida ON (Consulta2.PARTIDA = dbo_Partida.CodPartida) AND (Consulta2.PRESUPUESTO = dbo_Partida.CodPresupuesto)) INNER JOIN (((IF5PLA LEFT JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) INNER JOIN dbo_PartidaDetalle ON IF5PLA.F5CODPRO = dbo_PartidaDetalle.CodInsumo) LEFT JOIN [SELECT TB_DETSOLICITUD.F5CODCOSTO, TB_DETSOLICITUD.PARTIDA, TB_DETSOLICITUD.DES_PARTIDA, TB_DETSOLICITUD.cod_producto, Sum(TB_DETSOLICITUD.ds_cantidad) AS cantidad_comp "
        csql = csql & " FROM TB_CABSOLICITUD INNER JOIN TB_DETSOLICITUD ON (TB_CABSOLICITUD.cs_documento = TB_DETSOLICITUD.cs_documento) AND (TB_CABSOLICITUD.cod_solicitud = TB_DETSOLICITUD.cod_solicitud)"
        csql = csql & " GROUP BY TB_DETSOLICITUD.F5CODCOSTO, TB_DETSOLICITUD.PARTIDA, TB_DETSOLICITUD.DES_PARTIDA, TB_DETSOLICITUD.cod_producto]. as consumido ON (dbo_PartidaDetalle.CodInsumo = consumido.cod_producto) AND (dbo_PartidaDetalle.CodPartida = consumido.PARTIDA) AND (dbo_PartidaDetalle.CodPresupuesto = consumido.F5CODCOSTO)) ON (dbo_Partida.CodPresupuesto = dbo_PartidaDetalle.CodPresupuesto) AND (dbo_Partida.CodPartida = dbo_PartidaDetalle.CodPartida)"
        'csql = csql & " WHERE (((dbo_PartidaDetalle.CodPresupuesto)='" & wcodpresupuesto & "') AND ((dbo_PartidaDetalle.CodPartida)='" & wcodpartida & "'))"
        csql = csql & " WHERE dbo_PartidaDetalle.CodPresupuesto='" & wcodpresupuesto & "' AND dbo_PartidaDetalle.CodPartida='" & wcodpartida & "' and left(IF5PLA.F5CODPRO,2) <> 'MO'"
        csql = csql & " ORDER BY IF5PLA.F5NOMPRO"
        
'        csql = "SELECT DISTINCT IF5PLA.F4PERINT, IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, EF7MEDIDAS.F7SIGMED, Format(dbo_PartidaDetalle.Cantidad*dbo_Partida.MATERIAL1,'0.00') AS F5VALVTA, IF5PLA.F5AFECTO, dbo_PartidaDetalle.Precio AS f5vtanet, Val(Format(Consulta2.SumaDeMETRADO*dbo_PartidaDetalle.Cantidad,'0.00')) AS AUTXINSUMO"
'        csql = csql & " INTO tmpProductos IN '" & wrutatemp & "\Templus.mdb'"
'        csql = csql & " FROM ((SELECT PRESUPUESTO, PARTIDA, Sum(METRADO) AS SumaDeMETRADO FROM CRONOGRAMA GROUP BY PRESUPUESTO, PARTIDA) AS Consulta2 INNER JOIN dbo_Partida ON (Consulta2.PRESUPUESTO = dbo_Partida.CodPresupuesto) AND (Consulta2.PARTIDA = dbo_Partida.CodPartida)) INNER JOIN ((IF5PLA LEFT JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) "
'        csql = csql & " INNER JOIN dbo_PartidaDetalle ON IF5PLA.F5CODPRO = dbo_PartidaDetalle.CodInsumo) ON (dbo_Partida.CodPresupuesto = dbo_PartidaDetalle.CodPresupuesto) AND (dbo_Partida.CodPartida = dbo_PartidaDetalle.CodPartida)"
'        csql = csql & " WHERE (((dbo_PartidaDetalle.CodPresupuesto)='" & wcodpresupuesto & "') AND ((dbo_PartidaDetalle.CodPartida)='" & wcodpartida & "'))"
'        csql = csql & " ORDER BY IF5PLA.F5NOMPRO"
    Else
        csql = ""
        csql = "SELECT DISTINCT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.F4PERINT, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED,EF7MEDIDAS.F7SIGMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC,IF5PLA.f5vtanet,IF5PLA.f5vtanetdol"
        csql = csql & " INTO tmpProductos IN '" & wrutatemp & "\Templus.mdb'"
        csql = csql & " FROM ((EF2MARCAS RIGHT JOIN ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '" & wcod_alm & "')]. AS Consulta3 "
        csql = csql & " RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA) "
        csql = csql & " LEFT JOIN [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO) LEFT JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED "
        csql = csql & " GROUP BY Consulta3.CANTIDAD, IF5PLA.F4PERINT, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, EF7MEDIDAS.F7SIGMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC, IF5PLA.F5VTANET, IF5PLA.F5VTANETDOL, IF5PLA.F5STOCKACT"
        
        If oTipoRequerimiento = "ID_Nueva/O.S." Then
            csql = csql & " HAVING (IF5PLA.F5CODPRO Like 'SER%')"
        Else
            csql = csql & " HAVING (IF5PLA.F5CODPRO Not Like 'SER%')"
        End If
        
        csql = csql & " ORDER BY F5NOMPRO "
    End If
    
    dxDBGrid1.Dataset.Active = False
    If Len(Trim(wcodpartida)) > 0 Then
        cnn_dbbancos.Execute csql
        dxDBGrid1.Dataset.ADODataset.CommandText = "select * from tmpProductos"
'        dxDBGrid1.Columns.ColumnByFieldName("f6stockact").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("f5fecuc").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("f2desmar").Visible = False
'        dxDBGrid1.Columns.ColumnByFieldName("F5VALVTA").Visible = True '
    Else
        cnn_dbbancos.Execute csql
        dxDBGrid1.Dataset.ADODataset.CommandText = "select * from tmpProductos"
    End If
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F5CODPRO"
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
'    If KeyCode = 13 Then
'        dxDBGrid1_OnDblClick
'    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    dxDBGrid1.Dataset.Filter = ""
    txtbusqueda.SetFocus
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.left = 980
    Me.top = 1200
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\Templus.mdb;Persist Security Info=False"
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_form
    End With
    Me.MousePointer = vbDefault
    Me.Caption = "Ayuda de todos los Productos con movimiento"
    FILL
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    dxDBGrid1.Move 0, FraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    cnn_form.Close
    Set ayuda_productos = Nothing
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            mant_productos.Show 1
            If sw_mant_ayuda = False Then Unload Me
        Case "ID_Filtrar"
            
            If Tool.State = ssChecked Then
                dxDBGrid1.Filter.FilterActive = True
            Else
                dxDBGrid1.Filter.FilterActive = False
            End If
        Case "ID_Agrupar"
            If Tool.State = ssChecked Then
                dxDBGrid1.Options.Set (egoShowGroupPanel)
            Else
                dxDBGrid1.Options.Unset (egoShowGroupPanel)
            End If
        Case "ID_Compras":
            wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").Value
            wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value
            lista_compras.Show 1
            Unload Me
        Case "ID_Salir":
            Me.Hide
    End Select

End Sub

Private Sub txtbusqueda_Change()

    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F5NOMPRO LIKE '*" & txtbusqueda.Text & "*' or F5CODPRO like '*" & txtbusqueda.Text & "*' "
'    dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtbusqueda.Text & "*' " & _
'    "OR " & " F5CODFAB LIKE '*" & txtbusqueda.Text & "*' " & _
'    "or " & " F2DESMAR like '*" & txtbusqueda.Text & "*' " & _
'    "or " & " F5NOMPRO like '*" & txtbusqueda.Text & "*' " & _
'    "or " & " F7CODMED like '*" & txtbusqueda.Text & "*' "
''    "or " & " F5MARCA  like '*" & txtbusqueda.Text & "*' "

    If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtBusqueda_GotFocus()
    txtbusqueda.Text = Empty
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    On Error Resume Next
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True

        dxDBGrid1.Dataset.Filter = "F5NOMPRO LIKE '*" & txtbusqueda.Text & "*' or F5CODPRO like '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If
    
End Sub



