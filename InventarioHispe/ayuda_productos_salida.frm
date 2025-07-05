VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_productos_salida 
   Caption         =   "Ayuda de Productos con Stock"
   ClientHeight    =   8535
   ClientLeft      =   585
   ClientTop       =   3960
   ClientWidth     =   14115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_productos_salida.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   14115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   10665
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar todos los productos"
         Height          =   240
         Left            =   7920
         TabIndex        =   4
         Top             =   520
         Width           =   2445
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Incluir Productos sin Stock"
         Height          =   240
         Left            =   7920
         TabIndex        =   3
         Top             =   240
         Width           =   2445
      End
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   5445
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "ayuda_productos_salida.frx":000C
      ToolBars        =   "ayuda_productos_salida.frx":975E
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6465
      Left            =   135
      OleObjectBlob   =   "ayuda_productos_salida.frx":9939
      TabIndex        =   2
      Top             =   1350
      Width           =   13665
   End
End
Attribute VB_Name = "ayuda_productos_salida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn_Mov     As New ADODB.Connection
Dim csql        As String
Dim Estado      As Boolean

Private Sub Check1_Click()
        If Check1.value = 1 Then
            Me.Caption = "Ayuda de todos los Productos con movimiento"
            Check2.value = 0
            FILL
        Else
            Me.Caption = "Ayuda de Productos con Stock"
            FILL
        End If
End Sub

Private Sub Check2_Click()
        If Check2.value = 1 Then
            Me.Caption = "Ayuda de todos los Productos"
            Check1.value = 0
            FILL
        Else
            Me.Caption = "Ayuda de Productos con Stock"
            FILL
        End If
End Sub

Private Sub Checkagrupar_Click()
    If Checkagrupar.value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").value
    wcodfab = dxDBGrid1.Columns.ColumnByFieldName("f5codfab").value
    wmarca = dxDBGrid1.Columns.ColumnByFieldName("f2desmar").value
    wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").value
    wcodmed = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7CODMED", "EF7MEDIDAS", "F7SIGMED", dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").value, "T")
    wmedida = dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").value
    wstockact = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f6stockact").value), 0, dxDBGrid1.Columns.ColumnByFieldName("f6stockact").value)
    wprecos = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").value), 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").value)
    wprecosdol = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").value), 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").value)
    wafecto = dxDBGrid1.Columns.ColumnByFieldName("f5afecto").value
    wtipocc = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").value), 0, dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").value)
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

End Sub

Private Sub FILL()

'If Con_Ayu = 1 Then
 If Check2.value = 1 Then
    If wcod_alm = "" Then wcod_alm = "01"
        If Len(Trim(wcodpartida)) > 0 Then
            csql = "SELECT DISTINCT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED,EF7MEDIDAS.F7SIGMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC,IF5PLA.f5vtanet,IF5PLA.f5vtanetdol "
            csql = csql & "FROM (((EF2MARCAS RIGHT JOIN ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES "
            csql = csql & "GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '01')]. AS Consulta3 "
            csql = csql & "RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA) LEFT JOIN [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES "
            csql = csql & "GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO) "
            csql = csql & "LEFT JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) INNER JOIN dbo_PartidaDetalle ON IF5PLA.F5CODPRO = dbo_PartidaDetalle.CodInsumo "
            csql = csql & " GROUP BY Consulta3.CANTIDAD, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, EF7MEDIDAS.F7SIGMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC, IF5PLA.F5VTANET, IF5PLA.F5VTANETDOL, IF5PLA.F5STOCKACT, dbo_PartidaDetalle.CodPresupuesto, dbo_PartidaDetalle.CodPartida"
            csql = csql & " HAVING (((dbo_PartidaDetalle.CodPresupuesto)='" & wcodpresupuesto & "') AND ((dbo_PartidaDetalle.CodPartida)='" & wcodpartida & "')) "
            csql = csql & " ORDER BY F5NOMPRO "
        Else
            csql = "SELECT DISTINCT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED,EF7MEDIDAS.F7SIGMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC,IF5PLA.f5vtanet,IF5PLA.f5vtanetdol"
            csql = csql & " FROM ((EF2MARCAS RIGHT JOIN ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '" & wcod_alm & "')]. AS Consulta3 "
            csql = csql & " RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA) "
            csql = csql & " LEFT JOIN [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO) LEFT JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED "
            csql = csql & " GROUP BY Consulta3.CANTIDAD, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, EF7MEDIDAS.F7SIGMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC, IF5PLA.F5VTANET, IF5PLA.F5VTANETDOL, IF5PLA.F5STOCKACT"
            csql = csql & " ORDER BY F5NOMPRO "
        End If
 Else
    If wcod_alm = "" Then wcod_alm = "01"
        csql = ""
        csql = "SELECT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.*, EF2MARCAS.F2DESMAR, Consulta2.F2CODALM, EF7MEDIDAS.F7SIGMED"
        csql = csql + " FROM ((EF2MARCAS RIGHT JOIN ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES GROUP BY IF3VALES.F5CODPRO, "
        csql = csql + " IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '01')]. AS Consulta3 RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA) INNER JOIN [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO "
        csql = csql + " FROM IF3VALES GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO) INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED "
        csql = csql + " WHERE (Consulta2.F2CODALM)='" & wcod_alm & "'"
        If Check1.value <> 1 Then
            csql = csql + " and NOT(ISNULL(Consulta3.CANTIDAD))"
        End If
        csql = csql & " ORDER BY F5NOMPRO;"
        
''        csql = "SELECT DISTINCT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, if5pla.F5STOCKACT as F6STOCKACT, IF5PLA.F5FECUC, IF5PLA.F5ULTTC,IF5PLA.f5vtanet"
''        csql = csql + " FROM EF2MARCAS INNER JOIN IF5PLA ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA "
''        csql = csql + " GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC,if5pla.F5STOCKACT, IF5PLA.F5ULTTC,IF5PLA.f5vtanet"
''        If Check1.Value <> 1 Then
''           csql = csql & " HAVING  (((Sum(IF5PLA.F5STOCKACT))>0))"
''        End If
''        csql = csql & " ORDER BY F5NOMPRO;"
'Else
'    If wcod_alm = "" Then wcod_alm = "01"
'    csql = ""
'    csql = csql & "SELECT A.F5CODPRO,A.F5CODFAB,D.F2DESMAR,A.F5NOMPRO,A.F7CODMED," & _
'           "A.F5VALVTA, A.F5FOB, A.F5AFECTO, C.F7SIGMED, B.F6STOCKACT, A.f5vtanet, A.f5fecuc, A.F5ULTTC " & _
'           "FROM ((IF5PLA AS A INNER JOIN IF6ALMA AS B " & _
'           "ON A.F5CODPRO = B.F5CODPRO) INNER JOIN EF7MEDIDAS AS C ON " & _
'           "A.F7CODMED = C.F7CODMED) INNER JOIN EF2MARCAS AS D ON " & _
'           "A.F5MARCA = D.F2CODMAR WHERE "
'    If Check1.Value <> 1 Then
'        csql = csql & "B.F6STOCKACT > 0 AND "
'    End If
'    csql = csql & "B.F2CODALM='" & wcod_alm & "' ORDER BY A.F5NOMPRO;"
'End If
 End If
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F5CODPRO"
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 13 Then
        dxDBGrid1_OnDblClick
    End If
End Sub

Private Sub Form_Load()
    Me.MousePointer = 11
    
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.left = 980
    Me.top = 1200
    If Cnn_Mov.State = adStateOpen Then Cnn_Mov.Close
    Cnn_Mov.ConnectionString = cnn_dbbancos
    Cnn_Mov.Open cconexion

    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = Cnn_Mov
    End With
    'If Con_Ayu = 1 Then
    '    Check1.Value = 1
    If Con_Ayu <> 2 Then
        Check2.value = 1
    Else
        Check2.value = 0
    End If
'        If Len(Trim(wcodpartida)) > 0 Then
'            Check1.Visible = False
'            'Check2.Visible = False jcg
'        End If
        'FILL
    'Else
    '    FILL
    'End If
    Me.MousePointer = 1
    FILL
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    FraBusqueda.Move 0, 0, Me.ScaleWidth, 870
    Check1.left = FraBusqueda.Width - (Check1.Width + 250)
    Check2.left = FraBusqueda.Width - (Check2.Width + 250)
    txtbusqueda.Width = FraBusqueda.Width - (Check1.Width + 750)
    dxDBGrid1.Move 0, FraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    Cnn_Mov.Close
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
            wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").value
            wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").value
            lista_compras.Show 1
            Unload Me
        Case "ID_Salir":
            Unload Me
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

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    On Error Resume Next
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True

        dxDBGrid1.Dataset.Filter = "F5NOMPRO LIKE '*" & txtbusqueda.Text & "*' or F5CODPRO like '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub



