VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ConsProd_Proveedores 
   Caption         =   "Consulta de Productos - Proveedores"
   ClientHeight    =   7125
   ClientLeft      =   1005
   ClientTop       =   1350
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10620
   Begin VB.Frame Frame2 
      Height          =   7035
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   10410
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   6615
         Left            =   180
         OleObjectBlob   =   "ConsProd_Proveedores.frx":0000
         TabIndex        =   1
         Top             =   315
         Width           =   10155
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ConsProd_Proveedores.frx":10B5
      ToolBars        =   "ConsProd_Proveedores.frx":75E1
   End
End
Attribute VB_Name = "ConsProd_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nueva_Columna   As dxGridColumn
Dim contador        As Long
Dim csql            As String


Private Sub Form_Activate()

    With dxDBGrid1.Options
        '.Set (egoShowNewItemRow)
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Unset (egoPreview)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        '.Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowBands)
    End With

    FILL
    
End Sub

Private Sub Form_Load()

    Me.Height = 7550
    Me.Width = 10740
    Me.left = 1500
    Me.top = 980
    
    dxDBGrid1.Visible = True
    dxDBGrid1.Height = 6380

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    dxDBGrid1.Dataset.Close
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
         Case "ID_Salir"
            Unload Me
    End Select

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
Dim cIdtTip As String
Dim nValMes As Single
Dim nValAno As Single
    
    'ALMACENANDO LOS PRODUCTOS PARA SER COMPARADOS EN UNA CADENA -----------------'
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "SELECT * FROM tb_detsolicitud where cod_solicitud='" & Trim(solicitud.txtsolicitud.Text) & "' order by val(item)", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not Rs.EOF Then
        cad = ""
        Do While Not Rs.EOF
            If Trim(Rs.Fields("cod_producto")) <> "" Then
                cad = cad & Trim(Rs.Fields("cod_producto")) & ","
            End If
            Rs.MoveNext
            If Rs.EOF Then cad = left(cad, Len(Trim(cad)) - 1)
        Loop
    End If
    Rs.Close
    '-----------------------------------------------------------------------------'
    Rem EMB SQL = "Select F5CODPRO, F5MARCA, F2DESMAR, F5CODFAB, F5NOMPRO, F7CODMED, F5MONEDA, F5FOB, F5FACTOR FROM IF5PLA WHERE F5TIPO= 'P' ORDER BY F5CODFAB,F2DESMAR"
    sql = "SELECT EF2PROD_PROV.F5CODPRO, EF2PROD_PROV.F5CODFAB, EF2MARCAS.F2DESMAR, EF2PROD_PROV.F5VALVTA, EF2PROD_PROV.F2CODPRV " & _
        " FROM (EF2PROD_PROV INNER JOIN IF5PLA ON EF2PROD_PROV.F5CODPRO = IF5PLA.F5CODPRO) LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR;"
    'SQL = "select * from ConsProdProve"
    
    '----------- Para ver la consulta de referencias cruzadas -------------------'
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not Rs.EOF Then

        AddColumna
        AdicionaItems
   
        dxDBGrid1.Columns(0).Color = &HC0FFFF
        dxDBGrid1.Columns(1).Color = &HC0FFFF
        dxDBGrid1.Columns(2).Color = &HC0FFFF
        dxDBGrid1.Columns(3).Color = &HFFFFC0
        
    End If
    Rs.Close
    '-----------------------------------------------------------------------------'
    
End Sub

Private Sub AddColumna() 'ByVal Num As Long)

Dim NewColumn As dxGridColumn
Dim Num As Integer

    With dxDBGrid1
        .Dataset.Close
        SQL1 = "TRANSFORM First(ConsProdProve.F5VALVTA) AS [El Valor] " & _
            " SELECT ConsProdProve.F5CODPRO, ConsProdProve.F5CODFAB, ConsProdProve.F2DESMAR, Min(ConsProdProve.F5VALVTA) AS PreMínimo " & _
            " From ConsProdProve " & _
            " GROUP BY ConsProdProve.F5CODPRO, ConsProdProve.F5CODFAB, ConsProdProve.F2DESMAR " & _
            " ORDER BY ConsProdProve.F5CODPRO  PIVOT ConsProdProve.F2CODPRV;"
        If rsif5pla.State = adStateOpen Then rsif5pla.Close
        rsif5pla.Open SQL1, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif5pla.EOF Then
            'CODIGO
            Set NewColumn = .Columns.Add(gedTextEdit)
            NewColumn.ObjectName = "f5codpro"
            NewColumn.Caption = "CODIGO"
            NewColumn.FieldName = "F5CODPRO"
            'NewColumn.DisableCustomizing = True
            NewColumn.BandIndex = 1
            NewColumn.HeaderAlignment = taCenter
            NewColumn.Width = 20
            CadSql2 = CadSql2 & NewColumn.FieldName & "  TEXT(2),"
            'MODELO
            Set NewColumn = .Columns.Add(gedTextEdit)
            NewColumn.ObjectName = "f5codFAB"
            NewColumn.Caption = "MODELO"
            NewColumn.FieldName = "F5CODFAB"
            NewColumn.BandIndex = 1
            NewColumn.HeaderAlignment = taCenter
            NewColumn.Width = 30
            CadSql2 = CadSql2 & NewColumn.FieldName & "  TEXT(2),"
            'MARCA
            Set NewColumn = .Columns.Add(gedTextEdit)
            NewColumn.ObjectName = "f2desmar"
            NewColumn.Caption = "MARCA"
            NewColumn.FieldName = "F2DESMAR"
            NewColumn.BandIndex = 1
            NewColumn.HeaderAlignment = taCenter
            NewColumn.Width = 30
            CadSql2 = CadSql2 & NewColumn.FieldName & "  TEXT(2),"
            'PRECIO MINIMO
            Set NewColumn = .Columns.Add(gedTextEdit)
            NewColumn.ObjectName = "PreMínimo"
            NewColumn.Caption = "PreMínimo"
            NewColumn.FieldName = "PreMínimo"
            NewColumn.BandIndex = 1
            NewColumn.HeaderAlignment = taRightJustify
            NewColumn.Width = 25
            CadSql2 = CadSql2 & NewColumn.FieldName & "  TEXT(2),"
             
            '.Bands.Add
            '.Bands.ITEM(2).ObjectName = "Proveedor"
            '.Bands.ITEM(2).Caption = "Proveedor"
    
            For i = 0 To rsif5pla.Fields.Count - 6   'DO WHILE AL RECORDSET
                
                Set NewColumn = .Columns.Add(gedSpinEdit)
                NewColumn.ObjectName = rsif5pla.Fields(i + 5).Name    'Trim(Mid(lstproductos.List(i), 1, 7)) & "COCHES"
                NewColumn.Caption = "" & rsif5pla.Fields(i + 5).Name
                NewColumn.FieldName = Trim(Mid(rsif5pla.Fields(i + 5).Name, 1, 11)) 'CODSUBPROD
                'NewColumn.MinWidth = 20
                NewColumn.BandIndex = 2
                NewColumn.HeaderAlignment = taCenter
                NewColumn.Width = 50
                CadSql2 = CadSql2 & NewColumn.FieldName & "  DOUBLE,"
            Next
        End If
    End With

End Sub

Private Sub AdicionaItems()

Dim X As Integer
Dim VS As Integer

    '--------------------------'
    '---- GRID2 ------'
        dxDBGrid1.Dataset.ADODataset.CommandText = SQL1
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
        dxDBGrid1.Dataset.Active = True
        
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        dxDBGrid1.OptionEnabled = False
        dxDBGrid1.Dataset.DisableControls
    
        dxDBGrid1.Dataset.EnableControls
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        dxDBGrid1.OptionEnabled = True
        
    '--------------------------'
    
End Sub


