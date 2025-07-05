VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_productos_xalmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Productos con Stock"
   ClientHeight    =   5040
   ClientLeft      =   2865
   ClientTop       =   3000
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   10935
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   315
      TabIndex        =   6
      Top             =   1035
      Width           =   1455
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1755
      TabIndex        =   5
      Top             =   1035
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   10665
      Begin VB.CheckBox Check2 
         Caption         =   "Mostrar todos los productos"
         Height          =   240
         Left            =   7740
         TabIndex        =   7
         Top             =   520
         Width           =   2625
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Productos sin Stock"
         Height          =   240
         Left            =   7740
         TabIndex        =   4
         Top             =   240
         Width           =   2265
      End
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   5445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   315
         TabIndex        =   1
         Top             =   405
         Width           =   735
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
      Tools           =   "ayuda_productos_xalmacen.frx":0000
      ToolBars        =   "ayuda_productos_xalmacen.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3120
      Left            =   135
      OleObjectBlob   =   "ayuda_productos_xalmacen.frx":663F
      TabIndex        =   3
      Top             =   1350
      Width           =   10665
   End
End
Attribute VB_Name = "ayuda_productos_xalmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
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
    wmedida = dxDBGrid1.Columns.ColumnByFieldName("f7codmed").value
    wstockact = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f6stockact").value), 0, dxDBGrid1.Columns.ColumnByFieldName("f6stockact").value)
    wprecos = IIf(dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").value = Null, 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").value)
    wprecosdol = IIf(dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").value = Null, 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").value)
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
        csql = "SELECT DISTINCT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, iif(isnull(if5pla.F5STOCKACT),0,0) as F6STOCKACT, IF5PLA.F5FECUC, IF5PLA.F5ULTTC,IF5PLA.f5vtanet,IF5PLA.f5vtanetdol"
        csql = csql + " FROM EF2MARCAS INNER JOIN IF5PLA ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA "
        csql = csql + " GROUP BY IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC,if5pla.F5STOCKACT, IF5PLA.F5ULTTC,IF5PLA.f5vtanet,,IF5PLA.f5vtanetdol"
        csql = csql & " ORDER BY F5NOMPRO;"
 Else
    If wcod_alm = "" Then wcod_alm = "01"
        csql = ""
        csql = "SELECT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.*, EF2MARCAS.F2DESMAR, Consulta2.F2CODALM"
        csql = csql + " FROM (EF2MARCAS INNER JOIN ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '" & wcod_alm & "')]. AS Consulta3"
        csql = csql + " RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA)"
        csql = csql + " INNER JOIN  [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO"
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
    Me.MousePointer = vbHourglass
    
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.left = 500
    Me.top = 1950
    If cnn_mov.State = adStateOpen Then cnn_mov.Close
    cnn_mov.ConnectionString = cnn_dbbancos
    cnn_mov.Open cconexion
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_mov
    End With
    'If Con_Ayu = 1 Then
    '    Check1.Value = 1
        FILL
    'Else
    '    FILL
    'End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    cnn_mov.Close
    Set ayuda_productos = Nothing
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            mant_productos.Show 1
            If sw_mant_ayuda = False Then Unload Me
        Case "ID_Salir":
            Unload Me
    End Select

End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtBusqueda.Text & "*' " & _
    "OR " & " F5CODFAB LIKE '*" & txtBusqueda.Text & "*' " & _
    "or " & " F2DESMAR like '*" & txtBusqueda.Text & "*' " & _
    "or " & " F5NOMPRO like '*" & txtBusqueda.Text & "*' " & _
    "or " & " F7CODMED like '*" & txtBusqueda.Text & "*' "
'    "or " & " F5MARCA  like '*" & txtbusqueda.Text & "*' "

    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True

        dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtBusqueda.Text & "*' " & _
        "OR " & " F5CODFAB LIKE '*" & txtBusqueda.Text & "*' " & _
        "or " & " F2DESMAR like '*" & txtBusqueda.Text & "*' " & _
        "or " & " F5NOMPRO like '*" & txtBusqueda.Text & "*' " & _
        "or " & " F7CODMED like '*" & txtBusqueda.Text & "*' "
'        "or " & " F5MARCA  like '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    'End If
    
End Sub



