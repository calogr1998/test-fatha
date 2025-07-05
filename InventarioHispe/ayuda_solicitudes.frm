VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_solicitudes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Solicitudes"
   ClientHeight    =   5760
   ClientLeft      =   2790
   ClientTop       =   2055
   ClientWidth     =   9825
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
   ScaleHeight     =   5760
   ScaleWidth      =   9825
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9550
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   8160
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
      Tools           =   "ayuda_solicitudes.frx":0000
      ToolBars        =   "ayuda_solicitudes.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4245
      Left            =   120
      OleObjectBlob   =   "ayuda_solicitudes.frx":663F
      TabIndex        =   3
      Top             =   945
      Width           =   9555
   End
End
Attribute VB_Name = "ayuda_solicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String

Private Sub dxDBGrid1_OnDblClick()
    num_solcomp = dxDBGrid1.Columns(0).value
    
    sw_limpia = True
    txtBusqueda.Text = ""
    sw_limpia = False
    Unload Me
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

    'csql = "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM"
    
    csql = "SELECT tb_cabsolicitud.cod_solicitud, tb_cabsolicitud.cs_fecha, tb_cabsolicitud.cs_codsolicitante, tb_cabsolicitud.cs_moneda, tb_cabsolicitud.cs_total, EF2PROVEEDORES.F2NEWRUC, EF2PROVEEDORES.F2NOMPROV " _
    & "FROM tb_cabsolicitud LEFT JOIN EF2PROVEEDORES ON tb_cabsolicitud.CS_PROVEEDOR = EF2PROVEEDORES.F2NEWRUC " _
    & "WHERE (((tb_cabsolicitud.cod_solicitud) In (SELECT DISTINCTROW tb_cabsolicitud.cod_solicitud FROM tb_cabsolicitud INNER JOIN tb_detsolicitud ON tb_cabsolicitud.cod_solicitud = tb_detsolicitud.cod_solicitud GROUP BY tb_cabsolicitud.cod_solicitud HAVING (((Sum(tb_detsolicitud.candis))>0))))) " _
    & "ORDER BY tb_cabsolicitud.cod_solicitud DESC;"
    
'    csql = "SELECT TB_CABSOLICITUD.cod_solicitud, TB_CABSOLICITUD.cs_fecha, TB_CABSOLICITUD.cs_codsolicitante, TB_CABSOLICITUD.cs_moneda, TB_CABSOLICITUD.cs_total, EF2PROVEEDORES.F2NEWRUC, EF2PROVEEDORES.F2NOMPROV " _
'          & "FROM EF2PROVEEDORES INNER JOIN TB_CABSOLICITUD ON TB_CABSOLICITUD.CS_PROVEEDOR = EF2PROVEEDORES.F2NEWRUC " _
'          & "WHERE TB_CABSOLICITUD.cod_solicitud In (SELECT DISTINCTROW TB_CABSOLICITUD.cod_solicitud FROM TB_CABSOLICITUD INNER JOIN tb_detsolicitud ON tb_cabsolicitud.cod_solicitud = tb_detsolicitud.cod_solicitud GROUP BY tb_cabsolicitud.cod_solicitud HAVING (((Sum(tb_detsolicitud.candis))>0))))) " _
'          & "ORDER BY TB_CABSOLICITUD.cod_solicitud DESC);"

    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "cod_solicitud"
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 13 Then
    dxDBGrid1_OnDblClick
End If
End Sub

Private Sub Form_Activate()
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    
    dxDBGrid1.Columns.FocusedIndex = 1
    dxDBGrid1.SetFocus
    
    dxDBGrid1.OptionEnabled = 1
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    
    Me.left = 500
    Me.top = 1050
    
    sw_limpia = False
        
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
    FILL
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
'    Select Case Tool.ID
'        Case "ID_Nuevo":
'            sw_nuevo_doc = True
'            sw_mant_ayuda = True
'            mant_marcas.Show 1
'            If sw_mant_ayuda = False Then Unload Me
'        Case "ID_Salir":
'            Unload Me
'    End Select

End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "cod_solicitud LIKE '*" & txtBusqueda.Text & "*' OR " & " cs_fecha LIKE '*" & txtBusqueda.Text & "*' "
    
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "cod_solicitud LIKE '*" & txtBusqueda.Text & "*' OR " & " cs_fecha LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub





