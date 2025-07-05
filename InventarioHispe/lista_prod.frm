VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form lista_prod 
   Caption         =   "Lista de Productos"
   ClientHeight    =   7590
   ClientLeft      =   375
   ClientTop       =   1905
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   16200
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   855
      Left            =   10560
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6105
      Left            =   120
      OleObjectBlob   =   "lista_prod.frx":0000
      TabIndex        =   2
      Top             =   1080
      Width           =   16005
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   45
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   11
      Tools           =   "lista_prod.frx":2DAF
      ToolBars        =   "lista_prod.frx":D1C0
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10425
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10200
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Mnu 
         Caption         =   "Ver Producto"
         Index           =   0
      End
      Begin VB.Menu Mnu 
         Caption         =   "Ver Información Técnica"
         Index           =   1
      End
   End
End
Attribute VB_Name = "lista_prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstconsulta As New ADODB.Recordset
Dim rsfield As New ADODB.Recordset
Dim rstparametro As New ADODB.Recordset
Dim wcod(5)         As String
Dim wdes(5)         As String
Dim wlong1          As Integer
Dim wlong2          As Integer
Dim wniveles        As Integer
Dim wgraba          As String
Dim wtippro         As String


Private Sub dxDBGrid1_OnDblClick()
'    sw_nuevo_doc = False
'    codprod = dxDBGrid1.Columns(0).Value
'    mant_productos.Show 1
    
'    If ModUtilitario.validarFormAbierto("frmMantProducto") Then
'        Unload frmMantProducto
'    End If
'
'    With frmMantProducto
'        Me.MousePointer = vbHourglass
'
'        .Codigo = Trim(dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").value & "")
'
'        .Show 1
'
'        listarProductos
'
'        Me.MousePointer = vbDefault
'    End With
End Sub

Private Sub dxDBGrid1_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Single, ByVal Y As Single)
'If Shift = 0 And Button = vbRightButton Then
' PopupMenu Menu
'End If
End Sub

Private Sub Form_Activate()
 sw_load_mant = False
    
    'Adoprod.ConnectionString = cnn_dbbancos
    dxDBGrid1.Dataset.ADODataset.Requery
    'listarProductos
    
    
    'Adoprod.RecordSource = csql
    'Adoprod.Refresh
    
    wgraba = "0"
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
End Sub

Private Sub Form_Load()
Dim csql    As String
Dim i       As Integer

    Me.MousePointer = vbHourglass
    Me.left = 1600
    Me.top = 1050
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With
    listarProductos
    Me.MousePointer = vbDefault
 End Sub

Private Sub Form_Resize()
'    On Error Resume Next
'
'    FraBusqueda.Move 0, 0, Me.ScaleWidth, 870
'    txtbusqueda.Width = FraBusqueda.Width - 400
'    dxDBGrid1.Move 0, FraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height)
    
    On Error Resume Next
    
    dxDBGrid1.Move 0, fraBusqueda.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height) - 100
    
    fraBusqueda.left = 0
    fraBusqueda.top = 0
    fraProceso.left = fraBusqueda.Width + 100
    fraProceso.top = fraBusqueda.top
    fraProceso.Width = (dxDBGrid1.Width - fraBusqueda.Width) - 100
    fraProceso.Height = fraBusqueda.Height
    pgbProceso.Width = fraProceso.Width - 1000
    pgbProceso.left = (fraProceso.Width - pgbProceso.Width) / 2
End Sub

Private Sub Mnu_Click(Index As Integer)

Select Case Index
Case 0
    wNameImage = dxDBGrid1.Columns.ColumnByFieldName("F5RUTAIMAGE").value & ""
'''    ImagenProducto.Caption = "Imagen del Producto"
'''    ImagenProducto.Show 1
'''    Unload ImagenProducto
'''    Set ImagenProducto = Nothing
Case 1
    wNameImage = dxDBGrid1.Columns.ColumnByFieldName("F5RUTASERIE").value
'''    ImagenProducto.Caption = "Información Técnica"
'''    ImagenProducto.Show 1
'''    Unload ImagenProducto
'''    Set ImagenProducto = Nothing
End Select

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim csql As String

    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            'mant_productos.Show 1
        Case "ID_Imprimir":
            csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, EF2MARCAS.f2desmar, f5descontinuado " & _
               "FROM IF5PLA LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR " & _
               "WHERE IF5PLA.F5TIPO='P' " & _
               "ORDER BY IF5PLA.F5CODPRO;"
            
            With acr_productos
                .DataControl1.ConnectionString = cnn_dbbancos
                .DataControl1.Source = csql
                .fldFecha.Text = Format(Date, "DD/MM/YYYY")
                .lblempresa.Caption = wnomcia
                .Caption = "Relación de Productos"
                .Show 1
            End With
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
        Case "Importar"
            dxDBGrid1.Dataset.Close
            
            ModMilano.importarInsumoServidorExterno fraProceso, pgbProceso
            
            listarProductos
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub tdbprod_DblClick()

    sw_nuevo_mant = False
    'mant_productos.Show 1
    
End Sub
'
'Private Sub tdbprod_FilterChange()
'On Error GoTo errhandler
'Set cols = tdbprod.Columns
'Dim c As Integer
'
'    c = tdbprod.col
'    tdbprod.HoldFields
'    Adoprod.Recordset.Filter = getFilter()
'    tdbprod.col = c
'    tdbprod.EditActive = True
'    Exit Sub
'
'errhandler:
'
'    MsgBox Err.Source & ":" & vbCrLf & Err.Description
'    For Each col In tdbprod.Columns
'       col.FilterText = ""
'    Next col
'    Adoprod.Recordset.Filter = adFilterNone
'
'End Sub
'
'Private Function getFilter() As String
'Dim cadena As String
'Dim n As Integer
'
'    For Each col In cols
'        If Trim(col.FilterText) <> "" Then
'            n = n + 1
'            If n > 1 Then
'                cadena = cadena & " AND "
'            End If
'            cadena = cadena & col.DataField & " LIKE '" & col.FilterText & "*'"
'        End If
'    Next col
'   getFilter = cadena
'
'End Function

Private Sub listarProductos()
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "IF5PLA.F5CODPRO, "
    SqlCad = SqlCad & "IF5PLA.F5CODFAB, "
    SqlCad = SqlCad & "IF5PLA.F5NOMPRO, "
    SqlCad = SqlCad & "'' AS F2DESMAR, "
    SqlCad = SqlCad & "IF5PLA.F5DESCONTINUADO, "
    SqlCad = SqlCad & "EF7MEDIDAS.F7SIGMED as F7CODMED "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "IF5PLA "
    SqlCad = SqlCad & "INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "IF5PLA.F5CODPRO"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = SqlCad
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F5CODPRO"
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtBusqueda.Text & "*' " & _
    "OR " & " F5CODFAB LIKE '*" & txtBusqueda.Text & "*' " & _
    "or " & " F5NOMPRO like '*" & txtBusqueda.Text & "*' " & _
    "or " & " F7CODMED like '*" & txtBusqueda.Text & "*' " & _
    "or " & " f2desmar like '*" & txtBusqueda.Text & "*' " & _
    "or " & " f5descontinuado  like '*" & txtBusqueda.Text & "*' "

    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True

        dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtBusqueda.Text & "*' " & _
        "OR " & " F5CODFAB LIKE '*" & txtBusqueda.Text & "*' " & _
        "or " & " F5NOMPRO like '*" & txtBusqueda.Text & "*' " & _
        "or " & " F7CODMED like '*" & txtBusqueda.Text & "*' " & _
        "or " & " f2desmar like '*" & txtBusqueda.Text & "*' " & _
        "or " & " f5descontinuado  like '*" & txtBusqueda.Text & "*' "

    Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub


