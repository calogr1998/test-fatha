VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form lista_compras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultimas compras"
   ClientHeight    =   6645
   ClientLeft      =   1185
   ClientTop       =   2355
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11430
   Begin VB.TextBox txtbusqueda 
      Height          =   315
      Left            =   1020
      TabIndex        =   3
      Top             =   480
      Width           =   4140
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1740
      TabIndex        =   2
      Top             =   90
      Width           =   2055
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   90
      Width           =   1455
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   5205
      Left            =   120
      OleObjectBlob   =   "lista_compras.frx":0000
      TabIndex        =   0
      Top             =   930
      Width           =   11265
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   180
      Top             =   -45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "lista_compras.frx":400B
      ToolBars        =   "lista_compras.frx":7F50
   End
   Begin VB.Label LblNomProv 
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Búsqueda"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   735
   End
End
Attribute VB_Name = "lista_compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    mostrar = True ' se carga en el sistemas invenat6
    sw_nuevo_doc = False
    Mant_Proveedores.Show 1

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

Private Sub Form_Activate()
Dim csql    As String
    dxDBGrid1.Option = egoAutoSearch
    CONECTAR
    csql = "SELECT IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL, IF4VALES.F2CODPROV, IF4VALES.F4REFERE as PROVEEDOR, IF3VALES.F3CANPRO,IF3VALES.F5CODPRO, IF4VALES.F4MONEDA, IIf([IF4VALES].[F4MONEDA]='S',[F3VALVTA],[F3VALDOL]) AS MONTO " & _
           "FROM IF4VALES INNER JOIN IF3VALES ON IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL AND IF4VALES.F2CODALM = IF3VALES.F2CODALM " & _
           "WHERE IF4VALES.F1CODORI='XC0' ORDER BY IF4VALES.F4FECVAL DESC;"
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODPROV"

    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    sw_mant_ayuda = False
    txtBusqueda.Text = wcodproducto
    LblNomProv.Caption = wdesproducto
    txtBusqueda.SetFocus
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    Me.left = 500
    Me.top = 1050
    Me.MousePointer = vbDefault
    
End Sub

Public Sub CONECTAR()
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
     Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            Mant_Proveedores.Show 1
        Case "ID_Imprimir":
           With acr_proveedores
                .DataControl1.ConnectionString = cnn_dbbancos
                .DataControl1.Source = "SELECT * FROM EF2PROVEEDORES ORDER BY F2NOMPROV"
                .fldFecha.Text = Format(Date, "DD/MM/YYYY")
                .lblEmpresa.Caption = wempresa
                .Show 1
            End With
        Case "ID_Salir"
            Unload Me
    End Select
End Sub


Private Sub txtbusqueda_Change()

    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtBusqueda.Text & "*' OR " & " F2CODPROV LIKE '*" & txtBusqueda.Text & "*' OR " & " PROVEEDOR like '*" & txtBusqueda.Text & "*' "
        
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
    
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F2CODPROV LIKE '*" & txtBusqueda.Text & "*' OR " & " F2CODPROV LIKE '*" & txtBusqueda.Text & "*' OR " & " PROVEEDOR like '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub

