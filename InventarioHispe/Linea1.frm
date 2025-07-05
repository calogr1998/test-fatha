VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form linea1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Linea de Productos"
   ClientHeight    =   5940
   ClientLeft      =   6405
   ClientTop       =   3360
   ClientWidth     =   5010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4770
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
         Height          =   210
         Left            =   120
         TabIndex        =   2
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
      Tools           =   "Linea1.frx":0000
      ToolBars        =   "Linea1.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4380
      Left            =   120
      OleObjectBlob   =   "Linea1.frx":668D
      TabIndex        =   3
      Top             =   990
      Width           =   4770
   End
End
Attribute VB_Name = "linea1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset

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
     xcod = dxDBGrid1.Columns.ColumnByFieldName("F7CODCON").value
        xdes = dxDBGrid1.Columns.ColumnByFieldName("F7DESCON").value
        sw_nuevo_mant = False
        nuevalinea.Show 1 ' formulario
        dxDBGrid1.Dataset.Active = False
        dxDBGrid1.Dataset.Active = True
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
    Me.left = 1600
    Me.top = 1050
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_dbbancos
'    Set rst = New ADODB.Recordset
    CargarNivel
    
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    Me.MousePointer = vbDefault
End Sub

Public Sub CargarNivel()
    Dim sql As String
    sql = "select * from sf7nivel01 order by f7codcon"
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = sql
    dxDBGrid1.Dataset.Active = True
    'dxDBGrid1.Dataset.ADODataset.Requery
    'dxDBGrid1.Dataset.Refresh
    dxDBGrid1.KeyField = "f7codcon"
End Sub

Private Sub Form_Unload(Cancel As Integer)
dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo":
        sw_nuevo_mant = True
        nuevalinea.Show 1
    Case "ID_Imprimir":
        With Acr_Linea
            .DataControl1.ConnectionString = cnn_dbbancos
            .DataControl1.Source = "Select * From sf7nivel01 order by f7codcon"
            .fldFecha.Text = Format(Date, "DD/MM/YYYY")
            .lblempresa.Caption = wnomcia
            .Show 1
        End With
    Case "ID_Salir":
        Unload Me
End Select

End Sub

Private Sub txtbusqueda_Change()
   dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "f7codcon LIKE '*" & txtBusqueda.Text & "*' OR " & " f7descon LIKE '*" & txtBusqueda.Text & "*' "
    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "f7codcon LIKE '*" & txtBusqueda.Text & "*' OR " & " f7descon LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
End Sub

