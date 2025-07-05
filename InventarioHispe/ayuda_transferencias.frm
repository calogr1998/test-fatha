VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_transferencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Transferencias"
   ClientHeight    =   5775
   ClientLeft      =   315
   ClientTop       =   1905
   ClientWidth     =   9585
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
   ScaleHeight     =   5775
   ScaleWidth      =   9585
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9310
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   7800
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
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4245
      Left            =   120
      OleObjectBlob   =   "ayuda_transferencias.frx":0000
      TabIndex        =   3
      Top             =   945
      Width           =   9315
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ayuda_transferencias.frx":360D
      ToolBars        =   "ayuda_transferencias.frx":9B39
   End
End
Attribute VB_Name = "ayuda_transferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String

Private Sub dxDBGrid1_OnDblClick()
    wcod_alm = dxDBGrid1.Columns(0).value
    wnomalmacen = dxDBGrid1.Columns(1).value
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

    csql = "Select A.F5MARCA,A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta, " & _
    "C.F7SIGMED,D.F2DESMAR,A.F7CODMED FROM IF5PLA AS A,IF6ALMA AS B, " & _
    "EF7MEDIDAS AS C,EF2MARCAS AS D WHERE A.F5CODPRO=B.F5CODPRO AND " & _
    "B.F2CODALM='" & wcod_alm & "' AND A.F7CODMED=C.F7CODMED AND " & _
    "A.F5MARCA=D.F2CODMAR ORDER BY A.F5NOMPRO ASC"
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "f5codpro"
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
    
    Me.left = 1600
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
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            mant_marcas.Show 1
            If sw_mant_ayuda = False Then Unload Me
        Case "ID_Salir":
            Unload Me
    End Select

End Sub

Private Sub txtbusqueda_Change()


'A.F5MARCA,A.f5codpro,A.F5CODFAB,A.F5NOMPRO,A.F5valvta
'C.F7SIGMED,D.F2DESMAR,A.F7CODMED

  dxDBGrid1.Dataset.Filtered = True
  dxDBGrid1.Dataset.Filter = "F5MARCA LIKE '*" & txtBusqueda.Text & "*' OR " & _
  " f5codpro LIKE '*" & txtBusqueda.Text & "*' or " & _
  " F5CODFAB LIKE '*" & txtBusqueda.Text & "*' or " & _
  " F5NOMPRO LIKE '*" & txtBusqueda.Text & "*' or " & _
  " F5valvta LIKE '*" & txtBusqueda.Text & "*' or " & _
  " F7SIGMED LIKE '*" & txtBusqueda.Text & "*' or " & _
  " F2DESMAR LIKE '*" & txtBusqueda.Text & "*' or " & _
  " F7CODMED LIKE '*" & txtBusqueda.Text & "*' "

 
  

    If Len(Trim(txtBusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
    
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtBusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "F5MARCA LIKE '*" & txtBusqueda.Text & "*' OR " & _
            " f5codpro LIKE '*" & txtBusqueda.Text & "*' or " & _
            " F5CODFAB LIKE '*" & txtBusqueda.Text & "*' or " & _
            " F5NOMPRO LIKE '*" & txtBusqueda.Text & "*' or " & _
            " F5valvta LIKE '*" & txtBusqueda.Text & "*' or " & _
            " F7SIGMED LIKE '*" & txtBusqueda.Text & "*' or " & _
            " F2DESMAR LIKE '*" & txtBusqueda.Text & "*' or " & _
            " F7CODMED LIKE '*" & txtBusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub




