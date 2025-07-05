VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form ayuda_proveedores_ocl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Proveedores"
   ClientHeight    =   5700
   ClientLeft      =   3435
   ClientTop       =   3225
   ClientWidth     =   10815
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
   ScaleHeight     =   5700
   ScaleWidth      =   10815
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9000
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   360
         Width           =   6075
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
      Tools           =   "ayuda_proveedores_ocl.frx":0000
      ToolBars        =   "ayuda_proveedores_ocl.frx":652C
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4245
      Left            =   120
      OleObjectBlob   =   "ayuda_proveedores_ocl.frx":65FA
      TabIndex        =   3
      Top             =   945
      Width           =   10560
   End
End
Attribute VB_Name = "ayuda_proveedores_ocl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn_Mov     As New ADODB.Connection
Dim csql As String

Private Sub dxDBGrid1_OnDblClick()
    wcodprov = dxDBGrid1.Columns.ColumnByFieldName("f2codprov").value
    wrucprov = dxDBGrid1.Columns.ColumnByFieldName("f2newruc").value
    wnomprov = dxDBGrid1.Columns.ColumnByFieldName("f2nomprov").value
    wdirprov = dxDBGrid1.Columns.ColumnByFieldName("f2dirprov").value
    wfpagoprov = dxDBGrid1.Columns(3).value
    wcontacto = dxDBGrid1.Columns.ColumnByFieldName("f2contacto").value
    wdcto = dxDBGrid1.Columns.ColumnByFieldName("F2TIPDOC").value
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
        
    csql = "select A.F2CODPROV,A.F2NEWRUC,A.F2NOMPROV,A.F2FORPAG,A.F2CONTACTO,iif(A.F2TIPPROV='E','Extranjero','Nacional') as f2tipprov, A.f2dirprov, B.F2DESDOC,A.F2TIPDOC "
    csql = csql & "FROM EF2PROVEEDORES A LEFT JOIN DOCUMENTOS B ON A.F2TIPDOC = B.F2CODDOC order by A.F2NOMPROV"
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F2CODPROV"

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dxDBGrid1_OnDblClick
        Case vbKeyUp
            If dxDBGrid1.Dataset.RecNo = 1 Then
                txtBusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    
    Me.left = 1600
    Me.top = 1050
    
    If Cnn_Mov.State = adStateOpen Then Cnn_Mov.Close
    Cnn_Mov.ConnectionString = cnn_dbbancos
    Cnn_Mov.Open cconexion
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = Cnn_Mov
    End With
    FILL
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    Cnn_Mov.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            Mant_Proveedores.Show 1
                
           Unload Me
        Case "ID_Salir":
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F2CODPROV LIKE '*" & _
    txtBusqueda.Text & "*' OR " & " F2NEWRUC LIKE '*" & _
    txtBusqueda.Text & "*' OR " & " F2NOMPROV like '*" & _
    txtBusqueda.Text & "*' or " & " F2FORPAG like '*" & _
    txtBusqueda.Text & "*' or " & " F2CONTACTO like '*" & _
    txtBusqueda.Text & "*' or " & " f2tipprov like '*" & txtBusqueda.Text & "*' "
     
      If Len(Trim(txtBusqueda.Text)) = 0 Then
                dxDBGrid1.Dataset.Filtered = False
      End If
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dxDBGrid1.SetFocus
    End Select
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If Len(Trim(txtBusqueda.Text)) > 0 Then
       dxDBGrid1.Dataset.Filtered = True
       'dxDBGrid1.Dataset.Filter = "F2CODALM LIKE '*" & txtbusqueda.Text & "*' OR " & " F2NOMALM LIKE '*" & txtbusqueda.Text & "*' "
       dxDBGrid1.Dataset.Filter = "F2CODPROV LIKE '*" & _
       txtBusqueda.Text & "*' OR " & " F2NEWRUC LIKE '*" & _
       txtBusqueda.Text & "*' OR " & " F2NOMPROV like '*" & _
       txtBusqueda.Text & "*' or " & " F2FORPAG like '*" & _
       txtBusqueda.Text & "*' or " & " F2CONTACTO like '*" & _
       txtBusqueda.Text & "*' or " & " f2tipprov like '*" & txtBusqueda.Text & "*' "
       dxDBGrid1.Dataset.ADODataset.Requery
       dxDBGrid1.Columns.FocusedIndex = 2
       
     Else
       dxDBGrid1.Dataset.Filtered = False
     End If
 End If
End Sub




