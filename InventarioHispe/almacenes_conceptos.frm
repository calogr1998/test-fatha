VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form Almacenes_Conceptos 
   Caption         =   "Almacenes por Concepto"
   ClientHeight    =   5160
   ClientLeft      =   3495
   ClientTop       =   2670
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   5850
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4380
      Left            =   90
      OleObjectBlob   =   "almacenes_conceptos.frx":0000
      TabIndex        =   0
      Top             =   135
      Width           =   5625
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "almacenes_conceptos.frx":1BD3
      ToolBars        =   "almacenes_conceptos.frx":80FF
   End
End
Attribute VB_Name = "Almacenes_Conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim csql As String
Dim sw_nuevo As Boolean
Dim rsOri As ADODB.Recordset



Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    If dxDBGrid1.Columns.FocusedColumn.ObjectName = "check" Then
        dxDBGrid1.Dataset.Edit
        If dxDBGrid1.Columns.FocusedColumn.value = True Then
            dxDBGrid1.Dataset.FieldValues("opc") = True
        Else
            dxDBGrid1.Dataset.FieldValues("opc") = False
        End If
        dxDBGrid1.Dataset.Post
        
    End If
End Sub

Private Sub Form_Load()
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False

    Me.left = 3600
    Me.top = 1050

    cnombase = "TEMPLUS.mdb"
    cnomtabla = "tmpAlmacenesporConcepto"
    
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    
    sql = "delete * from " & cnomtabla
    cnn_form.Execute sql
    'AlmacenaQuery_sql sql, cnn_form
    
    FILL

End Sub

Private Sub FILL()
Dim J As Integer
sw_nuevo = True


    Set rsOri = New ADODB.Recordset

    csql = "SELECT ef2almacenes.f2codalm,ef2almacenes.f2nomalm" & _
           ", IIF((select almacen_concepto.f1codori " & _
           "from almacen_concepto where ef2almacenes.f2codalm = almacen_concepto.f2codalm and " & _
           "almacen_concepto.f1codori = '" & lista_conceptos.dxDBGrid1.Columns.ColumnByFieldName("f1codori").value & "') <>'',true,false) as OPC " & _
           "FROM ef2almacenes;"
    
    If rsOri.State = adStateOpen Then rsOri.Close
    rsOri.Open csql, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rsOri.EOF Then
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    
        rsOri.MoveFirst
            Do While Not rsOri.EOF
                csql = "insert into " & cnomtabla & " (f2codalm,f2nomalm,opc) values " & _
                       "('" & rsOri.Fields("f2codalm").value & "','" & rsOri.Fields("f2nomalm").value & "'," & rsOri.Fields("opc").value & ")"
                cnn_form.Execute csql
                'AlmacenaQuery_sql csql, cnn_form
                rsOri.MoveNext
            Loop
        
        dxDBGrid1.Dataset.Active = False
        dxDBGrid1.Dataset.ADODataset.CommandText = "select * from " & cnomtabla
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.KeyField = "F2CODALM"

        
    End If
    rsOri.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Dim J As Integer
    
    Select Case Tool.Id
        Case "ID_Grabar":
            
            csql = "delete * from almacen_concepto where f1codori = '" & lista_conceptos.dxDBGrid1.Columns.ColumnByFieldName("f1codori").value & "'"
            cnn_dbbancos.Execute csql
            'AlmacenaQuery_sql csql, cnn_dbbancos
            
            dxDBGrid1.Dataset.First
            Do While Not dxDBGrid1.Dataset.EOF
                If dxDBGrid1.Columns.ColumnByFieldName("OPC").value = True Then
                    csql = "insert into almacen_concepto (f2codalm,f1codori) values ('" & dxDBGrid1.Columns.ColumnByFieldName("f2codalm").value & "'" & _
                           ",'" & lista_conceptos.dxDBGrid1.Columns.ColumnByFieldName("f1codori").value & "')"
                    cnn_dbbancos.Execute csql
                    'AlmacenaQuery_sql csql, cnn_dbbancos
                End If
                dxDBGrid1.Dataset.Next
            Loop

        Case "ID_Salir"
            Unload Me
    End Select
End Sub


