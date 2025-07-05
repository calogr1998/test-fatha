VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form conceptos_almacenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Conceptos"
   ClientHeight    =   5160
   ClientLeft      =   9120
   ClientTop       =   1935
   ClientWidth     =   5835
   Icon            =   "conceptos_almacenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5835
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4380
      Left            =   90
      OleObjectBlob   =   "conceptos_almacenes.frx":058A
      TabIndex        =   0
      Top             =   135
      Width           =   5670
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "conceptos_almacenes.frx":2608
      ToolBars        =   "conceptos_almacenes.frx":8B34
   End
End
Attribute VB_Name = "conceptos_almacenes"
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
        If dxDBGrid1.Columns.FocusedColumn.Value = True Then
            dxDBGrid1.Dataset.FieldValues("opc") = True
        Else
            dxDBGrid1.Dataset.FieldValues("opc") = False
        End If
        dxDBGrid1.Dataset.Post
        
    End If
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False

    Me.left = 3600
    Me.top = 1050

    cnombase = "TEMPLUS.mdb"
    cnomtabla = "tmpConceptosporAlmacen"
    
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    
    sql = "delete * from " & cnomtabla
    cnn_form.Execute sql
    'AlmacenaQuery_sql sql, cnn_form

    FILL
    
    Me.MousePointer = vbDefault
End Sub

Private Sub FILL()
'    Dim J As Integer
    sw_nuevo = True

'    Set rsOri = New ADODB.Recordset
    
    Rem SK ADD:
    
'    If ctipoadm_bd = "M" Then
'        csql = "SELECT SF1ORIGENES.F1CODORI, SF1ORIGENES.F1NOMORI, SF1ORIGENES.F1TIPMOV, ALMACEN_CONCEPTO.F2CODALM " & _
'               "FROM ALMACEN_CONCEPTO RIGHT JOIN SF1ORIGENES ON ALMACEN_CONCEPTO.F1CODORI = SF1ORIGENES.F1CODORI AND ALMACEN_CONCEPTO.F2CODALM = '" & wcod_alm & "' " & _
'               "WHERE SF1ORIGENES.F1TIPMOV = 'S' " & _
'               "ORDER BY SF1ORIGENES.F1TIPMOV;"
'
'    Else
'        csql = "SELECT SF1ORIGENES.F1CODORI, SF1ORIGENES.F1NOMORI, sf1origenes.f1tipmov" & _
'               ", IIF((select almacen_concepto.f1codori " & _
'               "from almacen_concepto where SF1ORIGENES.F1TIPMOV AND SF1ORIGENES.F1CODORI = almacen_concepto.f1codori and " & _
'               "almacen_concepto.f2codalm = '" & wcod_alm & "') <>'',true,false) as OPC " & _
'               "FROM SF1ORIGENES order by sf1origenes.f1tipmov;"
'    End If
    'tmpConceptosporAlmacen
    'f1codori,f1nomori,f1tipmov,opc
    dxDBGrid1.Dataset.Close
    
    cnn_form.Execute "DELETE FROM TMPCONCEPTOSPORALMACEN"
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "INSERT INTO TMPCONCEPTOSPORALMACEN("
    SqlCad = SqlCad & "F1CODORI, F1NOMORI, F1TIPMOV, OPC) "
    SqlCad = SqlCad & "IN '" & wrutatemp & "TEMPLUS.MDB' "
    
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "ORI.F1CODORI, "
    SqlCad = SqlCad & "ORI.F1NOMORI, "
    SqlCad = SqlCad & "IIF(ORI.F1TIPMOV = 'I', 'Ingreso', 'Salida') AS F1TIPMOV,"
    SqlCad = SqlCad & "IIF(VAL(ALMORI.CANTIDAD & '') = 0, FALSE, TRUE) AS OPC "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "SF1ORIGENES AS ORI "
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "(SELECT "
    SqlCad = SqlCad & "F1CODORI, "
    SqlCad = SqlCad & "COUNT(F1CODORI) AS CANTIDAD "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "ALMACEN_CONCEPTO "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "F2CODALM = '" & wcod_alm & "' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "F1CODORI) AS ALMORI "
    SqlCad = SqlCad & "ON ALMORI.F1CODORI = ORI.F1CODORI "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "ORI.ESTADO = TRUE "
    SqlCad = SqlCad & "ORDER BY "
    SqlCad = SqlCad & "ORI.F1TIPMOV"
    
    cnn_dbbancos.Execute SqlCad
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_form
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = "SELECT * FROM TMPCONCEPTOSPORALMACEN"
        .Dataset.Active = True
        
        .KeyField = "F1CODORI"
        
        .Columns.ColumnByFieldName("F1TIPMOV").Sorted = csUp
    End With
    
'    If rsOri.State = adStateOpen Then rsOri.Close
'
'    rsOri.Open csql, cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
'
'    If Not rsOri.EOF Then
'        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
'
'        rsOri.MoveFirst
'            Do While Not rsOri.EOF
'                csql = "insert into " & cnomtabla & " (f1codori,f1nomori,f1tipmov,opc) values " & _
'                       "('" & rsOri.Fields("f1codori").Value & "','" & rsOri.Fields("f1nomori").Value & "','" & rsOri.Fields("f1tipmov") & "', " & rsOri.Fields("OPC") & ")"
'                cnn_form.Execute csql
'                'AlmacenaQuery_sql csql, cnn_form
'
'                rsOri.MoveNext
'            Loop
'
'        dxDBGrid1.Dataset.Active = False
'        dxDBGrid1.Dataset.ADODataset.CommandText = "select * from " & cnomtabla
'        dxDBGrid1.Dataset.Active = True
'        dxDBGrid1.KeyField = "F1CODORI"
'
'
'    End If
'
'    rsOri.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid1.Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Grabar"
            Me.MousePointer = vbHourglass
            
            guardarConcepto
            
            Me.MousePointer = vbDefault
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Private Sub guardarConcepto()
    Dim rstConcepto As New ADODB.Recordset
    
    If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
        dxDBGrid1.Dataset.Post
    End If
    
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "TMPCONCEPTOSPORALMACEN "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "OPC = TRUE"
    
    If rstConcepto.State = 1 Then rstConcepto.Close
    
    rstConcepto.Open SqlCad, cnn_form, adOpenForwardOnly, adLockReadOnly
    
    If Not rstConcepto.EOF Then
        rstConcepto.MoveFirst
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM ALMACEN_CONCEPTO WHERE F2CODALM = '" & wcod_alm & "'"
        
        cnn_dbbancos.Execute SqlCad
        Actualiza_Log SqlCad, StrConexDbBancos
        
        Do While Not rstConcepto.EOF
            SqlCad = vbNullString
            SqlCad = SqlCad & "INSERT INTO ALMACEN_CONCEPTO("
            SqlCad = SqlCad & "F1CODORI, F2CODALM) "
            SqlCad = SqlCad & "VALUES("
            SqlCad = SqlCad & "'" & Trim(rstConcepto!F1CODORI & "") & "', "
            SqlCad = SqlCad & "'" & wcod_alm & "')"
            
            cnn_dbbancos.Execute SqlCad
            
            Actualiza_Log SqlCad, StrConexDbBancos
            
            rstConcepto.MoveNext
        Loop
    Else
        SqlCad = vbNullString
        SqlCad = SqlCad & "DELETE FROM ALMACEN_CONCEPTO WHERE F2CODALM = '" & wcod_alm & "'"
        
        cnn_dbbancos.Execute SqlCad
        Actualiza_Log SqlCad, StrConexDbBancos
        'MsgBox "No se asociaron correctamente los conceptos, vuelva a intentarlo.", vbInformation + vbOKOnly, App.ProductName
    End If
    
    MsgBox "Conceptos Asociados.", vbInformation + vbOKOnly, App.ProductName
'            If ctipoadm_bd = "M" Then
'                sql = "delete from almacen_concepto where f2codalm = '" & wcod_alm & "'"
'                cnn_dbbancos.Execute sql
'                'AlmacenaQuery_sql sql, cnn_dbbancos
'            Else
'                sql = "delete * from almacen_concepto where f2codalm = '" & wcod_alm & "'"
'                cnn_dbbancos.Execute sql
'                'AlmacenaQuery_sql sql, cnn_dbbancos
'            End If
'
'            J = 1
'            dxDBGrid1.Dataset.First
'            Do While Not dxDBGrid1.Dataset.EOF
'                If dxDBGrid1.Columns.ColumnByFieldName("OPC").Value = True Then
'                    csql = "insert into almacen_concepto (f1codori,f2codalm) values ('" & dxDBGrid1.Columns.ColumnByFieldName("f1codori").Value & "'" & _
'                           ",'" & wcod_alm & "')"
'                    cnn_dbbancos.Execute csql
'                    'AlmacenaQuery_sql csql, cnn_dbbancos
'
'                End If
'                dxDBGrid1.Dataset.Next
'            Loop
End Sub
