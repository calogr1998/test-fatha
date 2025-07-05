VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strDescripcion              As String

Private strSQLSelectAlter           As String

Private rstTarea            As ADODB.Recordset
Private strSQLTarea         As String


'Propiedad Codigo de Tarea
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Descripcion de Tarea
Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property



'Propiedad SQLSelectAlter
Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    strDescripcion = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerTarea() As Boolean
    On Error GoTo errObtenerTarea
    
    Set rstTarea = New ADODB.Recordset
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "SELECT "
    strSQLTarea = strSQLTarea & "* "
    strSQLTarea = strSQLTarea & "FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS "
    strSQLTarea = strSQLTarea & "WHERE "
    strSQLTarea = strSQLTarea & "F2CODTAREA = '" & strCodigo & "'"
    
    If rstTarea.State = 1 Then rstTarea.Close
    
    rstTarea.Open strSQLTarea, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTarea.EOF Then
        strCodigo = Trim(rstTarea!F2CODTAREA & "")
        strDescripcion = Trim(rstTarea!F2DESTAREA & "")
        
        obtenerTarea = True
    Else
        obtenerTarea = False
    End If
    
    rstTarea.Close
    
    Set rstTarea = Nothing
    
    Exit Function
errObtenerTarea:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTarea: ObtenerTarea"
    End Select
    
    obtenerTarea = False
    
    Err.Clear
End Function

Public Sub obtenerConfigTarea()
    On Error GoTo errObtenerConfigTarea
    Set rstTarea = New ADODB.Recordset
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "SELECT "
    strSQLTarea = strSQLTarea & "* "
    strSQLTarea = strSQLTarea & "FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS "
    strSQLTarea = strSQLTarea & "WHERE "
    strSQLTarea = strSQLTarea & "F2CODTAREA = '" & strCodigo & "'"
        
    If rstTarea.State = 1 Then rstTarea.Close
    
    rstTarea.Open strSQLTarea, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTarea.EOF Then
        strCodigo = Trim(rstTarea!F2CODTAREA & "")
        strDescripcion = Trim(rstTarea!F2DESTAREA & "")
    End If
    
    rstTarea.Close
    
    Set rstTarea = Nothing
    
    Exit Sub
errObtenerConfigTarea:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsTarea: ObtenerConfigTarea"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstTarea = New ADODB.Recordset
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "SELECT "
    strSQLTarea = strSQLTarea & "EF2TAREAS.* "
    strSQLTarea = strSQLTarea & "FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS "
    strSQLTarea = strSQLTarea & "WHERE "
    strSQLTarea = strSQLTarea & "F2CODTAREA = '" & strCodigo & "'"
    
    If rstTarea.State = 1 Then rstTarea.Close
    
    rstTarea.Open strSQLTarea, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTarea.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstTarea.Close
    
    Set rstTarea = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTarea: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function generarCodigo() As String
    On Error GoTo errGenerarCodigo
    
    Dim rsTarea As New ADODB.Recordset
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "SELECT "
    strSQLTarea = strSQLTarea & "F2CODTAREA "
    strSQLTarea = strSQLTarea & "FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS "
    strSQLTarea = strSQLTarea & "ORDER BY "
    strSQLTarea = strSQLTarea & "F2CODTAREA DESC"
    
    If rsTarea.State = 1 Then rsTarea.Close
    
    rsTarea.Open strSQLTarea, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rsTarea.EOF Then
        generarCodigo = Format(Val(strCodigo) + 1, "0000")
    Else
        generarCodigo = "0001"
    End If
    
    rsTarea.Close
    
    Set rsTarea = Nothing
    
    Exit Function
errGenerarCodigo:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "TipoCadena: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTarea: GenerarCodigo"
    End Select
        
    generarCodigo = vbNullString
    
    Err.Clear
End Function

Public Function guardarTarea() As Boolean
    On Error GoTo errGuardarTarea
    
    strSQLTarea = vbNullString
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigo
        End If
        
        strSQLTarea = vbNullString
        strSQLTarea = strSQLTarea & "INSERT INTO EF2TAREAS("
        strSQLTarea = strSQLTarea & "F2CODTAREA, F2DESTAREA"
        strSQLTarea = strSQLTarea & ") "
        
        strSQLTarea = strSQLTarea & "VALUES("
        strSQLTarea = strSQLTarea & "'" & strCodigo & "', "
        strSQLTarea = strSQLTarea & "'" & strDescripcion & "'"
        strSQLTarea = strSQLTarea & ")"
    Else
        strSQLTarea = strSQLTarea & "UPDATE "
        strSQLTarea = strSQLTarea & "EF2TAREAS "
        
        strSQLTarea = strSQLTarea & "SET "
        strSQLTarea = strSQLTarea & "F2DESTAREA = '" & strDescripcion & "' "
        strSQLTarea = strSQLTarea & "WHERE "
        strSQLTarea = strSQLTarea & "F2CODTAREA = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLTarea
    
    guardarTarea = True
    
    Exit Function
errGuardarTarea:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTarea: GuardarTarea"
    End Select
    
    guardarTarea = False
    
    Err.Clear
End Function

Public Function eliminarTarea() As Boolean
    On Error GoTo errEliminarTarea
    
    If Val(validarUsoRegistro(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarTarea = False
        
        Exit Function
    End If
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "DELETE FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS "
    strSQLTarea = strSQLTarea & "WHERE "
    strSQLTarea = strSQLTarea & "F2CODTAREA = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLTarea
    
    eliminarTarea = True
    
    Exit Function
errEliminarTarea:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTarea: EliminarTarea"
    End Select
    
    eliminarTarea = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Tareas en Grilla (QuamtumGrid)
Public Sub vistaTarea(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaTarea
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "SELECT "
    strSQLTarea = strSQLTarea & "D.F2CODTAREA, "
    strSQLTarea = strSQLTarea & "D.F2DESTAREA "
    
    strSQLTarea = strSQLTarea & "FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS AS D "
    
    strSQLTarea = strSQLTarea & "ORDER BY "
    strSQLTarea = strSQLTarea & "D.F2CODTAREA"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2CODTAREA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
                .Visible = True
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DESTAREA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLTarea
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2CODTAREA"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLTarea = vbNullString
    
    Exit Sub
errVistaTarea:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsTarea: VistaTarea"
    End Select
    
    Err.Clear
End Sub

Public Sub listarTarea(ByVal comboList As Object, _
                            Optional ByVal bolSoloTareasConCodExterno2 As Boolean)
                            
    On Error GoTo errListarTarea
    
    Set rstTarea = New ADODB.Recordset
    
    strSQLTarea = vbNullString
    strSQLTarea = strSQLTarea & "SELECT "
    strSQLTarea = strSQLTarea & "F2CODTAREA, "
    strSQLTarea = strSQLTarea & "F2DESTAREA "
    strSQLTarea = strSQLTarea & "FROM "
    strSQLTarea = strSQLTarea & "EF2TAREAS "
    strSQLTarea = strSQLTarea & "WHERE "
    strSQLTarea = strSQLTarea & "TRIM(F2CODTAREA & '') <> '' "
        
        If bolSoloTareasConCodExterno2 Then
            strSQLTarea = strSQLTarea & "AND TRIM(F2CODUSEREXTERNO1 & '') <> '' "
        End If
        
    strSQLTarea = strSQLTarea & "ORDER BY "
    strSQLTarea = strSQLTarea & "F2DESTAREA"
    
    If rstTarea.State = 1 Then rstTarea.Close
    
    rstTarea.Open strSQLTarea, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstTarea.EOF Then
        
        rstTarea.MoveFirst
        
        Do While Not rstTarea.EOF
            comboList.AddItem Trim(rstTarea!F2DESTAREA & "") & Space(100) & Trim(rstTarea!F2CODTAREA & "")
            
            rstTarea.MoveNext
        Loop
    End If
    
    If rstTarea.State = 1 Then rstTarea.Close
    
    strSQLTarea = vbNullString
    
    Set rstTarea = Nothing
    
    Exit Sub
errListarTarea:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsTarea: ListarTarea"
    
    Err.Clear
End Sub

