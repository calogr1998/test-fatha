VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSectorEmpresarial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strDescripcion              As String

Private strSQLSelectAlter           As String

Private rstSector            As ADODB.Recordset
Private strSQLSector         As String


'Propiedad Codigo de Sector
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Descripcion de Sector
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

Public Function obtenerSector() As Boolean
    On Error GoTo errObtenerSector
    
    Set rstSector = New ADODB.Recordset
    
    strSQLSector = "SELECT * " & _
                    "FROM EF2TIPOS " & _
                    "WHERE CODTIPCLIE = '" & strCodigo & "'"
    
    If rstSector.State = 1 Then rstSector.Close
    
    rstSector.Open strSQLSector, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSector.EOF Then
        strCodigo = Trim(rstSector!CODTIPCLIE & "")
        strDescripcion = Trim(rstSector!DESTIPCLIE & "")
        
        obtenerSector = True
    Else
        obtenerSector = False
    End If
    
    rstSector.Close
    
    Set rstSector = Nothing
    
    Exit Function
errObtenerSector:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsSectorEmpresarial: ObtenerSector"
    End Select
    
    obtenerSector = False
    
    Err.Clear
End Function

Public Sub obtenerConfigSector()
    On Error GoTo errObtenerConfigSector
    Set rstSector = New ADODB.Recordset
    
    strSQLSector = "SELECT * " & _
                        "FROM EF2TIPOS " & _
                        "WHERE CODTIPCLIE = '" & strCodigo & "'"
        
    If rstSector.State = 1 Then rstSector.Close
    
    rstSector.Open strSQLSector, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSector.EOF Then
        strCodigo = Trim(rstSector!CODTIPCLIE & "")
        strDescripcion = Trim(rstSector!DESTIPCLIE & "")
    End If
    
    rstSector.Close
    
    Set rstSector = Nothing
    
    Exit Sub
errObtenerConfigSector:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsSectorEmpresarial: ObtenerConfigSector"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstSector = New ADODB.Recordset
    
    strSQLSector = "SELECT EF2TIPOS.* " & _
                    "FROM EF2TIPOS " & _
                    "WHERE CODTIPCLIE = '" & strCodigo & "'"
    
    If rstSector.State = 1 Then rstSector.Close
    
    rstSector.Open strSQLSector, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstSector.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstSector.Close
    
    Set rstSector = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsSectorEmpresarial: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function generarCodigo() As String
    On Error GoTo errGenerarCodigo
    
    Dim rsSector As New ADODB.Recordset
    
    strSQLSector = "SELECT CODTIPCLIE " & _
                        "FROM EF2TIPOS " & _
                        "ORDER BY CODTIPCLIE DESC"
    
    If rsSector.State = 1 Then rsSector.Close
    
    rsSector.Open strSQLSector, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rsSector.EOF Then
        generarCodigo = Format(Val(strCodigo) + 1, "00")
    Else
        generarCodigo = "01"
    End If
    
    rsSector.Close
    
    Set rsSector = Nothing
    
    Exit Function
errGenerarCodigo:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "TipoCadena: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsSectorEmpresarial: GenerarCodigo"
    End Select
        
    generarCodigo = vbNullString
    
    Err.Clear
End Function

Public Function guardarSector() As Boolean
    On Error GoTo errGuardarSector
    
    'cnn_dbbancos.BeginTrans
    
    strSQLSector = vbNullString
    
    If Not verificarExistencia Then
        strCodigo = generarCodigo
        
        strSQLSector = strSQLSector & "INSERT INTO EF2TIPOS(CODTIPCLIE, DESTIPCLIE) "
        
        strSQLSector = strSQLSector & "VALUES('" & strCodigo & "', '" & strDescripcion & "')"
    Else
        strSQLSector = strSQLSector & "UPDATE "
        strSQLSector = strSQLSector & "EF2TIPOS "
        
        strSQLSector = strSQLSector & "SET "
        strSQLSector = strSQLSector & "DESTIPCLIE = '" & strDescripcion & "' "
        
        strSQLSector = strSQLSector & "WHERE "
        strSQLSector = strSQLSector & "CODTIPCLIE = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLSector
    
    'cnn_dbbancos.CommitTrans
    
    guardarSector = True
    
    Exit Function
errGuardarSector:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsSectorEmpresarial: GuardarSector"
    End Select
    
    'cnn_dbbancos.RollbackTrans
    
    guardarSector = False
    
    Err.Clear
End Function

Public Function eliminarSector() As Boolean
    On Error GoTo errEliminarSector
    
    If Val(validarUsoRegistro(cnn_dbbancos, "F2TIPOCLIE", "EF2CLIENTES", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarSector = False
        
        Exit Function
    End If
    
    'cnn_dbbancos.BeginTrans
    
    strSQLSector = "DELETE FROM EF2TIPOS " & _
                    "WHERE CODTIPCLIE = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLSector
    
    'cnn_dbbancos.CommitTrans
    
    eliminarSector = True
    
    Exit Function
errEliminarSector:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsSectorEmpresarial: EliminarSector"
    End Select
    
    'cnn_dbbancos.RollbackTrans
    
    eliminarSector = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Sectors en Grilla (QuamtumGrid)
Public Sub vistaSector(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaSector
    
    strSQLSector = vbNullString
    strSQLSector = strSQLSector & "SELECT "
    strSQLSector = strSQLSector & "D.CODTIPCLIE, "
    strSQLSector = strSQLSector & "D.DESTIPCLIE "
    
    strSQLSector = strSQLSector & "FROM "
    strSQLSector = strSQLSector & "EF2TIPOS AS D "
    
    strSQLSector = strSQLSector & "ORDER BY "
    strSQLSector = strSQLSector & "D.CODTIPCLIE"
    
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
                .FieldName = "CODTIPCLIE"
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
                .FieldName = "DESTIPCLIE"
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
            .Dataset.ADODataset.CommandText = strSQLSector
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODTIPCLIE"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLSector = vbNullString
    
    Exit Sub
errVistaSector:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsSectorEmpresarial: VistaSector"
    End Select
    
    Err.Clear
End Sub

Public Sub listarSector(ByVal comboList As Object)
    On Error GoTo errListarSector
    
    Set rstSector = New ADODB.Recordset
    
    strSQLSector = "SELECT CODTIPCLIE, DESTIPCLIE " & _
                        "FROM EF2TIPOS " & _
                        "ORDER BY CODTIPCLIE"
    
    If rstSector.State = 1 Then rstSector.Close
    
    rstSector.Open strSQLSector, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstSector.EOF Then
        
        rstSector.MoveFirst
        
        Do While Not rstSector.EOF
            comboList.AddItem Trim(rstSector!DESTIPCLIE & "") & Space(100) & Trim(rstSector!CODTIPCLIE & "")
            
            rstSector.MoveNext
        Loop
    End If
    
    If rstSector.State = 1 Then rstSector.Close
    
    strSQLSector = vbNullString
    
    Set rstSector = Nothing
    
    Exit Sub
errListarSector:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsSectorEmpresarial: ListarSector"
    
    Err.Clear
End Sub
