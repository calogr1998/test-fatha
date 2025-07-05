VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngCodigo                   As Long
Private strDescripcion              As String

Private strSQLSelectAlter           As String

Private rstCategoria            As ADODB.Recordset
Private strSQLCategoria         As String


'Propiedad Codigo de Categoria
Public Property Let Codigo(ByVal value As Long)
    lngCodigo = value
End Property

Public Property Get Codigo() As Long
    Codigo = lngCodigo
End Property

'Propiedad Descripcion de Categoria
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
    lngCodigo = 0
    strDescripcion = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerCategoria() As Boolean
    On Error GoTo errObtenerCategoria
    
    Set rstCategoria = New ADODB.Recordset
    
    strSQLCategoria = vbNullString
    strSQLCategoria = strSQLCategoria & "SELECT "
    strSQLCategoria = strSQLCategoria & "* "
    strSQLCategoria = strSQLCategoria & "FROM "
    strSQLCategoria = strSQLCategoria & "CATEGORIA "
    strSQLCategoria = strSQLCategoria & "WHERE "
    strSQLCategoria = strSQLCategoria & "INTCODCATEGORIA = " & lngCodigo
    
    If rstCategoria.State = 1 Then rstCategoria.Close
    
    rstCategoria.Open strSQLCategoria, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCategoria.EOF Then
        lngCodigo = Val(rstCategoria!INTCODCATEGORIA & "")
        strDescripcion = Trim(rstCategoria!STRDESCATEGORIA & "")
        
        obtenerCategoria = True
    Else
        obtenerCategoria = False
    End If
    
    rstCategoria.Close
    
    Set rstCategoria = Nothing
    
    strSQLCategoria = vbNullString
    
    Exit Function
errObtenerCategoria:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCategoria: ObtenerCategoria"
    End Select
    
    obtenerCategoria = False
    
    Err.Clear
End Function

Public Sub obtenerConfigCategoria()
    On Error GoTo errObtenerConfigCategoria
    Set rstCategoria = New ADODB.Recordset
    
    strSQLCategoria = vbNullString
    strSQLCategoria = strSQLCategoria & "SELECT "
    strSQLCategoria = strSQLCategoria & "* "
    strSQLCategoria = strSQLCategoria & "FROM "
    strSQLCategoria = strSQLCategoria & "CATEGORIA "
    strSQLCategoria = strSQLCategoria & "WHERE "
    strSQLCategoria = strSQLCategoria & "INTCODCATEGORIA = " & lngCodigo
    
    If rstCategoria.State = 1 Then rstCategoria.Close
    
    rstCategoria.Open strSQLCategoria, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCategoria.EOF Then
        lngCodigo = Val(rstCategoria!INTCODCATEGORIA & "")
        strDescripcion = Trim(rstCategoria!STRDESCATEGORIA & "")
    End If
    
    rstCategoria.Close
    
    Set rstCategoria = Nothing
    
    Exit Sub
errObtenerConfigCategoria:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsCategoria: ObtenerConfigCategoria"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstCategoria = New ADODB.Recordset
    
    strSQLCategoria = vbNullString
    strSQLCategoria = strSQLCategoria & "SELECT "
    strSQLCategoria = strSQLCategoria & "CATEGORIA.* "
    strSQLCategoria = strSQLCategoria & "FROM "
    strSQLCategoria = strSQLCategoria & "CATEGORIA "
    strSQLCategoria = strSQLCategoria & "WHERE "
    strSQLCategoria = strSQLCategoria & "INTCODCATEGORIA = " & lngCodigo
    
    If rstCategoria.State = 1 Then rstCategoria.Close
    
    rstCategoria.Open strSQLCategoria, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCategoria.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstCategoria.Close
    
    Set rstCategoria = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCategoria: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarCategoria() As Boolean
    On Error GoTo errGuardarCategoria
    
    'cnn_dbbancos.BeginTrans
    
    strSQLCategoria = vbNullString
    
    If Not verificarExistencia Then
'        lngCodigo = generarCodigo
        
        strSQLCategoria = strSQLCategoria & "INSERT INTO CATEGORIA(STRDESCATEGORIA) "
        
        strSQLCategoria = strSQLCategoria & "VALUES('" & strDescripcion & "')"
    Else
        strSQLCategoria = strSQLCategoria & "UPDATE "
        strSQLCategoria = strSQLCategoria & "CATEGORIA "
        
        strSQLCategoria = strSQLCategoria & "SET "
        strSQLCategoria = strSQLCategoria & "STRDESCATEGORIA = '" & strDescripcion & "' "
        
        strSQLCategoria = strSQLCategoria & "WHERE "
        strSQLCategoria = strSQLCategoria & "INTCODCATEGORIA = " & lngCodigo
    End If
    
    cnn_dbbancos.Execute strSQLCategoria
    
    'cnn_dbbancos.CommitTrans
    
    guardarCategoria = True
    
    Exit Function
errGuardarCategoria:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCategoria: GuardarCategoria"
    End Select
    
    'cnn_dbbancos.RollbackTrans
    
    guardarCategoria = False
    
    Err.Clear
End Function

Public Function eliminarCategoria() As Boolean
    On Error GoTo errEliminarCategoria
    
    If Val(validarUsoRegistro(cnn_dbbancos, "INTCODCATEGORIA", "EF2PROVEEDORES", lngCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarCategoria = False
        
        Exit Function
    End If
    
    'cnn_dbbancos.BeginTrans
    
    strSQLCategoria = "DELETE FROM CATEGORIA " & _
                    "WHERE INTCODCATEGORIA = '" & lngCodigo & "'"
    
    cnn_dbbancos.Execute strSQLCategoria
    
    'cnn_dbbancos.CommitTrans
    
    eliminarCategoria = True
    
    Exit Function
errEliminarCategoria:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCategoria: EliminarCategoria"
    End Select
    
    'cnn_dbbancos.RollbackTrans
    
    eliminarCategoria = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Categorias en Grilla (QuamtumGrid)
Public Sub vistaCategoria(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaCategoria
    
    strSQLCategoria = vbNullString
    strSQLCategoria = strSQLCategoria & "SELECT "
    strSQLCategoria = strSQLCategoria & "D.INTCODCATEGORIA, "
    strSQLCategoria = strSQLCategoria & "D.STRDESCATEGORIA "
    
    strSQLCategoria = strSQLCategoria & "FROM "
    strSQLCategoria = strSQLCategoria & "CATEGORIA AS D "
    
    strSQLCategoria = strSQLCategoria & "ORDER BY "
    strSQLCategoria = strSQLCategoria & "D.INTCODCATEGORIA"
    
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
                .FieldName = "INTCODCATEGORIA"
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
                .FieldName = "STRDESCATEGORIA"
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
            .Dataset.ADODataset.CommandText = strSQLCategoria
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "INTCODCATEGORIA"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLCategoria = vbNullString
    
    Exit Sub
errVistaCategoria:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCategoria: VistaCategoria"
    End Select
    
    Err.Clear
End Sub

Public Sub listarCategoria(ByVal comboList As Object)
    On Error GoTo errListarCategoria
    
    Set rstCategoria = New ADODB.Recordset
    
    strSQLCategoria = "SELECT INTCODCATEGORIA, STRDESCATEGORIA " & _
                        "FROM CATEGORIA " & _
                        "ORDER BY INTCODCATEGORIA"
    
    If rstCategoria.State = 1 Then rstCategoria.Close
    
    rstCategoria.Open strSQLCategoria, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstCategoria.EOF Then
        
        rstCategoria.MoveFirst
        
        Do While Not rstCategoria.EOF
            comboList.AddItem Trim(rstCategoria!STRDESCATEGORIA & "") & Space(100) & "*" & Trim(rstCategoria!INTCODCATEGORIA & "")
            
            rstCategoria.MoveNext
        Loop
    End If
    
    If rstCategoria.State = 1 Then rstCategoria.Close
    
    strSQLCategoria = vbNullString
    
    Set rstCategoria = Nothing
    
    Exit Sub
errListarCategoria:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCategoria: ListarCategoria"
    
    Err.Clear
End Sub
