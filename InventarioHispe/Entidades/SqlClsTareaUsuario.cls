VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsTareaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigoUsuario            As String
Private strCodigoTarea              As String

Private strSQLSelectAlter           As String

Private rstTareaUsuario            As ADODB.Recordset
Private strSQLTareaUsuario         As String


'Propiedad Codigo de Usuario
Public Property Let CodigoUsuario(ByVal value As String)
    strCodigoUsuario = value
End Property

Public Property Get CodigoUsuario() As String
    CodigoUsuario = strCodigoUsuario
End Property

'Propiedad Codigo de Tarea
Public Property Let CodigoTarea(ByVal value As String)
    strCodigoTarea = value
End Property

Public Property Get CodigoTarea() As String
    CodigoTarea = strCodigoTarea
End Property




'Propiedad SQLSelectAlter
Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigoUsuario = vbNullString
    strCodigoTarea = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerTareaUsuario() As Boolean
    On Error GoTo errObtenerTareaUsuario
    
    Set rstTareaUsuario = New ADODB.Recordset
    
    strSQLTareaUsuario = vbNullString
    strSQLTareaUsuario = strSQLTareaUsuario & "SELECT "
    strSQLTareaUsuario = strSQLTareaUsuario & "* "
    strSQLTareaUsuario = strSQLTareaUsuario & "FROM "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS "
    strSQLTareaUsuario = strSQLTareaUsuario & "WHERE "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER = '" & strCodigoUsuario & "' AND "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA = '" & strCodigoTarea & "'"
    
    If rstTareaUsuario.State = 1 Then rstTareaUsuario.Close
    
    rstTareaUsuario.Open strSQLTareaUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTareaUsuario.EOF Then
        strCodigoUsuario = Trim(rstTareaUsuario!F2CODUSER & "")
        strCodigoTarea = Trim(rstTareaUsuario!F2CODTAREA & "")
        
        obtenerTareaUsuario = True
    Else
        obtenerTareaUsuario = False
    End If
    
    rstTareaUsuario.Close
    
    Set rstTareaUsuario = Nothing
    
    Exit Function
errObtenerTareaUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTareaUsuario: ObtenerTareaUsuario"
    End Select
    
    obtenerTareaUsuario = False
    
    Err.Clear
End Function

Public Sub obtenerConfigTareaUsuario()
    On Error GoTo errObtenerConfigTareaUsuario
    Set rstTareaUsuario = New ADODB.Recordset
    
    strSQLTareaUsuario = vbNullString
    strSQLTareaUsuario = strSQLTareaUsuario & "SELECT "
    strSQLTareaUsuario = strSQLTareaUsuario & "* "
    strSQLTareaUsuario = strSQLTareaUsuario & "FROM "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS "
    strSQLTareaUsuario = strSQLTareaUsuario & "WHERE "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER = '" & strCodigoUsuario & "' AND "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA = '" & strCodigoTarea & "'"
        
    If rstTareaUsuario.State = 1 Then rstTareaUsuario.Close
    
    rstTareaUsuario.Open strSQLTareaUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTareaUsuario.EOF Then
        strCodigoUsuario = Trim(rstTareaUsuario!F2CODUSER & "")
        strCodigoTarea = Trim(rstTareaUsuario!F2CODTAREA & "")
    End If
    
    rstTareaUsuario.Close
    
    Set rstTareaUsuario = Nothing
    
    Exit Sub
errObtenerConfigTareaUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - SqlClsTareaUsuario: ObtenerConfigTareaUsuario"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstTareaUsuario = New ADODB.Recordset
    
    strSQLTareaUsuario = vbNullString
    strSQLTareaUsuario = strSQLTareaUsuario & "SELECT "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS.* "
    strSQLTareaUsuario = strSQLTareaUsuario & "FROM "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS "
    strSQLTareaUsuario = strSQLTareaUsuario & "WHERE "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER = '" & strCodigoUsuario & "' AND "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA = '" & strCodigoTarea & "'"
    
    If rstTareaUsuario.State = 1 Then rstTareaUsuario.Close
    
    rstTareaUsuario.Open strSQLTareaUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTareaUsuario.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstTareaUsuario.Close
    
    Set rstTareaUsuario = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTareaUsuario: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarTareaUsuario() As Boolean
    On Error GoTo errGuardarTareaUsuario
    
    strSQLTareaUsuario = vbNullString
    
    If Not verificarExistencia Then
        strSQLTareaUsuario = vbNullString
        strSQLTareaUsuario = strSQLTareaUsuario & "INSERT INTO MAESTROS.EF2TAREAUSERS("
        strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER, F2CODTAREA"
        strSQLTareaUsuario = strSQLTareaUsuario & ") "
        
        strSQLTareaUsuario = strSQLTareaUsuario & "VALUES("
        strSQLTareaUsuario = strSQLTareaUsuario & "'" & strCodigoUsuario & "', "
        strSQLTareaUsuario = strSQLTareaUsuario & "'" & strCodigoTarea & "'"
        strSQLTareaUsuario = strSQLTareaUsuario & ")"
    Else
        strSQLTareaUsuario = strSQLTareaUsuario & "UPDATE "
        strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS "
        
        strSQLTareaUsuario = strSQLTareaUsuario & "SET "
        strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER = '" & strCodigoUsuario & "', "
        strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA = '" & strCodigoTarea & "' "
        strSQLTareaUsuario = strSQLTareaUsuario & "WHERE "
        strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER = '" & strCodigoUsuario & "' AND "
        strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA = '" & strCodigoTarea & "'"
    End If
    
    cnBdCPlus.Execute strSQLTareaUsuario
    
    guardarTareaUsuario = True
    
    Exit Function
errGuardarTareaUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTareaUsuario: GuardarTareaUsuario"
    End Select
    
    guardarTareaUsuario = False
    
    Err.Clear
End Function

Public Function eliminarTareaUsuario() As Boolean
    On Error GoTo errEliminarTareaUsuario
    
    strSQLTareaUsuario = vbNullString
    strSQLTareaUsuario = strSQLTareaUsuario & "DELETE FROM "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS "
    strSQLTareaUsuario = strSQLTareaUsuario & "WHERE "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER = '" & strCodigoUsuario & "' AND "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA = '" & strCodigoTarea & "'"
    
    cnBdCPlus.Execute strSQLTareaUsuario
    
    eliminarTareaUsuario = True
    
    Exit Function
errEliminarTareaUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTareaUsuario: EliminarTareaUsuario"
    End Select
    
    eliminarTareaUsuario = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de TareaUsuarios en Grilla (QuamtumGrid)
Public Sub vistaTareaUsuario(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaTareaUsuario
    
    strSQLTareaUsuario = vbNullString
    strSQLTareaUsuario = strSQLTareaUsuario & "SELECT "
    strSQLTareaUsuario = strSQLTareaUsuario & "D.F2CODUSER, "
    strSQLTareaUsuario = strSQLTareaUsuario & "D.F2CODTAREA "
    
    strSQLTareaUsuario = strSQLTareaUsuario & "FROM "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS AS D "
    
    strSQLTareaUsuario = strSQLTareaUsuario & "ORDER BY "
    strSQLTareaUsuario = strSQLTareaUsuario & "D.F2CODUSER"
    
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
                .FieldName = "F2CODUSER"
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
                .FieldName = "F2CODTAREA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLTareaUsuario
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2CODUSER"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLTareaUsuario = vbNullString
    
    Exit Sub
errVistaTareaUsuario:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTareaUsuario: VistaTareaUsuario"
    End Select
    
    Err.Clear
End Sub

Public Sub listarTareaUsuario(ByVal comboList As Object, _
                            Optional ByVal bolSoloTareaUsuariosConCodExterno2 As Boolean)
                            
    On Error GoTo errListarTareaUsuario
    
    Set rstTareaUsuario = New ADODB.Recordset
    
    strSQLTareaUsuario = vbNullString
    strSQLTareaUsuario = strSQLTareaUsuario & "SELECT "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODUSER, "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA "
    strSQLTareaUsuario = strSQLTareaUsuario & "FROM "
    strSQLTareaUsuario = strSQLTareaUsuario & "MAESTROS.EF2TAREAUSERS "
    strSQLTareaUsuario = strSQLTareaUsuario & "WHERE "
    strSQLTareaUsuario = strSQLTareaUsuario & "TRIM(F2CODUSER & '') <> '' "
        
        If bolSoloTareaUsuariosConCodExterno2 Then
            strSQLTareaUsuario = strSQLTareaUsuario & "AND TRIM(F2CODUSEREXTERNO1 & '') <> '' "
        End If
        
    strSQLTareaUsuario = strSQLTareaUsuario & "ORDER BY "
    strSQLTareaUsuario = strSQLTareaUsuario & "F2CODTAREA"
    
    If rstTareaUsuario.State = 1 Then rstTareaUsuario.Close
    
    rstTareaUsuario.Open strSQLTareaUsuario, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstTareaUsuario.EOF Then
        
        rstTareaUsuario.MoveFirst
        
        Do While Not rstTareaUsuario.EOF
            comboList.AddItem Trim(rstTareaUsuario!F2CODTAREA & "") & Space(100) & Trim(rstTareaUsuario!F2CODUSER & "")
            
            rstTareaUsuario.MoveNext
        Loop
    End If
    
    If rstTareaUsuario.State = 1 Then rstTareaUsuario.Close
    
    strSQLTareaUsuario = vbNullString
    
    Set rstTareaUsuario = Nothing
    
    Exit Sub
errListarTareaUsuario:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTareaUsuario: ListarTareaUsuario"
    
    Err.Clear
End Sub

