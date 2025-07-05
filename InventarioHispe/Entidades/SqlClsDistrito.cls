VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsDistrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strDescripcion              As String

Private strCodExterno1              As String
Private strCodExterno2              As String

Private strSQLSelectAlter           As String

Private rstDistrito            As ADODB.Recordset
Private strSQLDistrito         As String


'Propiedad Codigo de Distrito
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Descripcion de Distrito
Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property



'Propiedad Codigo de Distrito 1
Public Property Let CodigoExterno1(ByVal value As String)
    strCodExterno1 = value
End Property

Public Property Get CodigoExterno1() As String
    CodigoExterno1 = strCodExterno1
End Property

'Propiedad Codigo de Distrito 2
Public Property Let CodigoExterno2(ByVal value As String)
    strCodExterno2 = value
End Property

Public Property Get CodigoExterno2() As String
    CodigoExterno2 = strCodExterno2
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
    
    strCodExterno1 = vbNullString
    strCodExterno2 = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerDistrito() As Boolean
    On Error GoTo errObtenerDistrito
    
    Set rstDistrito = New ADODB.Recordset
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "SELECT "
    strSQLDistrito = strSQLDistrito & "* "
    strSQLDistrito = strSQLDistrito & "FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
    strSQLDistrito = strSQLDistrito & "WHERE "
    strSQLDistrito = strSQLDistrito & "F2CODZON = '" & strCodigo & "'"
    
    If rstDistrito.State = 1 Then rstDistrito.Close
    
    rstDistrito.Open strSQLDistrito, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDistrito.EOF Then
        strCodigo = Trim(rstDistrito!F2CODZON & "")
        strDescripcion = Trim(rstDistrito!F2DESZON & "")
        
        strCodExterno1 = Trim(rstDistrito!F2CODUSEREXTERNO1 & "")
        strCodExterno2 = Trim(rstDistrito!F2CODUSEREXTERNO2 & "")
        
        obtenerDistrito = True
    Else
        obtenerDistrito = False
    End If
    
    rstDistrito.Close
    
    Set rstDistrito = Nothing
    
    Exit Function
errObtenerDistrito:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsDistrito: ObtenerDistrito"
    End Select
    
    obtenerDistrito = False
    
    Err.Clear
End Function

Public Sub obtenerConfigDistrito()
    On Error GoTo errObtenerConfigDistrito
    Set rstDistrito = New ADODB.Recordset
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "SELECT "
    strSQLDistrito = strSQLDistrito & "* "
    strSQLDistrito = strSQLDistrito & "FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
    strSQLDistrito = strSQLDistrito & "WHERE "
    strSQLDistrito = strSQLDistrito & "F2CODZON = '" & strCodigo & "'"
        
    If rstDistrito.State = 1 Then rstDistrito.Close
    
    rstDistrito.Open strSQLDistrito, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDistrito.EOF Then
        strCodigo = Trim(rstDistrito!F2CODZON & "")
        strDescripcion = Trim(rstDistrito!F2DESZON & "")
        
        strCodExterno1 = Trim(rstDistrito!F2CODUSEREXTERNO1 & "")
        strCodExterno2 = Trim(rstDistrito!F2CODUSEREXTERNO2 & "")
    End If
    
    rstDistrito.Close
    
    Set rstDistrito = Nothing
    
    Exit Sub
errObtenerConfigDistrito:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - SqlClsDistrito: ObtenerConfigDistrito"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstDistrito = New ADODB.Recordset
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "SELECT "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS.* "
    strSQLDistrito = strSQLDistrito & "FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
    strSQLDistrito = strSQLDistrito & "WHERE "
    strSQLDistrito = strSQLDistrito & "F2CODZON = '" & strCodigo & "'"
    
    If rstDistrito.State = 1 Then rstDistrito.Close
    
    rstDistrito.Open strSQLDistrito, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstDistrito.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstDistrito.Close
    
    Set rstDistrito = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsDistrito: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function generarCodigo() As String
    On Error GoTo errGenerarCodigo
    
    Dim rsDistrito As New ADODB.Recordset
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "SELECT "
    strSQLDistrito = strSQLDistrito & "F2CODZON "
    strSQLDistrito = strSQLDistrito & "FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
    strSQLDistrito = strSQLDistrito & "ORDER BY "
    strSQLDistrito = strSQLDistrito & "F2CODZON DESC"
    
    If rsDistrito.State = 1 Then rsDistrito.Close
    
    rsDistrito.Open strSQLDistrito, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rsDistrito.EOF Then
        generarCodigo = Format(Val(strCodigo) + 1, "00")
    Else
        generarCodigo = "01"
    End If
    
    rsDistrito.Close
    
    Set rsDistrito = Nothing
    
    Exit Function
errGenerarCodigo:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "TipoCadena: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsDistrito: GenerarCodigo"
    End Select
        
    generarCodigo = vbNullString
    
    Err.Clear
End Function

Public Function guardarDistrito() As Boolean
    On Error GoTo errGuardarDistrito

    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigo
        End If
        
        strSQLDistrito = vbNullString
        strSQLDistrito = strSQLDistrito & "INSERT INTO MAESTROS.EF2ZONAS("
        strSQLDistrito = strSQLDistrito & "F2CODZON, F2DESZON, F2CODUSEREXTERNO1, F2CODUSEREXTERNO2"
        strSQLDistrito = strSQLDistrito & ") "
        
        strSQLDistrito = strSQLDistrito & "VALUES("
        strSQLDistrito = strSQLDistrito & "'" & strCodigo & "', "
        strSQLDistrito = strSQLDistrito & "'" & strDescripcion & "', "
        strSQLDistrito = strSQLDistrito & "'" & strCodExterno1 & "', "
        strSQLDistrito = strSQLDistrito & "'" & strCodExterno2 & "'"
        strSQLDistrito = strSQLDistrito & ")"
    Else
        strSQLDistrito = vbNullString
        strSQLDistrito = strSQLDistrito & "UPDATE "
        strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
        
        strSQLDistrito = strSQLDistrito & "SET "
        strSQLDistrito = strSQLDistrito & "F2DESZON = '" & strDescripcion & "', "
        strSQLDistrito = strSQLDistrito & "F2CODUSEREXTERNO1 = '" & strCodExterno1 & "', "
        strSQLDistrito = strSQLDistrito & "F2CODUSEREXTERNO2 = '" & strCodExterno2 & "' "
        
        strSQLDistrito = strSQLDistrito & "WHERE "
        strSQLDistrito = strSQLDistrito & "F2CODZON = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLDistrito
    
    guardarDistrito = True
    
    Exit Function
errGuardarDistrito:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsDistrito: GuardarDistrito"
    End Select
    
    guardarDistrito = False
    
    Err.Clear
End Function

Public Function eliminarDistrito() As Boolean
    On Error GoTo errEliminarDistrito
    
    If Val(validarUsoRegistro(cnBdCPlus, "F2ZONA", "MAESTROS.EF2CLIENTES", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarDistrito = False
        
        Exit Function
    End If
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "DELETE FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
    strSQLDistrito = strSQLDistrito & "WHERE "
    strSQLDistrito = strSQLDistrito & "F2CODZON = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLDistrito
    
    eliminarDistrito = True
    
    Exit Function
errEliminarDistrito:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsDistrito: EliminarDistrito"
    End Select
    
    eliminarDistrito = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Distritos en Grilla (QuamtumGrid)
Public Sub vistaDistrito(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaDistrito
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "SELECT "
    strSQLDistrito = strSQLDistrito & "D.F2CODZON, "
    strSQLDistrito = strSQLDistrito & "D.F2DESZON "
    
    strSQLDistrito = strSQLDistrito & "FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS AS D "
    
    strSQLDistrito = strSQLDistrito & "ORDER BY "
    strSQLDistrito = strSQLDistrito & "D.F2CODZON"
    
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
                .FieldName = "F2CODZON"
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
                .FieldName = "F2DESZON"
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
            .Dataset.ADODataset.CommandText = strSQLDistrito
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2CODZON"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLDistrito = vbNullString
    
    Exit Sub
errVistaDistrito:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsDistrito: VistaDistrito"
    End Select
    
    Err.Clear
End Sub

Public Sub listarDistrito(ByVal comboList As Object, _
                            Optional ByVal bolSoloDistritosConCodExterno2 As Boolean)
                            
    On Error GoTo errListarDistrito
    
    Set rstDistrito = New ADODB.Recordset
    
    strSQLDistrito = vbNullString
    strSQLDistrito = strSQLDistrito & "SELECT "
    strSQLDistrito = strSQLDistrito & "F2CODZON, "
    strSQLDistrito = strSQLDistrito & "F2DESZON "
    strSQLDistrito = strSQLDistrito & "FROM "
    strSQLDistrito = strSQLDistrito & "MAESTROS.EF2ZONAS "
    strSQLDistrito = strSQLDistrito & "WHERE "
    strSQLDistrito = strSQLDistrito & "TRIM(F2CODZON & '') <> '' "
        
        If bolSoloDistritosConCodExterno2 Then
            strSQLDistrito = strSQLDistrito & "AND TRIM(F2CODUSEREXTERNO1 & '') <> '' "
        End If
        
    strSQLDistrito = strSQLDistrito & "ORDER BY "
    strSQLDistrito = strSQLDistrito & "F2DESZON"
    
    If rstDistrito.State = 1 Then rstDistrito.Close
    
    rstDistrito.Open strSQLDistrito, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstDistrito.EOF Then
        
        rstDistrito.MoveFirst
        
        Do While Not rstDistrito.EOF
            comboList.AddItem Trim(rstDistrito!F2DESZON & "") & Space(100) & Trim(rstDistrito!F2CODZON & "")
            
            rstDistrito.MoveNext
        Loop
    End If
    
    If rstDistrito.State = 1 Then rstDistrito.Close
    
    strSQLDistrito = vbNullString
    
    Set rstDistrito = Nothing
    
    Exit Sub
errListarDistrito:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsDistrito: ListarDistrito"
    
    Err.Clear
End Sub

