VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strCodigoS10                As String
Private strDescripcion              As String
Private lngDias                     As Long
Private strTipo                     As String
Private bolParaLetra                As Boolean

Private strSQLSelectAlter           As String

Private rstFormaPago            As ADODB.Recordset
Private strSQLFormaPago         As String


'Propiedad Codigo de FormaPago
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Codigo de FormaPago en S10
Public Property Let CodigoS10(ByVal value As String)
    strCodigoS10 = value
End Property

Public Property Get CodigoS10() As String
    CodigoS10 = strCodigoS10
End Property

'Propiedad Descripcion de FormaPago
Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

'Propiedad Dias de FormaPago
Public Property Let Dias(ByVal value As Long)
    lngDias = value
End Property

Public Property Get Dias() As Long
    Dias = lngDias
End Property

'Propiedad Tipo de FormaPago
Public Property Let Tipo(ByVal value As String)
    strTipo = value
End Property

Public Property Get Tipo() As String
    Tipo = strTipo
End Property

'ParaLetra
Public Property Let ParaLetra(ByVal value As Boolean)
    bolParaLetra = value
End Property

Public Property Get ParaLetra() As Boolean
    ParaLetra = bolParaLetra
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
    strCodigoS10 = vbNullString
    strDescripcion = vbNullString
    lngDias = 0
    strTipo = vbNullString
    bolParaLetra = False
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerFormaPago() As Boolean
    On Error GoTo errObtenerFormaPago
    
    Set rstFormaPago = New ADODB.Recordset
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "SELECT "
    strSQLFormaPago = strSQLFormaPago & "* "
    strSQLFormaPago = strSQLFormaPago & "FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "WHERE "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG = '" & strCodigo & "'"
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    rstFormaPago.Open strSQLFormaPago, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFormaPago.EOF Then
        strCodigo = Trim(rstFormaPago!f2forpag & "")
        strCodigoS10 = Trim(rstFormaPago!F2CODS10 & "")
        strDescripcion = Trim(rstFormaPago!F2DESPAG & "")
        lngDias = Val(rstFormaPago!F2DIAS & "")
        strTipo = Trim(rstFormaPago!F2TIPO & "")
        bolParaLetra = IIf(Trim(rstFormaPago!F2LETRA & "") = "S", True, False)
        
        obtenerFormaPago = True
    Else
        obtenerFormaPago = False
    End If
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    Set rstFormaPago = Nothing
    
    Exit Function
errObtenerFormaPago:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: ObtenerFormaPago"
    End Select
    
    obtenerFormaPago = False
    
    Err.Clear
End Function

Public Sub obtenerConfigFormaPago()
    On Error GoTo errObtenerConfigFormaPago
    
    Set rstFormaPago = New ADODB.Recordset
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "SELECT "
    strSQLFormaPago = strSQLFormaPago & "* "
    strSQLFormaPago = strSQLFormaPago & "FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "WHERE "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG = '" & strCodigo & "'"
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    rstFormaPago.Open strSQLFormaPago, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFormaPago.EOF Then
        strCodigo = Trim(rstFormaPago!f2forpag & "")
        strCodigoS10 = Trim(rstFormaPago!F2CODS10 & "")
        strDescripcion = Trim(rstFormaPago!F2DESPAG & "")
        lngDias = Val(rstFormaPago!F2DIAS & "")
        strTipo = Trim(rstFormaPago!F2TIPO & "")
        bolParaLetra = IIf(Trim(rstFormaPago!F2LETRA & "") = "S", True, False)
    End If
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    Set rstFormaPago = Nothing
    
    Exit Sub
errObtenerConfigFormaPago:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: ObtenerConfigFormaPago"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstFormaPago = New ADODB.Recordset
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "SELECT "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG.* "
    strSQLFormaPago = strSQLFormaPago & "FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "WHERE "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG = '" & strCodigo & "'"
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    rstFormaPago.Open strSQLFormaPago, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstFormaPago.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    Set rstFormaPago = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function generarCodigo() As String
    On Error GoTo errGenerarCodigo
    
    Dim rsFormaPago As New ADODB.Recordset
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "SELECT "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "ORDER BY "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG DESC"
    
    If rsFormaPago.State = 1 Then rsFormaPago.Close
    
    rsFormaPago.Open strSQLFormaPago, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rsFormaPago.EOF Then
        generarCodigo = Format(Val(strCodigo) + 1, "000")
    Else
        generarCodigo = "001"
    End If
    
    If rsFormaPago.State = 1 Then rsFormaPago.Close
    
    Set rsFormaPago = Nothing
    
    Exit Function
errGenerarCodigo:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "TipoCadena: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: GenerarCodigo"
    End Select
        
    generarCodigo = vbNullString
    
    Err.Clear
End Function

Public Function guardarFormaPago() As Boolean
    On Error GoTo errGuardarFormaPago
    
    strSQLFormaPago = vbNullString
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigo
        End If
        
        strSQLFormaPago = vbNullString
        strSQLFormaPago = strSQLFormaPago & "INSERT INTO EF2FORPAG("
        strSQLFormaPago = strSQLFormaPago & "F2FORPAG, F2CODS10, F2DESPAG, "
        strSQLFormaPago = strSQLFormaPago & "F2DIAS, F2TIPO, F2LETRA"
        strSQLFormaPago = strSQLFormaPago & ") "
        strSQLFormaPago = strSQLFormaPago & "VALUES("
        strSQLFormaPago = strSQLFormaPago & "'" & strCodigo & "', "
        strSQLFormaPago = strSQLFormaPago & "'" & strCodigoS10 & "', "
        strSQLFormaPago = strSQLFormaPago & "'" & strDescripcion & "', "
        strSQLFormaPago = strSQLFormaPago & lngDias & ", "
        strSQLFormaPago = strSQLFormaPago & "'" & strTipo & "', "
        strSQLFormaPago = strSQLFormaPago & IIf(bolParaLetra, "S", "N")
        strSQLFormaPago = strSQLFormaPago & ")"
    Else
        strSQLFormaPago = vbNullString
        strSQLFormaPago = strSQLFormaPago & "UPDATE "
        strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
        
        strSQLFormaPago = strSQLFormaPago & "SET "
        strSQLFormaPago = strSQLFormaPago & "F2CODS10 = '" & strCodigoS10 & "', "
        strSQLFormaPago = strSQLFormaPago & "F2DESPAG = '" & strDescripcion & "', "
        strSQLFormaPago = strSQLFormaPago & "F2DIAS = " & lngDias & ", "
        strSQLFormaPago = strSQLFormaPago & "F2TIPO = '" & strTipo & "', "
        strSQLFormaPago = strSQLFormaPago & "F2LETRA = " & IIf(bolParaLetra, "S", "N") & " "
        
        strSQLFormaPago = strSQLFormaPago & "WHERE "
        strSQLFormaPago = strSQLFormaPago & "F2FORPAG = '" & strCodigo & "'"
    End If
    
    cnn_dbbancos.Execute strSQLFormaPago
    
    guardarFormaPago = True
    
    Exit Function
errGuardarFormaPago:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: GuardarFormaPago"
    End Select
    
    guardarFormaPago = False
    
    Err.Clear
End Function

Public Function eliminarFormaPago() As Boolean
    On Error GoTo errEliminarFormaPago
    
    If Val(validarUsoRegistro(cnn_dbbancos, "F2FORPAG", "EF2CLIENTES", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarFormaPago = False
        
        Exit Function
    End If
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "DELETE FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "WHERE "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG = '" & strCodigo & "'"
    
    cnn_dbbancos.Execute strSQLFormaPago
    
    eliminarFormaPago = True
    
    Exit Function
errEliminarFormaPago:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: EliminarFormaPago"
    End Select
    
    'cnn_dbbancos.RollbackTrans
    
    eliminarFormaPago = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de FormaPagos en Grilla (QuamtumGrid)
Public Sub vistaFormaPago(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaFormaPago
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "SELECT "
    strSQLFormaPago = strSQLFormaPago & "FP.F2FORPAG, "
    strSQLFormaPago = strSQLFormaPago & "FP.F2DESPAG, "
    strSQLFormaPago = strSQLFormaPago & "FP.F2DIAS, "
    strSQLFormaPago = strSQLFormaPago & "FP.F2TIPO "
    
    strSQLFormaPago = strSQLFormaPago & "FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG AS FP "
    
    strSQLFormaPago = strSQLFormaPago & "ORDER BY "
    strSQLFormaPago = strSQLFormaPago & "FP.F2FORPAG"
    
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
                .FieldName = "F2FORPAG"
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
                .FieldName = "F2DESPAG"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
                .Visible = True
            End With
            
            'Columna Dias
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Días"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2DIAS"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDias"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
                .Visible = True
            End With
            
            'Columna Tipo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F2TIPO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnn_dbbancos.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLFormaPago
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F2FORPAG"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLFormaPago = vbNullString
    
    Exit Sub
errVistaFormaPago:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: VistaFormaPago"
    End Select
    
    Err.Clear
End Sub

Public Sub listarFormaPago(ByVal comboList As Object)
    On Error GoTo errListarFormaPago
    
    Set rstFormaPago = New ADODB.Recordset
    
    strSQLFormaPago = vbNullString
    strSQLFormaPago = strSQLFormaPago & "SELECT "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG, "
    strSQLFormaPago = strSQLFormaPago & "F2DESPAG "
    strSQLFormaPago = strSQLFormaPago & "FROM "
    strSQLFormaPago = strSQLFormaPago & "EF2FORPAG "
    strSQLFormaPago = strSQLFormaPago & "ORDER BY "
    strSQLFormaPago = strSQLFormaPago & "F2FORPAG"
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    rstFormaPago.Open strSQLFormaPago, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstFormaPago.EOF Then
        
        rstFormaPago.MoveFirst
        
        Do While Not rstFormaPago.EOF
            comboList.AddItem Trim(rstFormaPago!F2DESPAG & "") & Space(100) & Trim(rstFormaPago!f2forpag & "")
            
            rstFormaPago.MoveNext
        Loop
    End If
    
    If rstFormaPago.State = 1 Then rstFormaPago.Close
    
    strSQLFormaPago = vbNullString
    
    Set rstFormaPago = Nothing
    
    Exit Sub
errListarFormaPago:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsFormaPagoEmpresarial: ListarFormaPago"
    
    Err.Clear
End Sub

