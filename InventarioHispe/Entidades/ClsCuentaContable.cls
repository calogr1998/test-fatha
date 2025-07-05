VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCuentaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String

Private strDescripcion              As String

Private dblGradoContable            As Double
Private bolTieneCentroCosto         As Boolean
Private strTipoAnexo                As String
Private strCodMoneda                As String

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter           As String

Private rstCuentaContable                As ADODB.Recordset
Private strSQLCuentaContable             As String

'Propiedad Codigo de Cuenta Contable
Public Property Let Codigo(ByVal Value As String)
    strCodigo = Value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Descripcion de CuentaContable
Public Property Let Descripcion(ByVal Value As String)
    strDescripcion = Value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

'Propiedad Grado Contable
Public Property Let GradoContable(ByVal Value As Double)
    dblGradoContable = Value
End Property

Public Property Get GradoContable() As Double
    GradoContable = dblGradoContable
End Property

'Propiedad Tiene Centro de Costo
Public Property Let TieneCentroCosto(ByVal Value As Boolean)
    bolTieneCentroCosto = Value
End Property

Public Property Get TieneCentroCosto() As Boolean
    TieneCentroCosto = bolTieneCentroCosto
End Property

'Propiedad Tipo de Anexo (Auxiliar)
Public Property Let TipoAnexo(ByVal Value As String)
    strTipoAnexo = Value
End Property

Public Property Get TipoAnexo() As String
    TipoAnexo = strTipoAnexo
End Property

'Propiedad Codigo de Moneda
Public Property Let CodMoneda(ByVal Value As String)
    strCodMoneda = Value
End Property

Public Property Get CodMoneda() As String
    CodMoneda = strCodMoneda
End Property



'Propiedad Fecha de Registro de CuentaContable
Public Property Let FechaReg(ByVal Value As String)
    strFecReg = Value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de CuentaContable
Public Property Let UsuarioReg(ByVal Value As String)
    strUsuReg = Value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de CuentaContable
Public Property Let FechaMod(ByVal Value As String)
    strFecMod = Value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de CuentaContable
Public Property Let UsuarioMod(ByVal Value As String)
    strUsuMod = Value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property


'Propiedad SQLSelectAlter
Public Property Let SQLSelectAlter(ByVal Value As String)
    strSQLSelectAlter = Value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    
    strDescripcion = vbNullString
    
    dblGradoContable = 0
    bolTieneCentroCosto = False
    strTipoAnexo = vbNullString
    strCodMoneda = vbNullString
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    
    strSQLSelectAlter = vbNullString
End Sub


Public Function obtenerCuentaContable() As Boolean
    On Error GoTo errObtenerCuentaContable
    
    Set rstCuentaContable = New ADODB.Recordset
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "SELECT "
    strSQLCuentaContable = strSQLCuentaContable & "* "
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA "
    strSQLCuentaContable = strSQLCuentaContable & "WHERE "
    strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA = '" & strCodigo & "'"
    
    If rstCuentaContable.State = 1 Then rstCuentaContable.Close
    
    rstCuentaContable.Open strSQLCuentaContable, cnn_ContawinTabla, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCuentaContable.EOF Then
        strCodigo = Trim(rstCuentaContable!F5CODCTA & "")
        strDescripcion = Trim(rstCuentaContable!F5NOMCTA & "")
        
        dblGradoContable = Trim(rstCuentaContable!F5GRDCTA & "")
        bolTieneCentroCosto = CBool(rstCuentaContable!F5CC)
        strTipoAnexo = Trim(rstCuentaContable!F5PVANEXOCONCAR & "")
        strCodMoneda = Trim(rstCuentaContable!F5MONEDA & "")
        
        obtenerCuentaContable = True
    Else
        obtenerCuentaContable = False
    End If
    
    rstCuentaContable.Close
    
    Set rstCuentaContable = Nothing
    
    Exit Function
errObtenerCuentaContable:
    obtenerCuentaContable = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: ObtenerCuentaContable"
    
    Err.Clear
End Function

Public Sub obtenerConfigCuentaContable()
    On Error GoTo errObtenerConfigCuentaContable
    Set rstCuentaContable = New ADODB.Recordset
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "SELECT "
    strSQLCuentaContable = strSQLCuentaContable & "* "
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA "
    strSQLCuentaContable = strSQLCuentaContable & "WHERE "
    strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA = '" & strCodigo & "'"
    
    If rstCuentaContable.State = 1 Then rstCuentaContable.Close
    
    rstCuentaContable.Open strSQLCuentaContable, cnn_ContawinTabla, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCuentaContable.EOF Then
        strCodigo = Trim(rstCuentaContable!F5CODCTA & "")
        strDescripcion = Trim(rstCuentaContable!F5NOMCTA & "")
        
        dblGradoContable = Trim(rstCuentaContable!F5GRDCTA & "")
        bolTieneCentroCosto = CBool(rstCuentaContable!F5CC)
        strTipoAnexo = Trim(rstCuentaContable!F5PVANEXOCONCAR & "")
        strCodMoneda = Trim(rstCuentaContable!F5MONEDA & "")
    End If
    
    rstCuentaContable.Close
    
    Set rstCuentaContable = Nothing
    
    Exit Sub
    Resume
errObtenerConfigCuentaContable:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: ObtenerConfigCuentaContable"
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstCuentaContable = New ADODB.Recordset
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "SELECT "
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA.* "
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA "
    strSQLCuentaContable = strSQLCuentaContable & "WHERE "
    strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA = '" & strCodigo & "'"
    
    If rstCuentaContable.State = 1 Then rstCuentaContable.Close
    
    rstCuentaContable.Open strSQLCuentaContable, cnn_ContawinTabla, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCuentaContable.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstCuentaContable.Close
    
    Set rstCuentaContable = Nothing
    
    strSQLCuentaContable = vbNullString
    
    Exit Function
errVerificarExistencia:
    verificarExistencia = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: VerificarExistencia"
    
    Err.Clear
End Function

Public Function generarCodigoCuentaContable() As String
    On Error GoTo errGenerarCodigoCuentaContable
    
    Dim rsCuentaContable As New ADODB.Recordset
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "SELECT "
    strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA "
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA "
    strSQLCuentaContable = strSQLCuentaContable & "ORDER BY "
    strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA"
    
    If rsCuentaContable.State = 1 Then rsCuentaContable.Close
    
    rsCuentaContable.Open strSQLCuentaContable, cnn_ContawinTabla, adOpenForwardOnly, adLockReadOnly
    
    If Not rsCuentaContable.EOF Then
        rsCuentaContable.MoveFirst
        
        Do While Not rsCuentaContable.EOF
            strCodigo = Format(Val(Trim(rsCuentaContable!F5CODCTA & "")) + 1, "00")
            
            If Not Me.verificarExistencia Then
                Exit Do
            End If
            
            rsCuentaContable.MoveNext
        Loop
        
        generarCodigoCuentaContable = strCodigo
    Else
        generarCodigoCuentaContable = "01"
    End If
    
    rsCuentaContable.Close
    
    Set rsCuentaContable = Nothing
    
    Exit Function
errGenerarCodigoCuentaContable:
    generarCodigoCuentaContable = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: GenerarCodigoCuentaContable"
    
    Err.Clear
End Function

Public Function guardarCuentaContable() As Boolean
    On Error GoTo errGuardarCuentaContable
    
    strSQLCuentaContable = vbNullString
    
    If Not verificarExistencia Then
'        If strCodigo = vbNullString Then
'            strCodigo = generarCodigoCuentaContable
'        End If
        
        strSQLCuentaContable = strSQLCuentaContable & "INSERT INTO CF5PLA("
        strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA, F5NOMCTA, F5GRDCTA, F5CC, F5PVANEXOCONCAR, F5MONEDA"
        strSQLCuentaContable = strSQLCuentaContable & ") "
        strSQLCuentaContable = strSQLCuentaContable & "VALUES("
        strSQLCuentaContable = strSQLCuentaContable & "'" & strCodigo & "', "
        strSQLCuentaContable = strSQLCuentaContable & "'" & strDescripcion & "', "
        strSQLCuentaContable = strSQLCuentaContable & dblGradoContable & ", "
        strSQLCuentaContable = strSQLCuentaContable & IIf(bolTieneCentroCosto, "TRUE", "FALSE") & ", "
        strSQLCuentaContable = strSQLCuentaContable & "'" & strTipoAnexo & "', "
        strSQLCuentaContable = strSQLCuentaContable & "'" & strCodMoneda & "'"
        strSQLCuentaContable = strSQLCuentaContable & ")"
    Else
        strSQLCuentaContable = strSQLCuentaContable & "UPDATE CF5PLA "
        strSQLCuentaContable = strSQLCuentaContable & "SET "
        strSQLCuentaContable = strSQLCuentaContable & "F5NOMCTA = '" & strDescripcion & "', "
        strSQLCuentaContable = strSQLCuentaContable & "F5GRDCTA = " & dblGradoContable & ","
        strSQLCuentaContable = strSQLCuentaContable & "F5CC = " & IIf(bolTieneCentroCosto, "TRUE", "FALSE") & ", "
        strSQLCuentaContable = strSQLCuentaContable & "F5PVANEXOCONCAR = '" & strTipoAnexo & "', "
        strSQLCuentaContable = strSQLCuentaContable & "F5MONEDA = '" & strCodMoneda & "' "
        strSQLCuentaContable = strSQLCuentaContable & "WHERE "
        strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA = '" & strCodigo & "'"
    End If
    
    cnn_ContawinTabla.Execute strSQLCuentaContable
    
    strSQLSelectAlter = strSQLCuentaContable
    strSQLCuentaContable = vbNullString
    
    guardarCuentaContable = True
    
    Exit Function
errGuardarCuentaContable:
    guardarCuentaContable = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: GuardarCuentaContable"
    
    Err.Clear
End Function

Public Function eliminarCuentaContable() As Boolean
    On Error GoTo errEliminarCuentaContable
    
    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F4CTACONT", "REGISDOC", strCodigo, "T")) <> 0 Or _
        Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F3CTACON", "REGISMOV", strCodigo, "T")) <> 0 Or _
        Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "F4CTACONT", "PAG_DCTO", strCodigo, "T")) <> 0 Then
        
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarCuentaContable = False
        
        Exit Function
    End If
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "DELETE "
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA "
    strSQLCuentaContable = strSQLCuentaContable & "WHERE "
    strSQLCuentaContable = strSQLCuentaContable & "F5CODCTA = '" & strCodigo & "'"
    
    cnn_ContawinTabla.Execute strSQLCuentaContable
    
    strSQLSelectAlter = strSQLCuentaContable
    strSQLCuentaContable = vbNullString
    
    eliminarCuentaContable = True
    
    Exit Function
errEliminarCuentaContable:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: EliminarCuentaContable"
    
    eliminarCuentaContable = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de CuentaContables en Grilla (QuamtumGrid)
Public Sub vistaCuentaContable(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaCuentaContable
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "SELECT "
    strSQLCuentaContable = strSQLCuentaContable & "PL.F5CODCTA, "
    strSQLCuentaContable = strSQLCuentaContable & "PL.F5NOMCTA, "
    strSQLCuentaContable = strSQLCuentaContable & "PL.F5GRDCTA, "
    'strSQLCuentaContable = strSQLCuentaContable & "IIF(PL.F5CC, 'S', 'N') AS TIENECC, "
    strSQLCuentaContable = strSQLCuentaContable & "PL.F5CC, "
    strSQLCuentaContable = strSQLCuentaContable & "PL.F5PVANEXOCONCAR, "
    strSQLCuentaContable = strSQLCuentaContable & "IIF(PL.F5MONEDA = 'S', 'MN', IIF(PL.F5MONEDA = 'D', 'US', '')) AS MONEDA "
    
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    
    strSQLCuentaContable = strSQLCuentaContable & "CF5PLA AS PL "
    
    strSQLCuentaContable = strSQLCuentaContable & "ORDER BY "
    strSQLCuentaContable = strSQLCuentaContable & "PL.F5CODCTA"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de CuentaContable
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Código"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5CODCTA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Descripcion
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripción"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5NOMCTA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 150
            End With
            
            'Columna Grado de Cuenta Contable
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Grado"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5GRDCTA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColGrado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Tipo de Anexo (Auxiliar)
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Tipo Aux."
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F5PVANEXOCONCAR"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipoAuxiliar"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 50
            End With
            
            'Columna Cuenta Contable con Centro Costo
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Tiene C.C."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "F5CC"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTieneCC"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            'Columna Moneda de Cuenta
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Moneda Ref."
                .Color = &HC0FFFF
                .DisableEditor = True
                .FieldName = "MONEDA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColMonedaRef"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 40
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexContawinTabla
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLCuentaContable
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F5CODCTA"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLCuentaContable = vbNullString
    
    Exit Sub
errVistaCuentaContable:
    Select Case Err.Number
        Case 3704, 3709
            cnn_ContawinTabla.Open StrConexContawinTabla
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCuentaContable: VistaCuentaContable"
    End Select
    
    Err.Clear
End Sub

'Vista de Detalle de CuentaContables Presupuestales en Grilla (QuamtumGrid)
Public Sub vistaCuentaContableDetalle(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaCuentaContableDetalle
    
    strSQLCuentaContable = vbNullString
    strSQLCuentaContable = strSQLCuentaContable & "SELECT "
    strSQLCuentaContable = strSQLCuentaContable & "CD.CODPARTIDA, "
    strSQLCuentaContable = strSQLCuentaContable & "CD.CODINSUMO, "
    strSQLCuentaContable = strSQLCuentaContable & "I.F5NOMPRO, "
    strSQLCuentaContable = strSQLCuentaContable & "M.F7NOMMED, "
    strSQLCuentaContable = strSQLCuentaContable & "VAL(CD.CANTIDAD & '') AS CANT "
    
    strSQLCuentaContable = strSQLCuentaContable & "FROM "
    
    strSQLCuentaContable = strSQLCuentaContable & "(CF5PLADETALLE AS CD "
    strSQLCuentaContable = strSQLCuentaContable & "LEFT JOIN IF5PLA AS I ON I.F5CODPRO = CD.CODINSUMO) "
    strSQLCuentaContable = strSQLCuentaContable & "LEFT JOIN EF7MEDIDAS AS M ON M.F7CODMED = I.F7CODMED "
    
    strSQLCuentaContable = strSQLCuentaContable & "WHERE "
    strSQLCuentaContable = strSQLCuentaContable & "CD.F5CODCTA = '" & strCodigo & "'"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo de Partida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Partida"
                .Color = &H80000000
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODPARTIDA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodPartida"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Codigo de Insumo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &H80000000
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "CODINSUMO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodInsumo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Descripcion de Insumo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .BandIndex = 0
                .Caption = "Descripción de Recurso"
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "F5NOMPRO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 160
            End With
            
            'Columna Unidad de Medida
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .BandIndex = 0
                .Caption = "Unidad Med."
                .Color = &HFFFFC0
                .DisableEditor = True
                .FieldName = "F7NOMMED"
                .HeaderAlignment = taCenter
                .ObjectName = "ColUnidadMedida"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            'Columna Cantidad
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taRightJustify
                .Caption = "Cantidad"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DecimalPlaces = 2
                .DisableEditor = True
                .FieldName = "CANT"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCantidad"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = StrConexContawinTabla
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLCuentaContable
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODPARTIDA"
        End With
    End If
    
    strSQLCuentaContable = vbNullString
    
    Exit Sub
errVistaCuentaContableDetalle:
    Select Case Err.Number
        Case 3704, 3709
            cnn_ContawinTabla.Open StrConexContawinTabla
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ClsCuentaContable: VistaCuentaContableDetalle"
    End Select
    
    Err.Clear
End Sub

Public Sub listarAnnoCuentaContable(ByVal combo As Object)
    On Error GoTo errListarAnnoCuentaContable
    
    Set rstCuentaContable = New ADODB.Recordset
    
    strSQLCuentaContable = "SELECT YEAR(FECHA) AS ANNO " & _
                    "FROM CF5PLA " & _
                    "GROUP BY YEAR(FECHA)"
    
    If rstCuentaContable.State = 1 Then rstCuentaContable.Close
    
    rstCuentaContable.Open strSQLCuentaContable, cnn_ContawinTabla, adOpenForwardOnly, adLockReadOnly
    
    combo.Clear
    combo.AddItem "(*) - Todos"
    
    If Not rstCuentaContable.EOF Then
        
        rstCuentaContable.MoveFirst
        
        Do While Not rstCuentaContable.EOF
            combo.AddItem Trim(rstCuentaContable!Anno & "")
            
            rstCuentaContable.MoveNext
        Loop
    End If
    
    rstCuentaContable.Close
    
    strSQLCuentaContable = vbNullString
    
    Set rstCuentaContable = Nothing
    
    Exit Sub
errListarAnnoCuentaContable:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: ListarAnnoCuentaContable"
    
    Err.Clear
End Sub

Public Sub listarGradoContable(ByVal comboList As Object, ByVal strIncluirGradoContableCorPorAorAll As String)
    On Error GoTo errListarGradoContable
    
    Set rstCuentaContable = New ADODB.Recordset
    
    strSQLCuentaContable = "SELECT F5CODCTA, F5PVANEXOCONCAR " & _
                        "FROM CF5PLA " & _
                        "WHERE F5GRDCTA IN (" & strIncluirGradoContableCorPorAorAll & ") " & _
                        "ORDER BY F5PVANEXOCONCAR"
    
    If rstCuentaContable.State = 1 Then rstCuentaContable.Close
    
    rstCuentaContable.Open strSQLCuentaContable, cnn_ContawinTabla, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstCuentaContable.EOF Then
        
        rstCuentaContable.MoveFirst
        
        Do While Not rstCuentaContable.EOF
            comboList.AddItem Trim(rstCuentaContable!F5PVANEXOCONCAR & "") & Space(100) & Trim(rstCuentaContable!F5CODCTA & "")
            
            rstCuentaContable.MoveNext
        Loop
    End If
    
    If rstCuentaContable.State = 1 Then rstCuentaContable.Close
    
    strSQLCuentaContable = vbNullString
    
    Set rstCuentaContable = Nothing
    
    Exit Sub
errListarGradoContable:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripcion: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsCuentaContable: ListarGradoContable"
    
    Err.Clear
End Sub
