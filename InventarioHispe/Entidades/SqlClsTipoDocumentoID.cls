VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsTipoDocumentoID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private strDescripcion              As String
Private strAbreviatura              As String
Private lngLongitud                 As Long
Private strTipoCadena               As String
Private bolTieneLargoFijo           As Boolean
Private strModulo11                 As String
Private strOrigen                   As String
Private strPersona                  As String

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

Private strSQLSelectAlter           As String

Private rstTipoDocumento            As ADODB.Recordset
Private strSQLTipoDocumento         As String


'Propiedad Codigo de Tipo de Documento
Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

'Propiedad Descripcion de Tipo de Documento
Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

'Propiedad Abreviatura de Tipo de Documento
Public Property Let Abreviatura(ByVal value As String)
    strAbreviatura = value
End Property

Public Property Get Abreviatura() As String
    Abreviatura = strAbreviatura
End Property

'Propiedad Longitud de Tipo de Documento de Identidad
Public Property Let Longitud(ByVal value As String)
    lngLongitud = value
End Property

Public Property Get Longitud() As String
    Longitud = lngLongitud
End Property

'Propiedad Tipo de Cadena de TipoDocumento
Public Property Let TipoCadena(ByVal value As String)
    strTipoCadena = value
End Property

Public Property Get TipoCadena() As String
    TipoCadena = strTipoCadena
End Property

'Propiedad Tiene Largo Fijo de TipoDocumento
Public Property Let TieneLargoFijo(ByVal value As String)
    bolTieneLargoFijo = value
End Property

Public Property Get TieneLargoFijo() As String
    TieneLargoFijo = bolTieneLargoFijo
End Property

'Propiedad Modulo 11
Public Property Let Modulo11(ByVal value As String)
    strModulo11 = value
End Property

Public Property Get Modulo11() As String
    Modulo11 = strModulo11
End Property

'Propiedad Origen
Public Property Let Origen(ByVal value As String)
    strOrigen = value
End Property

Public Property Get Origen() As String
    Origen = strOrigen
End Property

'Propiedad Persona
Public Property Let Persona(ByVal value As String)
    strPersona = value
End Property

Public Property Get Persona() As String
    Persona = strPersona
End Property




'Propiedad Fecha de Registro de TipoDocumento
Public Property Let FechaReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de TipoDocumento
Public Property Let UsuarioReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de TipoDocumento
Public Property Let FechaMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de TipoDocumento
Public Property Let UsuarioMod(ByVal value As String)
    strUsuMod = value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
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
    strAbreviatura = vbNullString
    lngLongitud = 0
    strTipoCadena = vbNullString
    bolTieneLargoFijo = False
    strModulo11 = vbNullString
    strOrigen = vbNullString
    strPersona = vbNullString
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerTipoDocumento() As Boolean
    On Error GoTo errObtenerTipoDocumento
    
    Set rstTipoDocumento = New ADODB.Recordset
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "SELECT "
    strSQLTipoDocumento = strSQLTipoDocumento & "* "
    strSQLTipoDocumento = strSQLTipoDocumento & "FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
    strSQLTipoDocumento = strSQLTipoDocumento & "WHERE "
    strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO = '" & strCodigo & "'"
    
    If rstTipoDocumento.State = 1 Then rstTipoDocumento.Close
    
    rstTipoDocumento.Open strSQLTipoDocumento, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTipoDocumento.EOF Then
        strCodigo = Trim(rstTipoDocumento!Codigo & "")
        strDescripcion = Trim(rstTipoDocumento!Descripcion & "")
        strAbreviatura = Trim(rstTipoDocumento!Abreviatura & "")
        lngLongitud = Val(rstTipoDocumento!Longitud & "")
        strTipoCadena = Trim(rstTipoDocumento!TipoCadena & "")
        bolTieneLargoFijo = CBool(rstTipoDocumento!TieneLargoFijo)
        strModulo11 = Trim(rstTipoDocumento!Modulo11 & "")
        strOrigen = Trim(rstTipoDocumento!Origen & "")
        strPersona = Trim(rstTipoDocumento!Persona & "")
        
        strFecReg = Trim(rstTipoDocumento!FecReg & "")
        strUsuReg = Trim(rstTipoDocumento!UsuReg & "")
        strFecMod = Trim(rstTipoDocumento!FecMod & "")
        strUsuMod = Trim(rstTipoDocumento!UsuMod & "")
        
        obtenerTipoDocumento = True
    Else
        obtenerTipoDocumento = False
    End If
    
    rstTipoDocumento.Close
    
    Set rstTipoDocumento = Nothing
    
    Exit Function
errObtenerTipoDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: ObtenerTipoDocumento"
    End Select
    
    obtenerTipoDocumento = False
    
    Err.Clear
End Function

Public Sub obtenerConfigTipoDocumento()
    On Error GoTo errObtenerConfigTipoDocumento
    Set rstTipoDocumento = New ADODB.Recordset
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "SELECT "
    strSQLTipoDocumento = strSQLTipoDocumento & "* "
    strSQLTipoDocumento = strSQLTipoDocumento & "FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
    strSQLTipoDocumento = strSQLTipoDocumento & "WHERE "
    strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO = '" & strCodigo & "'"
    
    If rstTipoDocumento.State = 1 Then rstTipoDocumento.Close
    
    rstTipoDocumento.Open strSQLTipoDocumento, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTipoDocumento.EOF Then
        strCodigo = Trim(rstTipoDocumento!Codigo & "")
        strDescripcion = Trim(rstTipoDocumento!Descripcion & "")
        strAbreviatura = Trim(rstTipoDocumento!Abreviatura & "")
        lngLongitud = Val(rstTipoDocumento!Longitud & "")
        strTipoCadena = Trim(rstTipoDocumento!TipoCadena & "")
        bolTieneLargoFijo = CBool(rstTipoDocumento!TieneLargoFijo)
        strModulo11 = Trim(rstTipoDocumento!Modulo11 & "")
        strOrigen = Trim(rstTipoDocumento!Origen & "")
        strPersona = Trim(rstTipoDocumento!Persona & "")
        
        strFecReg = Trim(rstTipoDocumento!FecReg & "")
        strUsuReg = Trim(rstTipoDocumento!UsuReg & "")
        strFecMod = Trim(rstTipoDocumento!FecMod & "")
        strUsuMod = Trim(rstTipoDocumento!UsuMod & "")
    End If
    
    rstTipoDocumento.Close
    
    Set rstTipoDocumento = Nothing
    
    Exit Sub
errObtenerConfigTipoDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description, _
                vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: ObtenerConfigTipoDocumento"
    End Select
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstTipoDocumento = New ADODB.Recordset
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "SELECT "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID.* "
    strSQLTipoDocumento = strSQLTipoDocumento & "FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
    strSQLTipoDocumento = strSQLTipoDocumento & "WHERE "
    strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO = '" & strCodigo & "'"
    
    If rstTipoDocumento.State = 1 Then rstTipoDocumento.Close
    
    rstTipoDocumento.Open strSQLTipoDocumento, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTipoDocumento.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstTipoDocumento.Close
    
    Set rstTipoDocumento = Nothing
    
    Exit Function
errVerificarExistencia:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: VerificarExistencia"
    End Select
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function generarCodigo() As String
    On Error GoTo errGenerarCodigo
    
    Dim rsTipoDocumento As New ADODB.Recordset
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "SELECT "
    strSQLTipoDocumento = strSQLTipoDocumento & "F2CODCLI "
    strSQLTipoDocumento = strSQLTipoDocumento & "FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
    strSQLTipoDocumento = strSQLTipoDocumento & "ORDER BY "
    strSQLTipoDocumento = strSQLTipoDocumento & "F2CODCLI DESC"
    
    If rsTipoDocumento.State = 1 Then rsTipoDocumento.Close
    
    rsTipoDocumento.Open strSQLTipoDocumento, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rsTipoDocumento.EOF Then
        generarCodigo = Str(Val(strCodigo) + 1)
    Else
        generarCodigo = "1"
    End If
    
    rsTipoDocumento.Close
    
    Set rsTipoDocumento = Nothing
    
    Exit Function
errGenerarCodigo:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "TipoCadena: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: GenerarCodigo"
    End Select
        
    generarCodigo = vbNullString
    
    Err.Clear
End Function

Public Function guardarTipoDocumento() As Boolean
    On Error GoTo errGuardarTipoDocumento
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigo
        End If
        
        strSQLTipoDocumento = vbNullString
        strSQLTipoDocumento = strSQLTipoDocumento & "INSERT INTO MAESTROS.EF2TIPODOCID("
        strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO, DESCRIPCION, ABREVIATURA, LONGITUD, TIPOCADENA, "
        strSQLTipoDocumento = strSQLTipoDocumento & "TIENELARGOFIJO, MODULO11, ORIGEN, PERSONA, FECREG, USUREG) "
        
        strSQLTipoDocumento = strSQLTipoDocumento & "VALUES("
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strCodigo & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strDescripcion & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strAbreviatura & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & lngLongitud & ", "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strTipoCadena & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & IIf(bolTieneLargoFijo, "1", "0") & ", "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strModulo11 & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strOrigen & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strPersona & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strFecReg & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "'" & strUsuReg & "'"
        strSQLTipoDocumento = strSQLTipoDocumento & ")"
    Else
        strSQLTipoDocumento = vbNullString
        strSQLTipoDocumento = strSQLTipoDocumento & "UPDATE "
        strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
        
        strSQLTipoDocumento = strSQLTipoDocumento & "SET "
        strSQLTipoDocumento = strSQLTipoDocumento & "DESCRIPCION = '" & strDescripcion & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "ABREVIATURA = '" & strAbreviatura & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "LONGITUD = " & lngLongitud & ", "
        strSQLTipoDocumento = strSQLTipoDocumento & "TIPOCADENA = '" & strTipoCadena & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "TIENELARGOFIJO = " & IIf(bolTieneLargoFijo, "1", "0") & ", "
        strSQLTipoDocumento = strSQLTipoDocumento & "MODULO11 = '" & strModulo11 & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "ORIGEN = '" & strOrigen & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "PERSONA = '" & strPersona & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "FECMOD = '" & strFecMod & "', "
        strSQLTipoDocumento = strSQLTipoDocumento & "USUMOD = " & strUsuMod & "' "
        
        strSQLTipoDocumento = strSQLTipoDocumento & "WHERE "
        strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO = '" & strCodigo & "'"
    End If
     
    cnBdCPlus.Execute strSQLTipoDocumento
    
    strSQLSelectAlter = strSQLTipoDocumento
    strSQLTipoDocumento = vbNullString
    
    guardarTipoDocumento = True
    
    Exit Function
errGuardarTipoDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: GuardarTipoDocumento"
    End Select
    
    guardarTipoDocumento = False
    
    Err.Clear
End Function

Public Function eliminarTipoDocumento() As Boolean
    On Error GoTo errEliminarTipoDocumento
    
    If Val(validarUsoRegistro(cnBdCPlus, "F2DOCCLI", "MAESTROS.EF2CLIENTES", strCodigo, "T")) > 0 Then
        
        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
        
        eliminarTipoDocumento = False
        
        Exit Function
    End If
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "DELETE FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
    strSQLTipoDocumento = strSQLTipoDocumento & "WHERE "
    strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLTipoDocumento
    
    eliminarTipoDocumento = True
    
    Exit Function
errEliminarTipoDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: EliminarTipoDocumento"
    End Select
    
    eliminarTipoDocumento = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de TipoDocumentos en Grilla (QuamtumGrid)
Public Sub vistaTipoDocumento(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaTipoDocumento
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "SELECT "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.CODIGO, "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.DESCRIPCION, "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.ABREVIATURA, "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.LONGITUD, "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.TIPOCADENA, "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.TIENELARGOFIJO, "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.MODULO11 "
    
    strSQLTipoDocumento = strSQLTipoDocumento & "FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID AS TD "
    
    strSQLTipoDocumento = strSQLTipoDocumento & "ORDER BY "
    strSQLTipoDocumento = strSQLTipoDocumento & "TD.CODIGO"
    
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
                .FieldName = "CODIGO"
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
                .FieldName = "DESCRIPCION"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
                .Visible = True
            End With
            
            'Columna Abreviatura
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Abreviatura"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "ABREVIATURA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColAbreviatura"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = True
            End With
            
            'Columna Longitud
            Set gColumn = .Columns.Add(gedSpinEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Longitud"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "LONGITUD"
                .HeaderAlignment = taCenter
                .ObjectName = "ColLongitud"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            'Columna Tipo Cadena
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Tipo de Cadena"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "TIPOCADENA"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTipoCadena"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Largo Fijo
            Set gColumn = .Columns.Add(gedCheckEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Largo Fijo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "TIENELARGOFIJO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColTieneLargoFijo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 70
                .Visible = True
            End With
            
            'Columna Modulo 11
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Modulo 11"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "MODULO11"
                .HeaderAlignment = taCenter
                .ObjectName = "ColModulo11"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = cnBdCPlus.ConnectionString
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLTipoDocumento
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "CODIGO"
            
            '.Columns.ColumnByFieldName("CENTRO").GroupIndex = 0
            '.Columns.ColumnByFieldName("CENTRO").Font.Bold = True
            '.Columns.ColumnByFieldName("CENTRO").FontColor = RGB(255, 255, 255)
            
            '.M.FullCollapse
        End With
    End If
    
    strSQLTipoDocumento = vbNullString
    
    Exit Sub
errVistaTipoDocumento:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: VistaTipoDocumento"
    End Select
    
    Err.Clear
End Sub

Public Sub listarTipoDocumento(ByVal comboList As Object)
    On Error GoTo errListarTipoDocumento
    
    Set rstTipoDocumento = New ADODB.Recordset
    
    strSQLTipoDocumento = vbNullString
    strSQLTipoDocumento = strSQLTipoDocumento & "SELECT "
    strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO, ABREVIATURA "
    strSQLTipoDocumento = strSQLTipoDocumento & "FROM "
    strSQLTipoDocumento = strSQLTipoDocumento & "MAESTROS.EF2TIPODOCID "
    strSQLTipoDocumento = strSQLTipoDocumento & "WHERE "
    strSQLTipoDocumento = strSQLTipoDocumento & "TRIM(CODIGO & '') <> '' "
    
        If strOrigen <> vbNullString Then
            strSQLTipoDocumento = strSQLTipoDocumento & "AND ORIGEN IN ('A', '" & strOrigen & "') "
        End If
        
        If strPersona <> vbNullString Then
            strSQLTipoDocumento = strSQLTipoDocumento & "AND PERSONA IN ('A', '" & strPersona & "') "
        End If
        
    strSQLTipoDocumento = strSQLTipoDocumento & "ORDER BY "
    strSQLTipoDocumento = strSQLTipoDocumento & "CODIGO"
    
    If rstTipoDocumento.State = 1 Then rstTipoDocumento.Close
    
    rstTipoDocumento.Open strSQLTipoDocumento, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    comboList.Clear
    
    If Not rstTipoDocumento.EOF Then
        
        rstTipoDocumento.MoveFirst
        
        Do While Not rstTipoDocumento.EOF
            comboList.AddItem Trim(rstTipoDocumento!Abreviatura & "") & Space(100) & Trim(rstTipoDocumento!Codigo & "")
            
            rstTipoDocumento.MoveNext
        Loop
    End If
    
    If rstTipoDocumento.State = 1 Then rstTipoDocumento.Close
    
    strSQLTipoDocumento = vbNullString
    
    Set rstTipoDocumento = Nothing
    
    Exit Sub
errListarTipoDocumento:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsTipoDocumentoID: ListarTipoDocumento"
    
    Err.Clear
End Sub

