VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlClsCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo                   As String
Private intCodigoNivel              As Integer
Private strCodigoExterno            As String
Private strCodigoConcar             As String
Private strDescripcion              As String
Private strAbreviatura              As String
Private strCodCliente               As String
Private dblUtilidad                 As Double

Private bolEstado                   As Boolean

Private strFecReg                   As String
Private strUsuReg                   As String
Private strFecMod                   As String
Private strUsuMod                   As String

'Atributos adicionales
Private strCodigoOrigen             As String

Private strSQLSelectAlter   As String

Private rstCentroCosto             As ADODB.Recordset
Private strSQLCentroCosto          As String



Public Property Let Codigo(ByVal value As String)
    strCodigo = value
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property
'intCodigoNivel
Public Property Let CodigoNivel(ByVal value As Integer)
    intCodigoNivel = value
End Property

Public Property Get CodigoNivel() As Integer
    CodigoNivel = intCodigoNivel
End Property

Public Property Let CodigoExterno(ByVal value As String)
    strCodigoExterno = value
End Property

Public Property Get CodigoExterno() As String
    CodigoExterno = strCodigoExterno
End Property

Public Property Let CodigoConcar(ByVal value As String)
    strCodigoConcar = value
End Property

Public Property Get CodigoConcar() As String
    CodigoConcar = strCodigoConcar
End Property

Public Property Let Descripcion(ByVal value As String)
    strDescripcion = value
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

Public Property Let Abreviatura(ByVal value As String)
    strAbreviatura = value
End Property

Public Property Get Abreviatura() As String
    Abreviatura = strAbreviatura
End Property

Public Property Let CodigoCliente(ByVal value As String)
    strCodCliente = value
End Property

Public Property Get CodigoCliente() As String
    CodigoCliente = strCodCliente
End Property

Public Property Let Utilidad(ByVal value As Double)
    dblUtilidad = value
End Property

Public Property Get Utilidad() As Double
    Utilidad = dblUtilidad
End Property




Public Property Let Estado(ByVal value As Boolean)
    bolEstado = value
End Property

Public Property Get Estado() As Boolean
    Estado = bolEstado
End Property



'Propiedad Fecha de Registro de Color de Bien
Public Property Let FechaReg(ByVal value As String)
    strFecReg = value
End Property

Public Property Get FechaReg() As String
    FechaReg = strFecReg
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioReg(ByVal value As String)
    strUsuReg = value
End Property

Public Property Get UsuarioReg() As String
    UsuarioReg = strUsuReg
End Property

'Propiedad Fecha de Modificacion de Color de Bien
Public Property Let FechaMod(ByVal value As String)
    strFecMod = value
End Property

Public Property Get FechaMod() As String
    FechaMod = strFecMod
End Property

'Propiedad Usuario de Registro de Color de Bien
Public Property Let UsuarioMod(ByVal value As String)
    strUsuMod = value
End Property

Public Property Get UsuarioMod() As String
    UsuarioMod = strUsuMod
End Property




'CodigoOrigen
Public Property Let CodigoOrigen(ByVal value As String)
    strCodigoOrigen = value
End Property

Public Property Get CodigoOrigen() As String
    CodigoOrigen = strCodigoOrigen
End Property



Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodigo = vbNullString
    intCodigoNivel = 0
    strCodigoExterno = vbNullString
    strCodigoConcar = vbNullString
    strDescripcion = vbNullString
    strAbreviatura = vbNullString
    strCodCliente = vbNullString
    dblUtilidad = 0
    
    bolEstado = False
    
    strFecReg = vbNullString
    strUsuReg = vbNullString
    strFecMod = vbNullString
    strUsuMod = vbNullString
    
    strCodigoOrigen = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerCentroCosto() As Boolean
    On Error GoTo errObtenerCentroCosto
    
    Set rstCentroCosto = New ADODB.Recordset
    
    strSQLCentroCosto = vbNullString
    strSQLCentroCosto = strSQLCentroCosto & "SELECT "
    strSQLCentroCosto = strSQLCentroCosto & "* "
    strSQLCentroCosto = strSQLCentroCosto & "FROM "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS "
    strSQLCentroCosto = strSQLCentroCosto & "WHERE "
    strSQLCentroCosto = strSQLCentroCosto & "F3COSTO = '" & strCodigo & "'"
    
    If rstCentroCosto.State = 1 Then rstCentroCosto.Close
    
    rstCentroCosto.Open strSQLCentroCosto, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCentroCosto.EOF Then
        strCodigo = Trim(rstCentroCosto!F3COSTO & "")
        intCodigoNivel = Val(rstCentroCosto!intCodigoNivel & "")
        strCodigoExterno = Trim(rstCentroCosto!CODEXT2 & "")
        strCodigoConcar = Trim(rstCentroCosto!CCONCAR & "")
        strDescripcion = Trim(rstCentroCosto!F3DESCRIP & "")
        strAbreviatura = Trim(rstCentroCosto!F3ABREV & "")
        strCodCliente = Trim(rstCentroCosto!F3CODCLI & "")
        dblUtilidad = Val(rstCentroCosto!Utilidad & "")
        
        bolEstado = IIf(Trim(rstCentroCosto!F3ESTNUL) = "N", True, False)
        
        obtenerCentroCosto = True
    Else
        obtenerCentroCosto = False
    End If
    
    rstCentroCosto.Close
    
    Set rstCentroCosto = Nothing
    
    Exit Function
errObtenerCentroCosto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsCentroCosto: ObtenerCentroCosto"
    
    obtenerCentroCosto = False
    
    Err.Clear
End Function

Public Sub obtenerConfigCentroCosto()
    On Error GoTo errObtenerConfigCentroCosto
    
    Set rstCentroCosto = New ADODB.Recordset
    
    strSQLCentroCosto = vbNullString
    strSQLCentroCosto = strSQLCentroCosto & "SELECT "
    strSQLCentroCosto = strSQLCentroCosto & "* "
    strSQLCentroCosto = strSQLCentroCosto & "FROM "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS "
    strSQLCentroCosto = strSQLCentroCosto & "WHERE "
    strSQLCentroCosto = strSQLCentroCosto & "F3COSTO = '" & strCodigo & "'"
    
    If rstCentroCosto.State = 1 Then rstCentroCosto.Close
    
    rstCentroCosto.Open strSQLCentroCosto, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCentroCosto.EOF Then
        strCodigo = Trim(rstCentroCosto!F3COSTO & "")
        intCodigoNivel = Val(rstCentroCosto!intCodigoNivel & "")
        strCodigoExterno = Trim(rstCentroCosto!CODEXT2 & "")
        strCodigoConcar = Trim(rstCentroCosto!CCONCAR & "")
        strDescripcion = Trim(rstCentroCosto!F3DESCRIP & "")
        strAbreviatura = Trim(rstCentroCosto!F3ABREV & "")
        strCodCliente = Trim(rstCentroCosto!F3CODCLI & "")
        dblUtilidad = Val(rstCentroCosto!Utilidad & "")
        
        bolEstado = IIf(Trim(rstCentroCosto!F3ESTNUL) = "N", True, False)
    End If
    
    rstCentroCosto.Close
    
    Set rstCentroCosto = Nothing
    
    Exit Sub
errObtenerConfigCentroCosto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsCentroCosto: ObtenerCentroCosto"
    
    Err.Clear
End Sub

Public Function generarCodigoCentroCosto() As String
    On Error GoTo errGenerarCodigoCentroCosto
    
    Set rstCentroCosto = New ADODB.Recordset
    
    strSQLCentroCosto = vbNullString
    strSQLCentroCosto = strSQLCentroCosto & "SELECT "
    strSQLCentroCosto = strSQLCentroCosto & "TOP 1 SUBSTRING(F3COSTO, " & Len(strCodigoOrigen) + 1 & ") AS CODIGO "
    strSQLCentroCosto = strSQLCentroCosto & "FROM "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS "
    strSQLCentroCosto = strSQLCentroCosto & "WHERE "
    strSQLCentroCosto = strSQLCentroCosto & "INTCODIGONIVEL = " & intCodigoNivel & " AND "
    strSQLCentroCosto = strSQLCentroCosto & "LEFT(F3COSTO, " & Len(strCodigoOrigen) & ") = '" & strCodigoOrigen & "' "
    strSQLCentroCosto = strSQLCentroCosto & "ORDER BY "
    strSQLCentroCosto = strSQLCentroCosto & "F3COSTO DESC"
    
    If rstCentroCosto.State = 1 Then rstCentroCosto.Close
    
    rstCentroCosto.Open strSQLCentroCosto, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCentroCosto.EOF Then
        generarCodigoCentroCosto = strCodigoOrigen & Format(Val(rstCentroCosto!Codigo & "") + 1, "0000")
    Else
        generarCodigoCentroCosto = strCodigoOrigen & "0001"
    End If
    
    rstCentroCosto.Close
    
    Set rstCentroCosto = Nothing
    
    strSQLCentroCosto = vbNullString
    
    Exit Function
errGenerarCodigoCentroCosto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsCentroCosto: GenerarCodigoCentroCosto"
    
    generarCodigoCentroCosto = vbNullString
    
    Err.Clear
End Function

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstCentroCosto = New ADODB.Recordset
    
    strSQLCentroCosto = vbNullString
    strSQLCentroCosto = strSQLCentroCosto & "SELECT "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS.* "
    strSQLCentroCosto = strSQLCentroCosto & "FROM "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS "
    strSQLCentroCosto = strSQLCentroCosto & "WHERE "
    strSQLCentroCosto = strSQLCentroCosto & "F3COSTO = '" & strCodigo & "'"
    
    If rstCentroCosto.State = 1 Then rstCentroCosto.Close
    
    rstCentroCosto.Open strSQLCentroCosto, cnBdCPlus, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCentroCosto.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstCentroCosto.Close
    
    strSQLCentroCosto = vbNullString
    
    Set rstCentroCosto = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsCentroCosto: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarCentroCosto() As Boolean
    On Error GoTo errGuardarCentroCosto
    
    If Not verificarExistencia Then
        If strCodigo = vbNullString Then
            strCodigo = generarCodigoCentroCosto
        End If
        
        strSQLCentroCosto = vbNullString
        strSQLCentroCosto = strSQLCentroCosto & "INSERT INTO MAESTROS.CENTROS("
        strSQLCentroCosto = strSQLCentroCosto & "F3COSTO, INTCODIGONIVEL, CODEXT2, CCONCAR, F3DESCRIP, "
        strSQLCentroCosto = strSQLCentroCosto & "F3ABREV, F3CODCLI, UTILIDAD, F3ESTNUL, F3FECGRA, USEGRA) "
        strSQLCentroCosto = strSQLCentroCosto & "VALUES("
        strSQLCentroCosto = strSQLCentroCosto & "'" & strCodigo & "', "
        strSQLCentroCosto = strSQLCentroCosto & intCodigoNivel & ", "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strCodigoExterno & "', "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strCodigoConcar & "', "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strDescripcion & "', "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strAbreviatura & "', "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strCodCliente & "', "
        strSQLCentroCosto = strSQLCentroCosto & dblUtilidad & ", "
        strSQLCentroCosto = strSQLCentroCosto & IIf(bolEstado, "'N'", "'S'") & ", "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strFecReg & "', "
        strSQLCentroCosto = strSQLCentroCosto & "'" & strUsuReg & "')"
    Else
        strSQLCentroCosto = vbNullString
        strSQLCentroCosto = strSQLCentroCosto & "UPDATE "
        strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS "
        strSQLCentroCosto = strSQLCentroCosto & "SET "
        strSQLCentroCosto = strSQLCentroCosto & "CODEXT2 = '" & strCodigoExterno & "', "
        strSQLCentroCosto = strSQLCentroCosto & "CCONCAR = '" & strCodigoConcar & "', "
        strSQLCentroCosto = strSQLCentroCosto & "F3DESCRIP = '" & strDescripcion & "', "
        strSQLCentroCosto = strSQLCentroCosto & "F3ABREV = '" & strAbreviatura & "', "
        strSQLCentroCosto = strSQLCentroCosto & "F3CODCLI = '" & strCodCliente & "', "
        strSQLCentroCosto = strSQLCentroCosto & "UTILIDAD = " & dblUtilidad & ", "
        strSQLCentroCosto = strSQLCentroCosto & "F3ESTNUL = " & IIf(bolEstado, "'N', ", "'S', ")
        strSQLCentroCosto = strSQLCentroCosto & "F3FECMOD = '" & strFecMod & "', "
        strSQLCentroCosto = strSQLCentroCosto & "USEMOD = '" & strUsuMod & "' "
        strSQLCentroCosto = strSQLCentroCosto & "WHERE "
        strSQLCentroCosto = strSQLCentroCosto & "F3COSTO = '" & strCodigo & "'"
    End If
    
    cnBdCPlus.Execute strSQLCentroCosto
    
    guardarCentroCosto = True
    
    strSQLSelectAlter = strSQLCentroCosto
    strSQLCentroCosto = vbNullString
    
    Exit Function
errGuardarCentroCosto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsCentroCosto: GuardarCentroCosto"
    
    guardarCentroCosto = False
    
    Err.Clear
End Function

Public Function eliminarCentroCosto() As Boolean
    On Error GoTo errEliminarCentroCosto
    
'    If Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F4OBRA", "REGISDOC", strCodigo, "T")) > 0 Or _
'        Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F3CENCOS", "REGISMOV", strCodigo, "T")) > 0 Or _
'        Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F4CENTRO", "TBVENTA_CAB", strCodigo, "T")) > 0 Or _
'        Val(ModUtilitario.validarUsoRegistro(cnBdCPlus, "F3COSTO", "TBVENTA_DET", strCodigo, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarCentroCosto = False
'
'        Exit Function
'    End If
    
    strSQLCentroCosto = vbNullString
    strSQLCentroCosto = strSQLCentroCosto & "DELETE FROM "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS "
    strSQLCentroCosto = strSQLCentroCosto & "WHERE "
    strSQLCentroCosto = strSQLCentroCosto & "F3COSTO = '" & strCodigo & "'"
    
    cnBdCPlus.Execute strSQLCentroCosto
    
    eliminarCentroCosto = True
    
    strSQLSelectAlter = strSQLCentroCosto
    strSQLCentroCosto = vbNullString
    
    Exit Function
errEliminarCentroCosto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - SqlClsCentroCosto: EliminarCentroCosto"
    
    eliminarCentroCosto = False
    
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: UTILITARIOS DE CLASE :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Vista de Centro de Costo en Grilla (QuamtumGrid)
Public Sub vistaCentroCosto(ByVal grilla As dxDBGrid)
    
    On Error GoTo errVistaCentroCosto
    
    strSQLCentroCosto = vbNullString
    strSQLCentroCosto = strSQLCentroCosto & "SELECT "
    strSQLCentroCosto = strSQLCentroCosto & "C.F3COSTO, "
    strSQLCentroCosto = strSQLCentroCosto & "C.F3DESCRIP, "
    strSQLCentroCosto = strSQLCentroCosto & "C.F3ABREV, "
    strSQLCentroCosto = strSQLCentroCosto & "IIF(C.F3ESTNUL = 'N', 'Habilitado', 'Anulado') AS ESTADO "
    strSQLCentroCosto = strSQLCentroCosto & "FROM "
    strSQLCentroCosto = strSQLCentroCosto & "MAESTROS.CENTROS AS C "
    strSQLCentroCosto = strSQLCentroCosto & "WHERE "
    strSQLCentroCosto = strSQLCentroCosto & "C.F3COSTO <> '' "
        
        If strCodigoOrigen <> vbNullString Then
            strSQLCentroCosto = strSQLCentroCosto & "AND LEFT(F3COSTO, " & Len(strCodigoOrigen) & ") = '" & strCodigoOrigen & "' "
        End If
        
        If intCodigoNivel <> 0 Then
            strSQLCentroCosto = strSQLCentroCosto & "AND INTCODIGONIVEL = " & intCodigoNivel & " "
        End If
        
    strSQLCentroCosto = strSQLCentroCosto & "ORDER BY "
    strSQLCentroCosto = strSQLCentroCosto & "C.F3COSTO"
    
    If Not grilla Is Nothing Then
        With grilla
            .Dataset.Close
            
            .Columns.DestroyColumns
        End With
        
        Dim gColumn As dxGridColumn
        
        With grilla
            'Columna Codigo Interno de Centro de Costo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Codigo"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F3COSTO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColCodigo"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 80
                .Visible = True
            End With
            
            'Columna Nombre de Centro de Costo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taLeftJustify
                .Caption = "Descripcion"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F3DESCRIP"
                .HeaderAlignment = taCenter
                .ObjectName = "ColDescripcion"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 180
                .Visible = True
            End With
            
            'Columna Abreviatura de Centro de Costo
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Abreviatura"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "F3ABREV"
                .HeaderAlignment = taCenter
                .ObjectName = "ColAbreviatura"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 90
                .Visible = True
            End With
            
            'Columna Estado
            Set gColumn = .Columns.Add(gedTextEdit)
            
            With gColumn
                .Alignment = taCenter
                .Caption = "Estado"
                .Color = &HC0FFFF
                .BandIndex = 0
                .DisableEditor = True
                .FieldName = "ESTADO"
                .HeaderAlignment = taCenter
                .ObjectName = "ColEstado"
                .SummaryFooterType = cstCount
                .SummaryFooterFormat = " "
                .Width = 60
                .Visible = True
            End With
            
            .DefaultFields = False
            .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
            
            .Dataset.Active = False
            .Dataset.ADODataset.CommandType = cmdText
            .Dataset.ADODataset.CursorLocation = clUseClient
            .Dataset.ADODataset.CursorType = ctStatic
            .Dataset.ADODataset.LockType = ltReadOnly
            .Dataset.ADODataset.CommandText = strSQLCentroCosto
            .Dataset.Active = True
            .Dataset.Refresh
            .KeyField = "F3COSTO"
        End With
    End If
    
    strSQLCentroCosto = vbNullString
    
    Exit Sub
errVistaCentroCosto:
    Select Case Err.Number
        Case 3704, 3709
            cnBdCPlus.Open strCadenaConexioBdCPlus
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - SqlClsCentroCosto: VistaCentroCosto"
    End Select
    
    Err.Clear
End Sub

