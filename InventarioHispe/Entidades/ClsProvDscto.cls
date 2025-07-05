VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProvDscto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodProv                  As String
Private strCodProd                  As String
Private strDesProd                  As String
Private strCodMed                   As String
Private dblCantidad                 As Double
Private dblPorcentaje               As Double
Private StrFecha                    As String

Private strSQLSelectAlter           As String

Private rstProvDscto             As ADODB.Recordset
Private strSQLProvDscto          As String

Public Property Let CodigoProveedor(ByVal value As String)
    strCodProv = value
End Property

Public Property Get CodigoProveedor() As String
    CodigoProveedor = strCodProv
End Property

Public Property Let CodigoProducto(ByVal value As String)
    strCodProd = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProd
End Property

Public Property Let DescripcionProducto(ByVal value As String)
    strDesProd = value
End Property

Public Property Get DescripcionProducto() As String
    DescripcionProducto = strDesProd
End Property

Public Property Let CodigoUM(ByVal value As String)
    strCodMed = value
End Property

Public Property Get CodigoUM() As String
    CodigoUM = strCodMed
End Property

Public Property Let Cantidad(ByVal value As Double)
    dblCantidad = value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property
'dblPorcentaje
Public Property Let Porcentaje(ByVal value As Double)
    dblPorcentaje = value
End Property

Public Property Get Porcentaje() As Double
    Porcentaje = dblPorcentaje
End Property
'strFecha
Public Property Let Fecha(ByVal value As String)
    StrFecha = value
End Property

Public Property Get Fecha() As String
    Fecha = StrFecha
End Property


Public Property Let SQLSelectAlter(ByVal value As String)
    strSQLSelectAlter = value
End Property

Public Property Get SQLSelectAlter() As String
    SQLSelectAlter = strSQLSelectAlter
End Property


Public Sub inicializarEntidades()
    strCodProv = vbNullString
    strCodProd = vbNullString
    strDesProd = vbNullString
    strCodMed = vbNullString
    dblCantidad = 0
    dblPorcentaje = 0
    StrFecha = vbNullString
    
    strSQLSelectAlter = vbNullString
End Sub

Public Function obtenerProvDscto() As Boolean
    On Error GoTo errObtenerProvDscto
    
    Set rstProvDscto = New ADODB.Recordset
    
    strSQLProvDscto = vbNullString
    strSQLProvDscto = strSQLProvDscto & "SELECT "
    strSQLProvDscto = strSQLProvDscto & "* "
    strSQLProvDscto = strSQLProvDscto & "FROM EF2PROV_DSCTO "
    strSQLProvDscto = strSQLProvDscto & "WHERE "
    strSQLProvDscto = strSQLProvDscto & "F2CODPRV = '" & strCodProv & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F5CODPRO = '" & strCodProd & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F7CODMED = '" & strCodMed & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F3CANPRO = " & dblCantidad & " AND "
    strSQLProvDscto = strSQLProvDscto & "CVDATE(F2FECHA) = CVDATE('" & StrFecha & "')"
    
    If rstProvDscto.State = 1 Then rstProvDscto.Close
    
    rstProvDscto.Open strSQLProvDscto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProvDscto.EOF Then
        strCodProv = Trim(rstProvDscto!F2CODPRV & "")
        strCodProd = Trim(rstProvDscto!f5codpro & "")
        strDesProd = Trim(rstProvDscto!f5nompro & "")
        strCodMed = Trim(rstProvDscto!f7codmed & "")
        dblCantidad = Val(rstProvDscto!F3CANPRO & "")
        dblPorcentaje = Val(rstProvDscto!f3pordct & "")
        StrFecha = Trim(rstProvDscto!F2FECHA & "")
        
        obtenerProvDscto = True
    Else
        obtenerProvDscto = False
    End If
    
    rstProvDscto.Close
    
    Set rstProvDscto = Nothing
    
    Exit Function
errObtenerProvDscto:
    obtenerProvDscto = False
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProvDscto: ObtenerProvDscto"
    
    Err.Clear
End Function

Public Sub obtenerConfigProvDscto()
    On Error GoTo errObtenerConfigProvDscto
    
    Set rstProvDscto = New ADODB.Recordset
    
    strSQLProvDscto = vbNullString
    strSQLProvDscto = strSQLProvDscto & "SELECT "
    strSQLProvDscto = strSQLProvDscto & "* "
    strSQLProvDscto = strSQLProvDscto & "FROM EF2PROV_DSCTO "
    strSQLProvDscto = strSQLProvDscto & "WHERE "
    strSQLProvDscto = strSQLProvDscto & "F2CODPRV = '" & strCodProv & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F5CODPRO = '" & strCodProd & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F7CODMED = '" & strCodMed & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F3CANPRO = " & dblCantidad & " AND "
    strSQLProvDscto = strSQLProvDscto & "CVDATE(F2FECHA) = CVDATE('" & StrFecha & "')"
    
    If rstProvDscto.State = 1 Then rstProvDscto.Close
    
    rstProvDscto.Open strSQLProvDscto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProvDscto.EOF Then
        strCodProv = Trim(rstProvDscto!F2CODPRV & "")
        strCodProd = Trim(rstProvDscto!f5codpro & "")
        strDesProd = Trim(rstProvDscto!f5nompro & "")
        strCodMed = Trim(rstProvDscto!f7codmed & "")
        dblCantidad = Val(rstProvDscto!F3CANPRO & "")
        dblPorcentaje = Val(rstProvDscto!f3pordct & "")
        StrFecha = Trim(rstProvDscto!F2FECHA & "")
    End If
    
    rstProvDscto.Close
    
    Set rstProvDscto = Nothing
    
    Exit Sub
errObtenerConfigProvDscto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProvDscto: ObtenerProvDscto"
    
    Err.Clear
End Sub

Public Function verificarExistencia() As Boolean
    On Error GoTo errVerificarExistencia
    
    Set rstProvDscto = New ADODB.Recordset
    
    strSQLProvDscto = vbNullString
    strSQLProvDscto = strSQLProvDscto & "SELECT "
    strSQLProvDscto = strSQLProvDscto & "EF2PROV_DSCTO.* "
    strSQLProvDscto = strSQLProvDscto & "FROM EF2PROV_DSCTO "
    strSQLProvDscto = strSQLProvDscto & "WHERE "
    strSQLProvDscto = strSQLProvDscto & "F2CODPRV = '" & strCodProv & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F5CODPRO = '" & strCodProd & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F7CODMED = '" & strCodMed & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F3CANPRO = " & dblCantidad & " AND "
    strSQLProvDscto = strSQLProvDscto & "CVDATE(F2FECHA) = CVDATE('" & StrFecha & "')"
    
    If rstProvDscto.State = 1 Then rstProvDscto.Close
    
    rstProvDscto.Open strSQLProvDscto, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    If Not rstProvDscto.EOF Then
        verificarExistencia = True
    Else
        verificarExistencia = False
    End If
    
    rstProvDscto.Close
    
    strSQLProvDscto = vbNullString
    
    Set rstProvDscto = Nothing
    
    Exit Function
errVerificarExistencia:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProvDscto: VerificarExistencia"
    
    verificarExistencia = False
    
    Err.Clear
End Function

Public Function guardarProvDscto() As Boolean
    On Error GoTo errGuardarProvDscto
    
    strSQLProvDscto = vbNullString
    
    'cnn_dbbancos.BeginTrans
    
    If Not verificarExistencia Then
        strSQLProvDscto = strSQLProvDscto & "INSERT INTO EF2PROV_DSCTO("
        strSQLProvDscto = strSQLProvDscto & "F2CODPRV, F5CODPRO, F5NOMPRO, F7CODMED, F3CANPRO, F3PORDCT, F2FECHA) "
        strSQLProvDscto = strSQLProvDscto & "VALUES("
        strSQLProvDscto = strSQLProvDscto & "'" & strCodProv & "', "
        strSQLProvDscto = strSQLProvDscto & "'" & strCodProd & "', "
        strSQLProvDscto = strSQLProvDscto & "'" & strDesProd & "', "
        strSQLProvDscto = strSQLProvDscto & "'" & strCodMed & "', "
        strSQLProvDscto = strSQLProvDscto & dblCantidad & ", "
        strSQLProvDscto = strSQLProvDscto & dblPorcentaje & ", "
        strSQLProvDscto = strSQLProvDscto & "CVDATE('" & StrFecha & "'))"
    Else
        strSQLProvDscto = strSQLProvDscto & "UPDATE "
        strSQLProvDscto = strSQLProvDscto & "EF2PROV_DSCTO "
        strSQLProvDscto = strSQLProvDscto & "SET "
        strSQLProvDscto = strSQLProvDscto & "F5NOMPRO = '" & strDesProd & "', "
        strSQLProvDscto = strSQLProvDscto & "F3PORDCT = " & dblPorcentaje & " "
        strSQLProvDscto = strSQLProvDscto & "WHERE "
        strSQLProvDscto = strSQLProvDscto & "F2CODPRV = '" & strCodProv & "' AND "
        strSQLProvDscto = strSQLProvDscto & "F5CODPRO = '" & strCodProd & "' AND "
        strSQLProvDscto = strSQLProvDscto & "F7CODMED = '" & strCodMed & "' AND "
        strSQLProvDscto = strSQLProvDscto & "F3CANPRO = " & dblCantidad & " AND "
        strSQLProvDscto = strSQLProvDscto & "CVDATE(F2FECHA) = CVDATE('" & StrFecha & "')"
    End If
    
    cnn_dbbancos.Execute strSQLProvDscto
    
    'cnn_dbbancos.CommitTrans
    
    guardarProvDscto = True
    
    strSQLProvDscto = vbNullString
    
    Exit Function
errGuardarProvDscto:
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProvDscto: GuardarProvDscto"
    
    'cnn_dbbancos.RollbackTrans
    
    guardarProvDscto = False
    
    Err.Clear
End Function

Public Function eliminarProvDscto() As Boolean
    On Error GoTo errEliminarProvDscto
    
'    If Val(ModUtilitario.validarUsoRegistro(cnn_dbbancos, "CODIGO", "IF3ORDEN", strCodProv, "T")) > 0 Then
'
'        MsgBox "Imposible eliminar registro actual, esta relacionado con otras tablas.", vbInformation + vbOKOnly, App.ProductName
'
'        eliminarProvDscto = False
'
'        Exit Function
'    End If
    
    'cnn_dbbancos.BeginTrans
    
    strSQLProvDscto = vbNullString
    strSQLProvDscto = strSQLProvDscto & "DELETE "
    strSQLProvDscto = strSQLProvDscto & "FROM "
    strSQLProvDscto = strSQLProvDscto & "EF2PROV_DSCTO "
    strSQLProvDscto = strSQLProvDscto & "WHERE "
    strSQLProvDscto = strSQLProvDscto & "F2CODPRV = '" & strCodProv & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F5CODPRO = '" & strCodProd & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F7CODMED = '" & strCodMed & "' AND "
    strSQLProvDscto = strSQLProvDscto & "F3CANPRO = " & dblCantidad & " AND "
    strSQLProvDscto = strSQLProvDscto & "CVDATE(F2FECHA) = CVDATE('" & StrFecha & "')"
    
    cnn_dbbancos.Execute strSQLProvDscto
    
    'cnn_dbbancos.CommitTrans
    
    eliminarProvDscto = True
    
    strSQLProvDscto = vbNullString
    
    Exit Function
errEliminarProvDscto:
    eliminarProvDscto = False

    'cnn_dbbancos.RollbackTrans
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ClsProvDscto: EliminarProvDscto"
    
    Err.Clear
End Function
