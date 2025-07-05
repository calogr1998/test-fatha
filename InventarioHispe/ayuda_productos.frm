VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ayuda_productos 
   Caption         =   "Ayuda de Productos"
   ClientHeight    =   7815
   ClientLeft      =   135
   ClientTop       =   1755
   ClientWidth     =   20235
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_productos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   20235
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAlmacen 
      Caption         =   " Almacen "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   8
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Timer timTemporizador 
      Interval        =   1000
      Left            =   0
      Top             =   360
   End
   Begin ActiveToolBars.SSActiveToolBars tlbProducto 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   12
      Tools           =   "ayuda_productos.frx":058A
      ToolBars        =   "ayuda_productos.frx":B69D
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgProducto 
      Height          =   6465
      Left            =   120
      OleObjectBlob   =   "ayuda_productos.frx":B85A
      TabIndex        =   4
      Top             =   960
      Width           =   20025
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   " Ingresar cadena de texto a buscar "
      Height          =   735
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Width           =   9705
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar todos los productos"
         Height          =   240
         Left            =   7920
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Incluir Productos sin Stock"
         Height          =   240
         Left            =   7920
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   9165
      End
      Begin MSComctlLib.ProgressBar pgbProgresoBusqueda 
         Height          =   135
         Left            =   360
         TabIndex        =   7
         Top             =   555
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   25
         Scrolling       =   1
      End
   End
   Begin CONTROLSLibCtl.dxCheckBox chkProductoProveedor 
      Height          =   270
      Left            =   13080
      TabIndex        =   6
      Top             =   600
      Width           =   2115
      _Version        =   65536
      _cx             =   3731
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Productos del Proveedor"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox chkSeleccionActual 
      Height          =   270
      Left            =   13080
      TabIndex        =   5
      Top             =   240
      Width           =   2760
      _Version        =   65536
      _cx             =   4868
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Mostrar productos seleccionados"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
End
Attribute VB_Name = "ayuda_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem SK ADD:
Private strCodAuxiliar As String
Private strCodRequerimiento As String
Private strCodProducto As String

Private strCadenaCorte As String

'Propiedad Codigo de Auxiliar
Public Property Let CodigoAuxiliar(ByVal Value As String)
    strCodAuxiliar = Value
End Property

Public Property Get CodigoAuxiliar() As String
    CodigoAuxiliar = strCodAuxiliar
End Property

'Propiedad Codigo de Requerimiento
Public Property Let CodigoRequerimiento(ByVal Value As String)
    strCodRequerimiento = Value
End Property

Public Property Get CodigoRequerimiento() As String
    CodigoRequerimiento = strCodRequerimiento
End Property

'Propiedad Codigo de Producto
Public Property Let CodigoProducto(ByVal Value As String)
    strCodProducto = Value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

'Propiedad Cadena de Corte de Informacion
Public Property Let CadenaCorte(ByVal Value As String)
    strCadenaCorte = Value
End Property

Public Property Get CadenaCorte() As String
    CadenaCorte = strCadenaCorte
End Property


Private Sub listarAlmacenEnCombo()
    Dim rstAlmacen As New ADODB.Recordset
    
    If rstAlmacen.State = 1 Then rstAlmacen.Close
    
    rstAlmacen.Open "SELECT F2CODALM, F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    
    cmbAlmacen.Clear
    
    If Not rstAlmacen.EOF Then
        rstAlmacen.MoveFirst
        
        Do While Not rstAlmacen.EOF
            cmbAlmacen.AddItem Trim(rstAlmacen!F2NOMALM & "") & Space(100) & Trim(rstAlmacen!f2codalm & "")
            
            rstAlmacen.MoveNext
        Loop
            If cmbAlmacen.ListCount > 0 Then
                cmbAlmacen.ListIndex = 0
            End If
    End If
End Sub

Private Sub cargarProductos()
    dbgProducto.Dataset.Close
    
    abrirCnTemporal
    
    ModUtilitario.borrarTablaEnBD cnDBTemp, "TMPPRODUCTOS"
    
'    If strCodRequerimiento <> vbNullString And strCodProducto = vbNullString Then
'        Rem SK ADD:
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "VAL(STOCK.CANTIDAD & '') AS F6STOCKACT, "
'        SqlCad = SqlCad & "PROD.F4PERINT, "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
'        SqlCad = SqlCad & "DET.COD_PRODUCTO AS F5CODPRO, "
'        SqlCad = SqlCad & "HISTORIAL.F5CODFAB, "
'        SqlCad = SqlCad & "MAR.F2DESMAR, "
'        SqlCad = SqlCad & "IIF(TRIM(HISTORIAL.F5NOMPRO & '') = '', DET.DS_DESCRIPCION, HISTORIAL.F5NOMPRO) AS F5NOMPRO, "
'        SqlCad = SqlCad & "DET.DS_DESCRIPCION AS F5NOMPRO_ING, "
'        SqlCad = SqlCad & "DET.DS_UNIDMED AS F7CODMED, "
'        SqlCad = SqlCad & "MED.F7SIGMED, "
'        SqlCad = SqlCad & "0 AS F5VALVTA, "
'        SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS F5FOB, "
'        SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS F5FOBMAX, "
'        SqlCad = SqlCad & "PROD.F5AFECTO, "
'        SqlCad = SqlCad & "HISTORIAL.FECHA AS F5FECUC, "
'        SqlCad = SqlCad & "HISTORIAL.F2MONEDA, "
'        SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'S', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANET, "
'        SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'D', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANETDOL "
'
'        SqlCad = SqlCad & "INTO TMPPRODUCTOS IN '" & wrutatemp & "TEMPLUS.MDB' "
'
'        SqlCad = SqlCad & "FROM "
'
'        SqlCad = SqlCad & "((((((TB_DETSOLICITUD AS DET "
'
'        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CAB ON CAB.COD_SOLICITUD = DET.COD_SOLICITUD) "
'        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.COD_PRODUCTO) "
'        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.DS_UNIDMED) "
'        SqlCad = SqlCad & "LEFT JOIN EF2MARCAS AS MAR ON MAR.F2CODMAR = PROD.F5MARCA) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "("
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
'        SqlCad = SqlCad & "DET.F3CODPRO, "
'        SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "IF3ORDEN AS DET "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD, DET.F3CODPRO) AS MOVPROD "
'        SqlCad = SqlCad & "ON MOVPROD.COD_SOLICITUD = DET.COD_SOLICITUD AND MOVPROD.F3CODPRO = DET.COD_PRODUCTO) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "("
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "F5CODPRO, F2CODALM, SUM(F3CANPRO * IIF(LEFT(F4NUMVAL,1) = 'I', 1, -1)) AS CANTIDAD "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "IF3VALES "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "F2CODALM = '" & right(cmbAlmacen.Text, 2) & "' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "F5CODPRO, F2CODALM) AS STOCK ON STOCK.F5CODPRO = PROD.F5CODPRO) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "("
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "PP.F5CODPRO, PP.F5NOMPRO, PP.F5CODFAB, "
'        SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)) AS F5VALVTA, "
'        SqlCad = SqlCad & "PP.F2MONEDA, "
'        SqlCad = SqlCad & "MAX(PP.F2FECHA) AS FECHA "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "EF2PROD_PROV AS PP "
'        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MV ON MV.F5CODPRO = PP.F5CODPRO AND MV.F7CODMED = PP.F7CODMED "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "PP.F2CODPRV = '" & strCodAuxiliar & "' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "PP.F5CODPRO, "
'        SqlCad = SqlCad & "PP.F5NOMPRO, "
'        SqlCad = SqlCad & "PP.F5CODFAB, "
'        SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)), PP.F2MONEDA) AS HISTORIAL "
'        SqlCad = SqlCad & "ON HISTORIAL.F5CODPRO = PROD.F5CODPRO "
'
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) > 0 AND "
'        SqlCad = SqlCad & "CAB.CS_ESTADO NOT IN ('0', '1', '5') AND "
'        SqlCad = SqlCad & "CAB.COD_SOLICITUD = '" & strCodRequerimiento & "' "
'        SqlCad = SqlCad & "ORDER BY "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD DESC, DET.DS_DESCRIPCION"
'    ElseIf strCodRequerimiento = vbNullString And strCodProducto <> vbNullString Then
'        Rem SK ADD:
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "VAL(STOCK.CANTIDAD & '') AS F6STOCKACT, "
'        SqlCad = SqlCad & "PROD.F4PERINT, "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
'        SqlCad = SqlCad & "DET.COD_PRODUCTO AS F5CODPRO, "
'        SqlCad = SqlCad & "HISTORIAL.F5CODFAB, "
'        SqlCad = SqlCad & "MAR.F2DESMAR, "
'        SqlCad = SqlCad & "IIF(TRIM(HISTORIAL.F5NOMPRO & '') = '', DET.DS_DESCRIPCION, HISTORIAL.F5NOMPRO) AS F5NOMPRO, "
'        SqlCad = SqlCad & "DET.DS_DESCRIPCION AS F5NOMPRO_ING, "
'        SqlCad = SqlCad & "DET.DS_UNIDMED AS F7CODMED, "
'        SqlCad = SqlCad & "MED.F7SIGMED, "
'        SqlCad = SqlCad & "0 AS F5VALVTA, "
'        SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS F5FOB, "
'        SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS F5FOBMAX, "
'        SqlCad = SqlCad & "PROD.F5AFECTO, "
'        SqlCad = SqlCad & "HISTORIAL.FECHA AS F5FECUC, "
'        SqlCad = SqlCad & "HISTORIAL.F2MONEDA, "
'        SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'S', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANET, "
'        SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'D', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANETDOL "
'
'        SqlCad = SqlCad & "INTO TMPPRODUCTOS IN '" & wrutatemp & "TEMPLUS.MDB' "
'
'        SqlCad = SqlCad & "FROM "
'
'        SqlCad = SqlCad & "((((((TB_DETSOLICITUD AS DET "
'
'        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CAB ON CAB.COD_SOLICITUD = DET.COD_SOLICITUD) "
'        SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.COD_PRODUCTO) "
'        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.DS_UNIDMED) "
'        SqlCad = SqlCad & "LEFT JOIN EF2MARCAS AS MAR ON MAR.F2CODMAR = PROD.F5MARCA) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "("
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
'        SqlCad = SqlCad & "DET.F3CODPRO, "
'        SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "IF3ORDEN AS DET "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD, DET.F3CODPRO) AS MOVPROD "
'        SqlCad = SqlCad & "ON MOVPROD.COD_SOLICITUD = DET.COD_SOLICITUD AND  MOVPROD.F3CODPRO = DET.COD_PRODUCTO) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "("
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "F5CODPRO, F2CODALM, SUM(F3CANPRO * IIF(LEFT(F4NUMVAL,1) = 'I', 1, -1)) AS CANTIDAD "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "IF3VALES "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "F2CODALM = '" & right(cmbAlmacen.Text, 2) & "' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "F5CODPRO, F2CODALM) AS STOCK ON STOCK.F5CODPRO = PROD.F5CODPRO) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "("
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "PP.F5CODPRO, PP.F5NOMPRO, PP.F5CODFAB, "
'        SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)) AS F5VALVTA, "
'        SqlCad = SqlCad & "PP.F2MONEDA, "
'        SqlCad = SqlCad & "MAX(PP.F2FECHA) AS FECHA "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "EF2PROD_PROV AS PP "
'        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MV ON MV.F5CODPRO = PP.F5CODPRO AND MV.F7CODMED = PP.F7CODMED "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "PP.F2CODPRV = '" & strCodAuxiliar & "' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "PP.F5CODPRO, "
'        SqlCad = SqlCad & "PP.F5NOMPRO, "
'        SqlCad = SqlCad & "PP.F5CODFAB, "
'        SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)), PP.F2MONEDA) AS HISTORIAL "
'        SqlCad = SqlCad & "ON HISTORIAL.F5CODPRO = PROD.F5CODPRO "
'
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) > 0 AND "
'        SqlCad = SqlCad & "CAB.CS_ESTADO NOT IN ('0', '1', '5') AND "
'        SqlCad = SqlCad & "DET.COD_PRODUCTO = '" & strCodProducto & "' "
'        SqlCad = SqlCad & "ORDER BY "
'        SqlCad = SqlCad & "DET.COD_SOLICITUD DESC, DET.DS_DESCRIPCION"
'    Else
'        Rem  SK ADD:
'        SqlCad = vbNullString
'        SqlCad = SqlCad & "SELECT "
'        SqlCad = SqlCad & "VAL(STOCK.CANTIDAD & '') AS F6STOCKACT, "
'        SqlCad = SqlCad & "PROD.F4PERINT, "
'
'        SqlCad = SqlCad & "'' AS COD_SOLICITUD, "
'        SqlCad = SqlCad & "PROD.F5CODPRO, "
'
'        SqlCad = SqlCad & "HISTORIAL.F5CODFAB, "
'        SqlCad = SqlCad & "MAR.F2DESMAR, "
'
'        SqlCad = SqlCad & "IIF(TRIM(HISTORIAL.F5NOMPRO & '') = '', PROD.F5NOMPRO, HISTORIAL.F5NOMPRO) AS F5NOMPRO, "
'        SqlCad = SqlCad & "PROD.F5NOMPRO AS F5NOMPRO_ING, "
'        SqlCad = SqlCad & "PROD.F7CODMED, "
'
'        SqlCad = SqlCad & "MED.F7SIGMED, "
'        SqlCad = SqlCad & "0 AS F5VALVTA, "
'
'        SqlCad = SqlCad & "0 AS F5FOB, "
'        SqlCad = SqlCad & "0 AS F5FOBMAX, "
'
'        SqlCad = SqlCad & "PROD.F5AFECTO, "
'        SqlCad = SqlCad & "HISTORIAL.FECHA AS F5FECUC, "
'        SqlCad = SqlCad & "HISTORIAL.F2MONEDA, "
'        SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'S', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANET, "
'        SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'D', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANETDOL "
'
'        SqlCad = SqlCad & "INTO TMPPRODUCTOS IN '" & wrutatemp & "TEMPLUS.MDB' "
'
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "(((IF5PLA AS PROD "
'        SqlCad = SqlCad & "LEFT JOIN EF2MARCAS AS MAR ON MAR.F2CODMAR = PROD.F5MARCA) "
'        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "(SELECT "
'        SqlCad = SqlCad & "F5CODPRO, F2CODALM, SUM(F3CANPRO * IIF(LEFT(F4NUMVAL,1) = 'I', 1, -1)) AS CANTIDAD "
'        SqlCad = SqlCad & "FROM IF3VALES "
'        SqlCad = SqlCad & "WHERE F2CODALM = '" & right(cmbAlmacen.Text, 2) & "' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "F5CODPRO, F2CODALM) AS STOCK ON STOCK.F5CODPRO = PROD.F5CODPRO) "
'
'        SqlCad = SqlCad & "LEFT JOIN "
'        SqlCad = SqlCad & "(SELECT "
'        SqlCad = SqlCad & "PP.F5CODPRO, PP.F5NOMPRO, PP.F5CODFAB, "
'        SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)) AS F5VALVTA, "
'        SqlCad = SqlCad & "PP.F2MONEDA, MAX(PP.F2FECHA) AS FECHA "
'        SqlCad = SqlCad & "FROM "
'        SqlCad = SqlCad & "EF2PROD_PROV AS PP "
'        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MV ON MV.F5CODPRO = PP.F5CODPRO AND MV.F7CODMED = PP.F7CODMED "
'        SqlCad = SqlCad & "WHERE "
'        SqlCad = SqlCad & "PP.F2CODPRV = '" & strCodAuxiliar & "' "
'        SqlCad = SqlCad & "GROUP BY "
'        SqlCad = SqlCad & "PP.F5CODPRO, PP.F5NOMPRO, PP.F5CODFAB, (PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)), PP.F2MONEDA) AS HISTORIAL "
'        SqlCad = SqlCad & "ON HISTORIAL.F5CODPRO = PROD.F5CODPRO "
'    End If
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    'SqlCad = SqlCad & "VAL(STOCK.CANTIDAD & '') AS F6STOCKACT, "
    
    SqlCad = SqlCad & "VAL(STOCK.STOCKCOMPROMETIDO & '') AS STOCKCOMPROMETIDO, " 'COMPROMETIDO EN ALMACEN
    SqlCad = SqlCad & "VAL(STOCK.STOCKPORLLEGARCOMP & '') AS STOCKPORLLEGARCOMP, " 'COMPROMETIDO POR LLEGAR
    SqlCad = SqlCad & "VAL(STOCK.STOCKLIBRE & '') AS STOCKLIBRE, " 'LIBRE EN ALMACEN
    SqlCad = SqlCad & "VAL(STOCK.STOCKPORLLEGARLIBRE & '') AS STOCKPORLLEGARLIBRE, " 'LIBRE POR LLEGAR
    
    SqlCad = SqlCad & "VAL(STOCK.STOCK & '') AS STOCK, "
    SqlCad = SqlCad & "PROD.F4PERINT, "
    
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "DET.COD_SOLICITUD, "
            SqlCad = SqlCad & "DET.COD_PRODUCTO AS F5CODPRO, "
        Else
            SqlCad = SqlCad & "'' AS COD_SOLICITUD, "
            SqlCad = SqlCad & "PROD.F5CODPRO, "
        End If
    
    SqlCad = SqlCad & "HISTORIAL.F5CODFAB, "
    SqlCad = SqlCad & "MAR.F2DESMAR, "
    
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            'SqlCad = SqlCad & "IIF(TRIM(HISTORIAL.F5NOMPRO & '') = '', DET.DS_DESCRIPCION, HISTORIAL.F5NOMPRO) AS F5NOMPRO, "
            SqlCad = SqlCad & "DET.DS_DESCRIPCION AS F5NOMPRO, "
            SqlCad = SqlCad & "DET.DS_DESCRIPCION AS F5NOMPRO_ING, "
            SqlCad = SqlCad & "DET.DS_UNIDMED AS F7CODMED, "
        Else
            SqlCad = SqlCad & "PROD.F5NOMPRO AS F5NOMPRO, "
            SqlCad = SqlCad & "PROD.F5NOMPRO AS F5NOMPRO_ING, "
            SqlCad = SqlCad & "PROD.F7CODMED, "
        End If
        
    SqlCad = SqlCad & "MED.F7SIGMED, "
    SqlCad = SqlCad & "0 AS F5VALVTA, "
    
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            If strCodRequerimiento = vbNullString Then
                SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS F5FOB, "
            ElseIf strCodProducto = vbNullString Then
                SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) - (VAL(STOCK.STOCKCOMPROMETIDO & '') + VAL(STOCK.STOCKPORLLEGARCOMP & '')) AS F5FOB, "
            End If
            
            SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) AS F5FOBMAX, "
        Else
            SqlCad = SqlCad & "0.0000001 AS F5FOB, "
            SqlCad = SqlCad & "0.0000001 AS F5FOBMAX, "
        End If
        
    SqlCad = SqlCad & "PROD.F5AFECTO, "
    SqlCad = SqlCad & "HISTORIAL.FECHA AS F5FECUC, "
    SqlCad = SqlCad & "HISTORIAL.F2MONEDA, "
    SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'S', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANET, "
    SqlCad = SqlCad & "IIF(HISTORIAL.F2MONEDA = 'D', VAL(HISTORIAL.F5VALVTA & ''), NULL) AS F5VTANETDOL "
    
    SqlCad = SqlCad & "INTO TMPPRODUCTOS IN '" & wrutatemp & "TEMPLUS.MDB' "
    
    SqlCad = SqlCad & "FROM "
            
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "(((((("
            SqlCad = SqlCad & "TB_DETSOLICITUD AS DET "
            SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS CAB ON CAB.COD_SOLICITUD = DET.COD_SOLICITUD) "
            SqlCad = SqlCad & "LEFT JOIN IF5PLA AS PROD ON PROD.F5CODPRO = DET.COD_PRODUCTO) "
            SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.DS_UNIDMED) "
        Else
            SqlCad = SqlCad & "((("
            SqlCad = SqlCad & "IF5PLA AS PROD "
            SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PROD.F7CODMED) "
        End If
        
    
    SqlCad = SqlCad & "LEFT JOIN EF2MARCAS AS MAR ON MAR.F2CODMAR = PROD.F5MARCA) "
    
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "LEFT JOIN "
            SqlCad = SqlCad & "("
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "DET.COD_SOLICITUD, "
            SqlCad = SqlCad & "DET.F3CODPRO, "
            SqlCad = SqlCad & "SUM(DET.F3CANPRO) AS CANTIDAD "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "IF3ORDEN AS DET "
            SqlCad = SqlCad & "WHERE "
            SqlCad = SqlCad & "DET.COD_SOLICITUD <> '' "
            SqlCad = SqlCad & "GROUP BY "
            SqlCad = SqlCad & "DET.COD_SOLICITUD, DET.F3CODPRO) AS MOVPROD "
            SqlCad = SqlCad & "ON MOVPROD.COD_SOLICITUD = DET.COD_SOLICITUD AND MOVPROD.F3CODPRO = DET.COD_PRODUCTO) "
        End If
        
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "("
    'SqlCad = SqlCad & "SELECT "
    'SqlCad = SqlCad & "F5CODPRO, F2CODALM, SUM(F3CANPRO * IIF(TIPO = 'S', -1,1)) AS CANTIDAD "
    'SqlCad = SqlCad & "FROM "
    'SqlCad = SqlCad & "IF3VALES "
    'SqlCad = SqlCad & "WHERE "
    'SqlCad = SqlCad & "F2CODALM = '" & right(cmbAlmacen.Text, 2) & "' "
    'SqlCad = SqlCad & "GROUP BY "
    'SqlCad = SqlCad & "F5CODPRO, F2CODALM"
    
        With objAyudaVale
            .listarGrillaMovimientoProductoResumen Nothing, right(cmbAlmacen.Text, 2), vbNullString, strCodRequerimiento
            
            SqlCad = SqlCad & .SQLSelectAlter
        End With
    
    SqlCad = SqlCad & ") AS STOCK ON STOCK.F5CODPRO = PROD.F5CODPRO) "
    
    SqlCad = SqlCad & "LEFT JOIN "
    SqlCad = SqlCad & "("
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "PP.F5CODPRO, PP.F5NOMPRO, PP.F5CODFAB, "
    SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)) AS F5VALVTA, "
    SqlCad = SqlCad & "PP.F2MONEDA, "
    SqlCad = SqlCad & "MAX(PP.F2FECHA) AS FECHA "
    SqlCad = SqlCad & "FROM "
    SqlCad = SqlCad & "EF2PROD_PROV AS PP "
    SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MV ON MV.F5CODPRO = PP.F5CODPRO AND MV.F7CODMED = PP.F7CODMED "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "PP.F2CODPRV = '" & strCodAuxiliar & "' "
    SqlCad = SqlCad & "GROUP BY "
    SqlCad = SqlCad & "PP.F5CODPRO, "
    SqlCad = SqlCad & "PP.F5NOMPRO, "
    SqlCad = SqlCad & "PP.F5CODFAB, "
    SqlCad = SqlCad & "(PP.F5VALVTA / IIF(ISNULL(MV.F5FACTOR), 1, MV.F5FACTOR)), PP.F2MONEDA) AS HISTORIAL "
    SqlCad = SqlCad & "ON HISTORIAL.F5CODPRO = PROD.F5CODPRO "
    
    If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
        SqlCad = SqlCad & "WHERE "
        
        'SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) > 0 AND "
        SqlCad = SqlCad & "PROD.F5DESCONTINUADO = 'N' AND "
        
        If strCodRequerimiento = vbNullString Then
            SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) > 0 AND "
        ElseIf strCodProducto = vbNullString Then
            SqlCad = SqlCad & "(DET.DS_CANTIDAD - VAL(MOVPROD.CANTIDAD & '')) - (VAL(STOCK.STOCKCOMPROMETIDO & '') + VAL(STOCK.STOCKPORLLEGARCOMP & '')) > 0 AND "
        End If
        
        SqlCad = SqlCad & "CAB.CS_ESTADO NOT IN ('0', '1', '5') AND "
            
            If strCodRequerimiento <> vbNullString Then
                SqlCad = SqlCad & "CAB.COD_SOLICITUD = '" & strCodRequerimiento & "' "
            ElseIf strCodProducto <> vbNullString Then
                SqlCad = SqlCad & "DET.COD_PRODUCTO = '" & strCodProducto & "' "
            End If
            
            If strCadenaCorte <> vbNullString Then
                SqlCad = SqlCad & "AND IIF(TRIM(HISTORIAL.F5NOMPRO & '') = '', DET.DS_DESCRIPCION, HISTORIAL.F5NOMPRO) LIKE '%" & strCadenaCorte & "%' "
                'SqlCad = SqlCad & "DET.DS_DESCRIPCION LIKE '%" & strCadenaCorte & "%' "
            End If
    Else
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "PROD.F5DESCONTINUADO = 'N' "
        
        If strCadenaCorte <> vbNullString Then
            SqlCad = SqlCad & "AND IIF(TRIM(HISTORIAL.F5NOMPRO & '') = '', PROD.F5NOMPRO, HISTORIAL.F5NOMPRO) LIKE '%" & strCadenaCorte & "%' "
            'SqlCad = SqlCad & "PROD.F5NOMPRO LIKE '%" & strCadenaCorte & "%' "
        End If
    End If
    
    SqlCad = SqlCad & "ORDER BY "
    
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "DET.COD_SOLICITUD DESC, DET.DS_DESCRIPCION "
        Else
            SqlCad = SqlCad & "PROD.F5NOMPRO"
        End If
    
    cnn_dbbancos.Execute SqlCad
    
    If strCodRequerimiento = vbNullString Or strCodProducto = vbNullString Then
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "TMPPRODUCTOS "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "F5FOB = 0 "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F5FOB = 0.0000001"
        
        abrirCnTemporal
        
        cnDBTemp.Execute SqlCad
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "UPDATE "
        SqlCad = SqlCad & "TMPPRODUCTOS "
        SqlCad = SqlCad & "SET "
        SqlCad = SqlCad & "F5FOBMAX = 0 "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "F5FOBMAX = 0.0000001"
        
        abrirCnTemporal
        
        cnDBTemp.Execute SqlCad
    End If
    
    SqlCad = vbNullString
End Sub

Public Sub listarProductos()
    dbgProducto.Dataset.Close
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "SELECT "
    SqlCad = SqlCad & "* "
    SqlCad = SqlCad & "FROM TMPPRODUCTOS "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(F5CODPRO & '') <> '' "
        
        If chkSeleccionActual.Checked Then
            SqlCad = SqlCad & "AND F4PERINT = -1 "
        End If
        
        If chkProductoProveedor.Checked Then
            SqlCad = SqlCad & "AND NOT ISNULL(F5FECUC) "
        End If
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND ("
            'SqlCad = SqlCad & "F5CODPRO LIKE '%" & txtBusqueda.Text & "%' OR "
            'SqlCad = SqlCad & "F5CODFAB LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "F5NOMPRO LIKE '%" & txtBusqueda.Text & "%'"
                
                If strCodProducto <> vbNullString Then
                    SqlCad = SqlCad & " OR COD_SOLICITUD LIKE '%" & txtBusqueda.Text & "%'"
                End If
                
            SqlCad = SqlCad & ") "
        End If
    
    SqlCad = SqlCad & "ORDER BY "
        
        If strCodProducto <> vbNullString Then
            SqlCad = SqlCad & "COD_SOLICITUD DESC, "
        End If
        
    SqlCad = SqlCad & "F5NOMPRO"
    
    With dbgProducto
        abrirCnTemporal
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnDBTemp.ConnectionString
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        
        'If strCodRequerimiento = vbNullString And strCodProducto <> vbNullString Then
        '    .KeyField = "COD_SOLICITUD"
        'Else
        '    .KeyField = "F5CODPRO"
        'End If
        
        Select Case strCodProducto
            Case Is <> vbNullString
                .KeyField = "COD_SOLICITUD"
                
                .Columns.ColumnByFieldName("F5FOB").SummaryFooterType = cstSum
            Case Else
                .KeyField = "F5CODPRO"
        End Select
        
        .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterType = cstCount
        .Columns.ColumnByFieldName("F5NOMPRO").SummaryFooterFormat = "Cantidad de Registros = " & .Dataset.RecordCount
        
        .Columns.ColumnByFieldName("STOCKCOMPROMETIDO").Color = &HC0&
            .Columns.ColumnByFieldName("STOCKCOMPROMETIDO").Font.Bold = True
            .Columns.ColumnByFieldName("STOCKCOMPROMETIDO").FontColor = &HFFFFFF
            
        .Columns.ColumnByFieldName("STOCKPORLLEGARCOMP").Color = &H80FFFF
            .Columns.ColumnByFieldName("STOCKPORLLEGARCOMP").Font.Bold = False
            .Columns.ColumnByFieldName("STOCKPORLLEGARCOMP").FontColor = &H80000012
        
        .Columns.ColumnByFieldName("STOCKLIBRE").Color = &HC000&
            .Columns.ColumnByFieldName("STOCKLIBRE").Font.Bold = True
            .Columns.ColumnByFieldName("STOCKLIBRE").FontColor = &HFFFFFF
        
        .Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Color = &H80FFFF
            .Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Font.Bold = False
            .Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").FontColor = &H80000012
        
        .Columns.ColumnByFieldName("STOCK").Color = &HFFFFFF
            .Columns.ColumnByFieldName("STOCK").Font.Bold = True
            .Columns.ColumnByFieldName("STOCK").FontColor = &H80000012
        
        .Columns.ColumnByFieldName("F5FOB").DecimalPlaces = 2
        .Columns.ColumnByFieldName("F5FOBMAX").DecimalPlaces = 2
    End With
End Sub

Private Sub estadoSeleccion(ByVal bolEstado As Boolean)
    On Error GoTo errEstadoSeleccion
    
    dbgProducto.Dataset.Close
    
    Dim dblCantidad As Double
    
    SqlCad = vbNullString
    SqlCad = SqlCad & "UPDATE "
    SqlCad = SqlCad & "TMPPRODUCTOS "
    SqlCad = SqlCad & "SET "
    SqlCad = SqlCad & "F4PERINT = " & IIf(bolEstado, "-1", "0") & " "
    SqlCad = SqlCad & "WHERE "
    SqlCad = SqlCad & "TRIM(F5CODPRO & '') <> '' "
        
        If chkSeleccionActual.Checked Then
            SqlCad = SqlCad & "AND F4PERINT = -1 "
        End If
        
        If chkProductoProveedor.Checked Then
            SqlCad = SqlCad & "AND NOT ISNULL(F5FECUC) "
        End If
        
        If txtBusqueda.Text <> vbNullString Then
            SqlCad = SqlCad & "AND (F5CODPRO LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "F5CODFAB LIKE '%" & txtBusqueda.Text & "%' OR "
            SqlCad = SqlCad & "F5NOMPRO LIKE '%" & txtBusqueda.Text & "%') "
        End If
    
    
    abrirCnTemporal
    
    cnDBTemp.Execute SqlCad, dblCantidad
    
    SqlCad = vbNullString
    
    listarProductos
    
    MsgBox dblCantidad & " item(s) actualizado(s).", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
errEstadoSeleccion:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub chkSeleccionActual_Click()
    listarProductos
End Sub

Private Sub chkProductoProveedor_Click()
    listarProductos
End Sub

Private Sub cmbAlmacen_Click()
    If dbgProducto.Dataset.RecordCount > 0 Then
        dbgProducto.Dataset.Close
        
        Me.MousePointer = vbHourglass
        
        cargarProductos
        
        listarProductos
        
        inicializarControles
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub dbgProducto_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "F4PERINT"
            If dbgProducto.Dataset.State = dsEdit Then
                dbgProducto.Dataset.Post
            End If
    End Select
End Sub

Private Sub dbgProducto_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "STOCKCOMPROMETIDO", "STOCKPORLLEGARCOMP", "STOCKLIBRE", "STOCKPORLLEGARLIBRE", "STOCKTOTAL", "STOCKPORLLEGAR", "STOCK"
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub dbgProducto_OnDblClick()
    Select Case UCase(dbgProducto.Columns.FocusedColumn.FieldName)
        Case "F5CODPRO", "F5CODFAB"
            If Trim(dbgProducto.Columns.ColumnByFieldName("COD_SOLICITUD").Value & "") <> vbNullString Then
                With frmUtilDetalleConsolidadoPedido
                    .NroPedido = dbgProducto.Columns.ColumnByFieldName("COD_SOLICITUD").Value
                    .CodigoProducto = dbgProducto.Columns.ColumnByFieldName("F5CODPRO").Value
                    
                    .Show 1
                End With
            End If
        Case "F5NOMPRO"
            With dbgProducto
                .Dataset.Edit
            
                .Columns.ColumnByFieldName("F4PERINT").Value = IIf(Not CBool(dbgProducto.Columns.ColumnByFieldName("F4PERINT").Value), True, False)
            
                .Dataset.Post
            End With
        Case "STOCKCOMPROMETIDO"
'            If Val(dbgProducto.Columns.ColumnByFieldName("STOCKCOMPROMETIDO").value & "") <= 0 Then
'                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgProducto.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = right(cmbAlmacen.Text, 2)
                
                .Show 1
            End With
        Case "STOCKPORLLEGARCOMP"
            If Val(dbgProducto.Columns.ColumnByFieldName("STOCKPORLLEGARCOMP").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName

                Exit Sub
            End If

            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "C" 'Comprometido
                .CodigoProducto = Trim(dbgProducto.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = right(cmbAlmacen.Text, 2)

                .Show 1
            End With
        Case "STOCKLIBRE"
'            If Val(dbgProducto.Columns.ColumnByFieldName("STOCKLIBRE").value & "") <= 0 Then
'                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName
'
'                Exit Sub
'            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "F" 'Stock Fisico
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgProducto.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = right(cmbAlmacen.Text, 2)
                .left = 0
                
                .Show 1
            End With
        Case "STOCKPORLLEGARLIBRE"
            If Val(dbgProducto.Columns.ColumnByFieldName("STOCKPORLLEGARLIBRE").Value & "") <= 0 Then
                MsgBox "Stock insuficiente.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            With frmUtilStockDetalle
                .TipoNaturaleza = "V" 'Stock Virtual
                .TipoDetalle = "L" 'Libre
                .CodigoProducto = Trim(dbgProducto.Columns.ColumnByFieldName("F5CODPRO").Value & "")
                .CodigoAlmacen = right(cmbAlmacen.Text, 2)
                
                .Show 1
            End With
    End Select
    
    Select Case UCase(dbgProducto.Columns.FocusedColumn.FieldName)
        Case "STOCKCOMPROMETIDO", "STOCKPORLLEGARCOMP", "STOCKLIBRE", "STOCKPORLLEGARLIBRE"
            If frmUtilStockDetalle.RedistribucionEjecutada Then
                cargarProductos
                
                inicializarControles
                
                listarProductos
            End If
    End Select
End Sub

Private Sub dbgProducto_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
'    With dbgProducto.Dataset
'        If dbgProducto.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
'            If .State = dsEdit Then
'                dbgProducto.m.HideEditor
'                .Post
'                .DisableControls
'                .Close
'                .Open
'                .EnableControls
'            End If
'        End If
'    End With
'
'    If dbgProducto.Columns.FocusedColumn.FieldName = "f5fob" Then
'        dbgProducto.Dataset.Edit
'
'        If strCodRequerimiento <> vbNullString Then
'            If Val(dbgProducto.Columns.ColumnByFieldName("F5FOB").Value & "") > Val(dbgProducto.Columns.ColumnByFieldName("F5FOBMAX").Value & "") Then
'                MsgBox "La cantidad no puede exceder al saldo pendiente de atención del requerimiento, verifique.", vbInformation + vbOKOnly, App.ProductName
'
'                dbgProducto.Dataset.Cancel
'
'                Exit Sub
'            End If
'        End If
'
'        dbgProducto.Columns.ColumnByFieldName("F4PERINT").Value = True
'        dbgProducto.Dataset.Post
'    End If
    
    Select Case UCase(dbgProducto.Columns.FocusedColumn.FieldName)
        Case "F5FOB"
            With dbgProducto
                If .Dataset.State = dsEdit Then
                    If strCodRequerimiento <> vbNullString Then
                        If Val(dbgProducto.Columns.ColumnByFieldName("F5FOB").Value & "") > Val(dbgProducto.Columns.ColumnByFieldName("F5FOBMAX").Value & "") Then
                            MsgBox "La cantidad no puede exceder al saldo pendiente de atención del requerimiento, verifique.", vbInformation + vbOKOnly, App.ProductName
                            
                            dbgProducto.Dataset.Cancel
                            
                            Exit Sub
                        End If
                    End If
                    
                    dbgProducto.Dataset.Post
                    
                    .Dataset.Edit
                    .Columns.ColumnByFieldName("F4PERINT").Value = True
                    .Dataset.Post
                End If
            End With
    End Select
End Sub

Private Sub dbgProducto_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            If dbgProducto.Dataset.State = dsEdit Or dbgProducto.Dataset.State = dsInsert Then
                dbgProducto.Dataset.Post
            End If
            
            txtBusqueda.SetFocus
            
            dbgProducto_OnDblClick
        Case vbKeyUp
            txtBusqueda.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    
    'Me.left = 980
    'Me.top = 1200
    
    listarAlmacenEnCombo
    
'    abrirCnTemporal
    
'    With dbgProducto
'        .Options.Unset (egoShowGroupPanel)
'        .Filter.FilterActive = False
'
'        .Dataset.ADODataset.ConnectionString = cnDBTemp
'    End With
    
    cargarProductos
    
    inicializarControles
    
    listarProductos
    
    Me.MousePointer = vbDefault
End Sub

Private Sub inicializarControles()
'    If strCodRequerimiento <> vbNullString And strCodProducto = vbNullString Then
'        Me.Caption = "Productos del Requerimiento N° " & strCodRequerimiento
'    ElseIf strCodRequerimiento = vbNullString And strCodProducto <> vbNullString Then
'        Me.Caption = "Producto Pendiente de Atención en otro(s) Requerimiento(s)"
'    Else
'        Me.Caption = "Ayuda de todos los Productos"
'    End If
    
    If strCodRequerimiento <> vbNullString Then
        Me.Caption = "Productos del Requerimiento N° " & strCodRequerimiento & " - " & ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "CS_OBSERVACIONES", "TB_CABSOLICITUD", "COD_SOLICITUD", strCodRequerimiento, "T")
    ElseIf strCodProducto <> vbNullString Then
        Me.Caption = "Producto Pendiente de Atención en otro(s) Requerimiento(s)"
    Else
        Me.Caption = "Ayuda de todos los Productos"
    End If
    
    With chkProductoProveedor
        .Enabled = False
        .Visible = False
        .Checked = False
        dbgProducto.Columns.ColumnByFieldName("COD_SOLICITUD").Visible = False
        
        If strCodRequerimiento <> vbNullString Or strCodProducto <> vbNullString Then
            .Enabled = True
            .Visible = True
            .Checked = True
            
            If strCodProducto <> vbNullString Then
                .Enabled = False
                .Visible = False
                .Checked = False
                dbgProducto.Columns.ColumnByFieldName("COD_SOLICITUD").Visible = True
            End If
        End If
    End With
    
    'txtBusqueda.Text = vbNullString
    
'    timTemporizador.Enabled = False
'    timTemporizador.Interval = 0
'    pgbProgresoBusqueda.value = 0
'    pgbProgresoBusqueda.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgProducto.Dataset.Close
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    dbgProducto.Move 0, FraBusqueda.Height + 300, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + 300)
End Sub

Private Sub tlbProducto_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Nuevo":
'            sw_nuevo_doc = True
'            sw_mant_ayuda = True
'            mant_productos.Show 1
'            If sw_mant_ayuda = False Then Unload Me
            MsgBox "Opción no disponible.", vbInformation + vbOKOnly, App.ProductName
        Case "ID_Filtrar"
            If Tool.State = ssChecked Then
                dbgProducto.Filter.FilterActive = True
            Else
                dbgProducto.Filter.FilterActive = False
            End If
        Case "ID_Agrupar"
            If Tool.State = ssChecked Then
                dbgProducto.Options.Set (egoShowGroupPanel)
            Else
                dbgProducto.Options.Unset (egoShowGroupPanel)
            End If
        Case "ID_Compras":
'            wcodproducto = dbgProducto.Columns.ColumnByFieldName("f5codpro").Value
'            wdesproducto = dbgProducto.Columns.ColumnByFieldName("f5nompro").Value
'            lista_compras.Show 1
'            Unload Me
            MsgBox "Opción no disponible.", vbInformation + vbOKOnly, App.ProductName
        Case "ID_Salir":
            If dbgProducto.Dataset.State = dsEdit Then
                dbgProducto.Dataset.Post
            End If
            
            dbgProducto.Dataset.Close
            
            Me.Hide
        Case "SeleccionarTodo"
            estadoSeleccion True
        Case "QuitarSeleccion"
            estadoSeleccion False
    End Select
End Sub

Private Sub timTemporizador_Timer()
'    If timTemporizador.Interval = 25 Then
'        listarProductos
'
'        timTemporizador.Enabled = False
'        pgbProgresoBusqueda.value = 0
'        pgbProgresoBusqueda.Visible = False
'    Else
'        timTemporizador.Interval = timTemporizador.Interval + 1
'        pgbProgresoBusqueda.value = timTemporizador.Interval
'    End If
End Sub

Private Sub txtBusqueda_GotFocus()
    ModUtilitario.seleccionarTextoCaja txtBusqueda
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn
            listarProductos
        Case vbKeyDown
            dbgProducto.SetFocus
'        Case vbKeyDelete
'            txtBusqueda_KeyPress KeyCode
    End Select
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 8, 22, 24, 26, 32, 37, 45, 46, 58, 65 To 90, 97 To 122, 209, 40, 41, 46, 48 To 57, 241 - 32
'            timTemporizador.Interval = 0
'            timTemporizador.Enabled = True
'            pgbProgresoBusqueda.value = 0
'            pgbProgresoBusqueda.Visible = True
'
'            timTemporizador_Timer
'        Case Else
'            timTemporizador.Interval = 0
'            timTemporizador.Enabled = False
'            pgbProgresoBusqueda.value = 0
'            pgbProgresoBusqueda.Visible = False
'    End Select
End Sub
