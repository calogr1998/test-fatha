VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmListaProvDscto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Descuentos"
   ClientHeight    =   6255
   ClientLeft      =   2520
   ClientTop       =   3495
   ClientWidth     =   11895
   Icon            =   "frmListaProvDscto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11895
   Begin VB.Frame fraBusqueda 
      Caption         =   " Buscar "
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10215
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese cadena a buscar"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbProvDscto 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "frmListaProvDscto.frx":058A
      ToolBars        =   "frmListaProvDscto.frx":1EF6
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgProvDscto 
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmListaProvDscto.frx":1F87
      TabIndex        =   0
      Top             =   840
      Width           =   11655
   End
End
Attribute VB_Name = "frmListaProvDscto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolAyuda        As Boolean

Private strCodProv                  As String
Private strCodProd                  As String
Private strCodMed                   As String
Private dblCantidad                 As Double

Public Property Let CodigoProveedor(ByVal Value As String)
    strCodProv = Value
End Property

Public Property Get CodigoProveedor() As String
    CodigoProveedor = strCodProv
End Property

Public Property Let CodigoProducto(ByVal Value As String)
    strCodProd = Value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProd
End Property

Public Property Let CodigoUM(ByVal Value As String)
    strCodMed = Value
End Property

Public Property Get CodigoUM() As String
    CodigoUM = strCodMed
End Property

Public Property Let Cantidad(ByVal Value As Double)
    dblCantidad = Value
End Property

Public Property Get Cantidad() As Double
    Cantidad = dblCantidad
End Property

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property



Private Sub dbgProvDscto_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaProvDscto.Porcentaje = Val(dbgProvDscto.Columns.ColumnByFieldName("Porcentaje").Value & "")
        
        'Me.Hide
        Unload Me
    Else
'        With frmMantBienColor
'            .Ayuda = bolAyuda
'            .Codigo = Trim(dbgProvDscto.Columns.ColumnByFieldName("Codigo").Value & "")
'
'            .Show 1
'
'            'Me.Hide
'
'            listarProvDscto
'        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgProvDscto_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgProvDscto_OnDblClick
        Case vbKeyUp
            If dbgProvDscto.Dataset.RecNo = 1 Then
                txtbusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    If Not bolAyuda Then
        Me.top = 1000
        Me.left = 1250
    End If
    
    With dbgProvDscto
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    End With
    
    listarProvDscto
End Sub

Public Sub listarProvDscto()
    Dim strSQL      As String
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "PORC.F5CODPRO,  "
    strSQL = strSQL & "PORC.F5NOMPRO, "
    strSQL = strSQL & "MED.F7SIGMED, "
    strSQL = strSQL & "PORC.F3PORDCT, "
    strSQL = strSQL & "PORC.F2FECHA "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "EF2PROV_DSCTO AS PORC "
    strSQL = strSQL & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = PORC.F7CODMED "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "NOT ISNULL(PORC.F2FECHA) "
        
        If strCodProv <> vbNullString Then
            strSQL = strSQL & "AND PORC.F2CODPRV = '" & strCodProv & "' "
        End If
        
        If strCodProd <> vbNullString Then
            strSQL = strSQL & "AND PORC.F5CODPRO = '" & strCodProd & "' "
        End If
        
        If strCodMed <> vbNullString Then
            strSQL = strSQL & "AND PORC.F7CODMED = '" & strCodMed & "' "
        End If
        
        If dblCantidad <> 0 Then
            strSQL = strSQL & "AND PORC.F3CANPRO = " & dblCantidad & " "
        End If
    
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "PORC.F2FECHA DESC"
    
    With dbgProvDscto
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "PORCENTAJE"
        
        .Columns.ColumnByFieldName("F5CODPRO").Width = 80
        .Columns.ColumnByFieldName("F5NOMPRO").Width = 200
        .Columns.ColumnByFieldName("F7SIGMED").Width = 30
        .Columns.ColumnByFieldName("F3PORDCT").Width = 60
        .Columns.ColumnByFieldName("F2FECHA").Width = 45
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

Private Sub tlbProvDscto_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "Nuevo"
'            With frmMantBienColor
'                .Ayuda = bolAyuda
'                .Codigo = vbNullString
'
'                .Show 1
'
'                'Me.Hide
'
'                If Not bolAyuda Then
'                    listarProvDscto
'                Else
'                    Unload Me
'                End If
'            End With
        Case "Salir"
            objAyudaProvDscto.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    With dbgProvDscto
        .Dataset.Filtered = True
        .Dataset.Filter = "F5CODPRO LIKE '*" & txtbusqueda.Text & "*' " & _
                            "OR F5NOMPRO LIKE '*" & txtbusqueda.Text & "*'"
        
        If Trim(txtbusqueda.Text) = vbNullString Then
            .Dataset.Filtered = False
        End If
    End With
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dbgProvDscto.SetFocus
    End Select
End Sub

