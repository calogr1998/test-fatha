VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ayuda_prov_prod 
   Caption         =   "Historial de Precios por Producto"
   ClientHeight    =   7245
   ClientLeft      =   720
   ClientTop       =   2580
   ClientWidth     =   15960
   Icon            =   "ayuda_prov_prod.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   15960
   Begin VB.Frame fraRango 
      Caption         =   " Rango de Historial "
      Height          =   870
      Left            =   8160
      TabIndex        =   3
      Top             =   45
      Width           =   4575
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   3075
         TabIndex        =   7
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   114425857
         CurrentDate     =   41961
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   795
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   114425857
         CurrentDate     =   41961
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   400
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Busqueda "
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   7930
      Begin VB.TextBox txtBusqueda 
         Height          =   345
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   7545
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tlbHistorial 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ayuda_prov_prod.frx":058A
      ToolBars        =   "ayuda_prov_prod.frx":6AB7
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgHistorial 
      Height          =   5775
      Left            =   120
      OleObjectBlob   =   "ayuda_prov_prod.frx":6B86
      TabIndex        =   2
      Top             =   960
      Width           =   15735
   End
End
Attribute VB_Name = "ayuda_prov_prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strCodProducto             As String

Public Property Let CodigoProducto(ByVal value As String)
    strCodProducto = value
End Property

Public Property Get CodigoProducto() As String
    CodigoProducto = strCodProducto
End Property

Public Sub listarHistorial()
    Screen.MousePointer = vbHourglass
    
    dbgHistorial.Dataset.Close
    
    objAyudaVale.listarGrillaHistorialPrecioPorProducto dbgHistorial, Nothing, strCodProducto, dtpDesde.value, dtpHasta.value, wIgv, txtBusqueda.Text
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub dbgHistorial_OnDblClick()
    With objAyudaOrden
        .CodMoneda = Trim(dbgHistorial.Columns.ColumnByFieldName("F4MONEDA").value & "")
        .CodigoUM = Trim(dbgHistorial.Columns.ColumnByFieldName("F7CODMED").value & "")
        .PrecioSinImpuesto = Val(dbgHistorial.Columns.ColumnByFieldName("COSTOUNITARIO").value & "")
        .PorcentajeDscto = Val(dbgHistorial.Columns.ColumnByFieldName("PORCENTAJEDSCTO").value & "")
    End With
    
    Me.Hide
End Sub

Private Sub dtpDesde_CloseUp()
    listarHistorial
End Sub

Private Sub dtpHasta_CloseUp()
    listarHistorial
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    
    Me.left = (Screen.Width / 2) - (Me.Width / 2)
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
    dtpDesde.value = DateSerial(Year(Date), Month(Date) + 0, 1)
    dtpHasta.value = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    listarHistorial
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dbgHistorial.Dataset.Close
End Sub

Private Sub tlbHistorial_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarHistorial
    End Select
End Sub
