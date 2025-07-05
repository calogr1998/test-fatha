VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListaCliente 
   Caption         =   "Lista de Clientes"
   ClientHeight    =   7230
   ClientLeft      =   210
   ClientTop       =   1800
   ClientWidth     =   14895
   Icon            =   "frmListaCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   14895
   Begin VB.TextBox txtBusqueda 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   8040
      TabIndex        =   2
      Top             =   0
      Width           =   5100
   End
   Begin MSComDlg.CommonDialog cmdlgCliente 
      Left            =   240
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlbCliente 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imglstCliente"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo  "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList imglstCliente 
         Left            =   5400
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":0B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":1658
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":1BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":218C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":2726
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":2CC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":325A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":37F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListaCliente.frx":3D8E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgCliente 
      Height          =   6705
      Left            =   120
      OleObjectBlob   =   "frmListaCliente.frx":4328
      TabIndex        =   1
      Top             =   480
      Width           =   14655
   End
End
Attribute VB_Name = "frmListaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Af As New ADOFunctions
Private Rs As New ADODB.Recordset

Private bolAyuda        As Boolean

Public Property Let Ayuda(ByVal Value As Boolean)
    bolAyuda = Value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property

Private Sub Form_Activate()
    listarCliente
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width - Me.Width) / 2
    Me.top = (Screen.Height - Me.Height) / 3
    
    listarCliente
    
    tlbCliente.Buttons(3).Visible = Not bolAyuda
    tlbCliente.Buttons(4).Visible = Not bolAyuda
End Sub

Public Sub listarCliente()
    objAyudaCliente.vistaCliente dbgCliente
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    dbgCliente.Dataset.Close
    
    bolAyuda = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With txtBusqueda
        .Width = (tlbCliente.Width / 10) * 4
        .left = (tlbCliente.Width - .Width) - 200
        .top = 25
    End With
    
    dbgCliente.Move 0, tlbCliente.Height, Me.ScaleWidth, Me.ScaleHeight - tlbCliente.Height
End Sub

Private Sub dbgCliente_OnDblClick()
    Me.MousePointer = vbHourglass
    
    If bolAyuda Then
        objAyudaCliente.Codigo = Trim(dbgCliente.Columns.ColumnByFieldName("F2CODCLI").Value & "")
        objAyudaCliente.NombreCliente = Trim(dbgCliente.Columns.ColumnByFieldName("F2NOMCLI").Value & "")
        
        Me.Hide
        
        'Unload Me
    Else
        With frmMantCliente
            .Codigo = Trim(dbgCliente.Columns.ColumnByFieldName("F2CODCLI").Value & "")
            
            .Show 1

            listarCliente
        End With
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub dbgCliente_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyUp
            If dbgCliente.Dataset.RecNo = 1 Then
                txtBusqueda.SetFocus
            End If
        Case vbKeyReturn
            dbgCliente_OnDblClick
    End Select
End Sub

Private Sub dbgCliente_OnKeyPress(Key As Integer)
    Select Case Key
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.Buttons(2)
        Case 9 'Ctrl + G (Imprimir)
            tlbCliente_ButtonClick tlbCliente.Buttons(3)
        Case 5 'Ctrl + E (Excel)
            tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            tlbCliente_ButtonClick tlbCliente.Buttons(6)
    End Select
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Trim(Button.Caption)
        Case "Nuevo"
            With frmMantCliente
                .Ayuda = bolAyuda
                .Codigo = vbNullString
                
                .Show 1
                
                If Not bolAyuda Then
                    listarCliente
                Else
                    Unload Me
                End If
            End With
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgCliente
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                
                .ShowSave
                
                If .FileName <> Empty Then
                    dbgCliente.m.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                Else
                    MsgBox "Exportación cancelada.", vbInformation, App.ProductName
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "Imprimir":
            SqlCad = vbNullString
            SqlCad = SqlCad & "SELECT "
            SqlCad = SqlCad & "EF2CLIENTES.F2NEWRUC, "
            SqlCad = SqlCad & "EF2CLIENTES.F2CODCLI, "
            SqlCad = SqlCad & "EF2CLIENTES.F2NOMCLI, "
            SqlCad = SqlCad & "EF2CLIENTES.F2DIRCLI, "
            SqlCad = SqlCad & "EF2CLIENTES.F2NEWRUC, "
            SqlCad = SqlCad & "EF2CLIENTES.F2TELCLI, "
            SqlCad = SqlCad & "EF2ZONAS_CLIENTE.DESCRIPCION, "
            SqlCad = SqlCad & "EF2ZONAS_CLIENTE.DISTRITO, "
            SqlCad = SqlCad & "EF2VENDEDORES.F2NOMVEN "
            SqlCad = SqlCad & "FROM "
            SqlCad = SqlCad & "(EF2CLIENTES "
            SqlCad = SqlCad & "LEFT JOIN EF2ZONAS_CLIENTE ON EF2CLIENTES.F2ZONA_CLIENTE = EF2ZONAS_CLIENTE.CODIGO) "
            SqlCad = SqlCad & "LEFT JOIN EF2VENDEDORES ON EF2CLIENTES.F2CODVEN = EF2VENDEDORES.F2CODVEN "
            SqlCad = SqlCad & "ORDER BY "
            SqlCad = SqlCad & "EF2ZONAS_CLIENTE.CODIGO, EF2CLIENTES.F2NOMCLI"
            
            With rptImpresionCliente
                .Caption = "Relación de Clientes"
                .DataControl1.ConnectionString = StrConexDbBancos
                .DataControl1.Source = SqlCad
                .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                
                .lblempresa.Caption = wNomCia
                
                .Show 1
            End With
        Case "Salir":
            objAyudaCliente.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    With dbgCliente.Dataset
        .Filtered = True
        .Filter = "F2NEWRUC LIKE '*" & txtBusqueda.Text & "*' " & _
                    "OR F2NOMCLI LIKE '*" & txtBusqueda.Text & "*' " & _
                    "OR F2DIRCLI LIKE '*" & txtBusqueda.Text & "*' " & _
                    "OR ORIGEN LIKE '*" & txtBusqueda.Text & "*' " & _
                    "OR PERSONA LIKE '*" & txtBusqueda.Text & "*' "
        
        If Len(Trim(txtBusqueda.Text)) = 0 Then
            .Filtered = False
        End If
    End With
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dbgCliente.SetFocus
    End Select
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 14 'Ctrl + N (Nuevo)
            tlbCliente_ButtonClick tlbCliente.Buttons(2)
        Case 9 'Ctrl + G (Imprimir)
            tlbCliente_ButtonClick tlbCliente.Buttons(3)
        Case 5 'Ctrl + E (Excel)
            tlbCliente_ButtonClick tlbCliente.Buttons(4)
        Case 19 'Ctrl + S (Salir)
            tlbCliente_ButtonClick tlbCliente.Buttons(6)
    End Select
End Sub
