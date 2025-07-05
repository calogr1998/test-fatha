VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ayuda_gastos 
   Caption         =   "Ayuda de Gastos"
   ClientHeight    =   5775
   ClientLeft      =   3405
   ClientTop       =   2700
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ayuda_gastos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9615
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   6810
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   6465
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4545
      Left            =   60
      OleObjectBlob   =   "ayuda_gastos.frx":058A
      TabIndex        =   1
      Top             =   1380
      Width           =   9435
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Todos   "
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtrar   "
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir     "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   4680
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":2D79
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":3313
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":38AD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":3E47
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":43E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":497B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":4F15
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":54AF
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":5A49
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_gastos.frx":5FE3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "ayuda_gastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim csql        As String
Dim wBase As String
Dim StrTipoConcepto As String
Dim StrClaseConcepto As String

Public Property Get ClaseConcepto() As String
    ClaseConcepto = StrClaseConcepto
End Property

Public Property Let ClaseConcepto(ByVal vNewValue As String)
    StrClaseConcepto = vNewValue
End Property


Public Property Get TipoConcepto() As String
    TipoConcepto = StrTipoConcepto
End Property

Public Property Let TipoConcepto(ByVal vNewValue As String)
    StrTipoConcepto = vNewValue
End Property

Private Sub dxDBGrid1_OnDblClick()
    
    wgastos = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value)
    wnomgasto = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("NOMBRE").value)
    wctacont = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("CUENTA").value)
    Me.Hide
    
    
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.M.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With

End Sub

Private Sub FILL(DestinoDeConceptos As String)
Dim StrFiltro As String
If Sw_AyuCodProv = True Then
    StrFiltro = " and TIPO&''='P' "
Else
    'StrFiltro = " and TIPO&''<>'P' "
    StrFiltro = ""
End If

    
csql = "SELECT CODIGO,UCASE(NOMBRE) AS NOMBRE,CUENTA,iif(base='I','Ingresos',iif(base='G','Egresos','Ambos')) as Tipo FROM BF9GIN "
csql = csql & "WHERE BASE like '" & DestinoDeConceptos & "%' " & StrFiltro & " "
csql = csql & "ORDER BY tipo, NOMBRE"
    
    
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.ConnectionString = StrConexDbBancos
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Open
    dxDBGrid1.KeyField = "CODIGO"
    dxDBGrid1.M.FullExpand
Select Case DestinoDeConceptos
Case "I"
    Me.Caption = "Ayuda Conceptos de Ingresos"
Case "G"
    Me.Caption = "Ayuda Conceptos de Egresos"
Case ""
    Me.Caption = "Ayuda de Todos los Conceptos"
End Select

If Sw_AyuCodProv = True Then
    Me.Caption = Me.Caption & " - Tipo Proveedor"
End If

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    
    If KeyCode = 13 Then
        wgastos = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value)
        wnomgasto = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("NOMBRE").value)
        wctacont = Trim("" & dxDBGrid1.Columns.ColumnByFieldName("CUENTA").value)
        Me.Hide
    End If

End Sub

Private Sub Form_Activate()
    
    dxDBGrid1.Option = egoAutoSearch
    dxDBGrid1.OptionEnabled = 0
    
    dxDBGrid1.Columns.FocusedIndex = 2
    
    dxDBGrid1.OptionEnabled = 1
    If TipoConcepto = "E" Then wBase = "G"
    If TipoConcepto = "I" Then wBase = "I"
    If wBase = "I" Then
        Me.Caption = "Ayuda Conceptos de Ingresos"
    Else
        Me.Caption = "Ayuda Conceptos de Egresos"
    End If
    txtbusqueda.Text = ""
    FILL wBase
End Sub

Private Sub Form_Load()
 

    If TipoConcepto = "E" Then wBase = "G"
'    If TipoConcepto = "I" Then wBase = "I"
'    If wBase = "I" Then
'        Me.Caption = "Ayuda Conceptos de Ingresos"
'    Else
'        Me.Caption = "Ayuda Conceptos de Egresos"
'    End If
'
'    fill wBase
                
End Sub

Private Sub Form_Resize()
On Error Resume Next
fraBusqueda.Move 0, 0 + Toolbar.Height, Me.ScaleWidth, 870
txtbusqueda.Width = fraBusqueda.Width - 400
dxDBGrid1.Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
Case "Todos"
    If Button.value = tbrPressed Then
        Button.Image = 9
        dxDBGrid1.Columns.ColumnByFieldName("tipo").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("tipo").GroupIndex = 0
        
        FILL ""
    Else
        Button.Image = 10
        
        dxDBGrid1.Columns.ColumnByFieldName("tipo").GroupIndex = -1
        dxDBGrid1.Columns.ColumnByFieldName("tipo").Visible = False
        FILL wBase
    End If
Case "Filtrar"
    If Button.value = False Then
        dxDBGrid1.Filter.FilterActive = False
        Button.Image = 3
        Button.ToolTipText = "Activar Filtro"
    Else
        dxDBGrid1.Filter.FilterActive = True
        Button.Image = 6
        Button.ToolTipText = "Desactivar Filtro"
    End If

Case "Agrupar"
    If Button.value = False Then
        dxDBGrid1.Options.Unset (egoShowGroupPanel)
        Button.Image = 4
        Button.ToolTipText = "Agrupar Columnas"
    Else
        dxDBGrid1.Options.Set (egoShowGroupPanel)
        Button.Image = 7
        Button.ToolTipText = "Desagrupar Columnas"

    End If
Case "Salir"
    Unload Me
End Select

End Sub

Private Sub txtbusqueda_Change()
    
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "CODIGO LIKE '*" & txtbusqueda.Text & "*' OR " & " NOMBRE LIKE '*" & txtbusqueda.Text & "*' OR " & " CUENTA LIKE '*" & txtbusqueda.Text & "*'"
    
    If Len(Trim(txtbusqueda.Text)) = 0 Then
        dxDBGrid1.Dataset.Filtered = False
    End If
    
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 2
        dxDBGrid1.SetFocus
    End If
    
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) >= 3 Then
            dxDBGrid1.Dataset.Filtered = True
            dxDBGrid1.Dataset.Filter = "CODIGO LIKE '*" & txtbusqueda.Text & "*' OR " & " NOMBRE LIKE '*" & txtbusqueda.Text & "*' OR " & " CUENTA LIKE '*" & txtbusqueda.Text & "*'"
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    End If
    
End Sub
