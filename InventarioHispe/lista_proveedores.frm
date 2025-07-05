VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Lista_Proveedores 
   Caption         =   "Lista de Proveedores"
   ClientHeight    =   6870
   ClientLeft      =   2235
   ClientTop       =   2835
   ClientWidth     =   10020
   Icon            =   "lista_proveedores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   10020
   Begin VB.Frame fraProceso 
      Caption         =   " Procesando "
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   5535
      Begin ComctlLib.ProgressBar pgbProceso 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   5595
      Left            =   60
      OleObjectBlob   =   "lista_proveedores.frx":058A
      TabIndex        =   1
      Top             =   1260
      Width           =   9885
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Nuevo  "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   6240
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_proveedores.frx":3273
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_proveedores.frx":380D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_proveedores.frx":3DA7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "lista_proveedores.frx":4341
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   9885
      Begin VB.TextBox txtbusqueda 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9660
      End
   End
End
Attribute VB_Name = "Lista_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub FILL()
Dim csql As String

'csql = "select F2CODPROV,F2NOMPROV,F2NEWRUC,f2telprov from EF2PROVEEDORES order by F2CODPROV"

csql = "SELECT EF2PROVEEDORES.F2CODPROV, ucase(EF2PROVEEDORES.F2NOMPROV) as F2NOMPROV, EF2PROVEEDORES.F2NEWRUC, EF2PROVEEDORES.F2TELPROV, EF2FORPAG.F2DESPAG "
csql = csql & "FROM EF2PROVEEDORES LEFT JOIN EF2FORPAG ON EF2PROVEEDORES.F2FORPAG = EF2FORPAG.F2FORPAG "
csql = csql & "ORDER BY EF2PROVEEDORES.F2CODPROV"

dxDBGrid1.Dataset.Active = False
dxDBGrid1.Dataset.ADODataset.ConnectionString = StrConexDbBancos
dxDBGrid1.Dataset.ADODataset.CommandText = csql
dxDBGrid1.Dataset.Active = True
dxDBGrid1.KeyField = "f2newruc"
End Sub

Private Sub dxDBGrid1_OnDblClick()
dxDBGrid1_OnKeyDown 13, 0
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
   Case 13:
      If dxDBGrid1.Count > 0 Then
         sw_nuevo_documento = False
         Me.MousePointer = vbHourglass

         Cod_Prove = dxDBGrid1.Columns.ColumnByFieldName("f2codprov").value & ""
         FrmName = UCase(Me.Name)
         Me.Hide
         Mant_Proveedores.Show 1
         Unload Mant_Proveedores
         Set Mant_Proveedores = Nothing
         FILL
         Me.MousePointer = 0
      End If
   Case 27:
      Unload Me
End Select
End Sub

Private Sub Form_Activate()
    wgraba = "0"
     FILL
End Sub




Private Sub Form_Load()
 
Me.left = 0
Me.top = 0

 
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
        Case "Nuevo":
            Me.MousePointer = vbHourglass
            sw = True
            sw_nuevo_documento = True
            Cod_Prove = ""
            FrmName = UCase(Me.Name)
            Me.Hide
            Mant_Proveedores.Show 1
            FILL
            Me.MousePointer = 0
        Case "Imprimir"
            
            Dim X As Object
        
            Set X = New acr_proveedores
        
            With X
                .Caption = "Relación de Proveedores"
                 .DataControl1.ConnectionString = StrConexDbBancos
                 .DataControl1.Source = "SELECT * FROM EF2PROVEEDORES ORDER BY F2NOMPROV"
                 .fldFecha.Text = Format(Date, "DD/MM/YYYY")
                 .lblEmpresa.Caption = wnomcia
            End With
            If Not X Is Nothing Then
                Load ReporteChildTrue
                ReporteChildTrue.Caption = "Relación de Proveedores"
                Set ReporteChildTrue.arvPreview.object = X
                ReporteChildTrue.Show
            End If
        Case "Importar"
            'Importa_Proveedores.Show 1
            Me.MousePointer = vbHourglass
            
            dxDBGrid1.Dataset.Close
            
            ModMilano.importarPersonasServidorExterno fraProceso, pgbProceso
            
            FILL
            
            Me.MousePointer = vbDefault
        Case "Salir"
            Unload Me
    End Select
End Sub

Private Sub txtbusqueda_Change()
    dxDBGrid1.Dataset.Filtered = True
    dxDBGrid1.Dataset.Filter = "F2CODPROV LIKE '*" & txtbusqueda.Text & "*' " & _
    "OR " & " F2NEWRUC LIKE '*" & txtbusqueda.Text & "*' OR " & " F2NOMPROV LIKE '*" & txtbusqueda.Text & "*' "


    If Len(Trim(txtbusqueda.Text)) = 0 Then
            dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid1.Columns.FocusedIndex = 1
        dxDBGrid1.SetFocus
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
    End If
End Sub
