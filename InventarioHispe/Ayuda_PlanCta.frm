VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ayuda_PlanCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda del Plan de Cuentas"
   ClientHeight    =   5955
   ClientLeft      =   3555
   ClientTop       =   1590
   ClientWidth     =   8760
   ClipControls    =   0   'False
   Icon            =   "Ayuda_PlanCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8610
      Begin VB.TextBox txtbusqueda 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   8190
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   4665
      Left            =   0
      OleObjectBlob   =   "Ayuda_PlanCta.frx":058A
      TabIndex        =   0
      Top             =   1380
      Width           =   8610
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro     "
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar"
            Object.ToolTipText     =   "Agrupar Columnas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir      "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   8460
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":290A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":2EA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":343E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":39D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":3F72
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":450C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_PlanCta.frx":4AA6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu MnuPri 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuFiltro 
         Caption         =   "Filtrar"
      End
      Begin VB.Menu MnuFiltroavanz 
         Caption         =   "Filtro Avanzado:"
      End
      Begin VB.Menu MnuOrdAsc 
         Caption         =   "Ord. Asc"
      End
      Begin VB.Menu MnuOrdDesc 
         Caption         =   "Ord. Desc"
      End
      Begin VB.Menu MnuTodo 
         Caption         =   "Mostrar Todos"
      End
   End
End
Attribute VB_Name = "Ayuda_PlanCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dxDBGrid1_OnDblClick()
dxDBGrid1_OnKeyDown 13, 0
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
Select Case KeyCode
   Case 13:
      wctacont = dxDBGrid1.Columns.ColumnByFieldName("F5CODCTA").value & ""
      wnomctacont = dxDBGrid1.Columns.ColumnByFieldName("F5NOMCTA").value & ""
      Me.Hide
   Case 27:
      wctacont = ""
      wnomctacont = ""
      Me.Hide
End Select
End Sub

Private Sub Form_Load()
 

csql = "Select * From CF5PLA ORDER BY F5CODCTA"
With dxDBGrid1
    .Dataset.ADODataset.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\DB_TABLA.MDB" & ";Persist Security Info=False"
    .Dataset.Active = False
    .Dataset.ADODataset.CommandText = csql
    .Dataset.Active = True
    .KeyField = "f5codcta"
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
fraBusqueda.Move 0, 0 + Toolbar.Height, Me.ScaleWidth, 870
txtBusqueda.Width = fraBusqueda.Width - 400
dxDBGrid1.Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
Case "Filtro"
    If dxDBGrid1.Filter.FilterActive = True Then
        dxDBGrid1.Filter.FilterActive = False
        Me.Toolbar.Buttons.ITEM(2).Image = 2
        Me.Toolbar.Buttons.ITEM(2).ToolTipText = "Activar Filtro"
    Else
        dxDBGrid1.Filter.FilterActive = True
        Me.Toolbar.Buttons.ITEM(2).Image = 5
        Me.Toolbar.Buttons.ITEM(2).ToolTipText = "Desactivar Filtro"
    End If
Case "Agrupar"
    If Button.ToolTipText = "Desagrupar Columnas" Then
        dxDBGrid1.Options.Unset (egoShowGroupPanel)
        Me.Toolbar.Buttons.ITEM(3).Image = 3
        Me.Toolbar.Buttons.ITEM(3).ToolTipText = "Agrupar Columnas"
    Else
        dxDBGrid1.Options.Set (egoShowGroupPanel)
        Me.Toolbar.Buttons.ITEM(3).Image = 6
        Me.Toolbar.Buttons.ITEM(3).ToolTipText = "Desagrupar Columnas"
    End If
Case "Salir"
    Me.Hide
End Select
End Sub

Private Sub txtbusqueda_Change()
dxDBGrid1.Dataset.Filtered = True
dxDBGrid1.Dataset.Filter = "F5CODCTA LIKE '*" & txtBusqueda.Text & "*' OR " & " F5NOMCTA LIKE '*" & txtBusqueda.Text & "*' OR " & " F5CCOSTO LIKE '*" & txtBusqueda.Text & "*'"

If Len(Trim(txtBusqueda.Text)) = 0 Then
    dxDBGrid1.Dataset.Filtered = False
End If
End Sub
