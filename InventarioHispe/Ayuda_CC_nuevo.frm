VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ayuda_CC_nuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Centro de Costos"
   ClientHeight    =   8040
   ClientLeft      =   3555
   ClientTop       =   1590
   ClientWidth     =   14430
   ClipControls    =   0   'False
   Icon            =   "Ayuda_CC_nuevo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8610
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   8190
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14430
      _ExtentX        =   25453
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
               Picture         =   "Ayuda_CC_nuevo.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_CC_nuevo.frx":05A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_CC_nuevo.frx":0B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_CC_nuevo.frx":10DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_CC_nuevo.frx":1674
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_CC_nuevo.frx":1C0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Ayuda_CC_nuevo.frx":21A8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11033
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ayuda_CC_nuevo.frx":2742
            Key             =   "Usuario"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ayuda_CC_nuevo.frx":3594
            Key             =   "normal"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ayuda_CC_nuevo.frx":3B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ayuda_CC_nuevo.frx":40C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ayuda_CC_nuevo.frx":4662
            Key             =   "seleccionada"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Ayuda_CC_nuevo.frx":4BFC
            Key             =   ""
         EndProperty
      EndProperty
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
Attribute VB_Name = "Ayuda_CC_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nodX As Node
Private Sub llena_treeview()
Dim rsmov As New ADODB.Recordset

Dim previo(10000) As String
Dim X As Integer, z As Integer, w As Integer
Dim longitud_previa As Integer
Dim wruta As String

On Error GoTo Errores
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cnn_dbbancos.Open "Provider =Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB;Persist Security Info=False"
    If rsmov.State = adStateOpen Then rsmov.Close
    rsmov.Open "Select * From CENTROS ORDER BY LEN(TRIM(F3COSTO)), F3COSTO", cnn_dbbancos, adOpenStatic, adLockOptimistic
    Set nodX = TreeView1.Nodes.Add(, , , "Centros de Costo")
    rsmov.MoveFirst
    Do While Not rsmov.EOF
        If Len(Trim(rsmov.Fields("F3COSTO"))) = 1 Then
            Set nodX = TreeView1.Nodes.Add(1, tvwChild, "a" & rsmov.Fields("F3COSTO"), rsmov.Fields("F3COSTO") & " " & rsmov.Fields("F3DESCRIP"))
'            For z = 1 To X
'                previo(X) = ""
'            Next
'            X = 1
'            previo(X) = Trim(rsmov.Fields("F3COSTO"))
'            X = X + 1
        Else
'            If Len(Trim(previo(X - 1))) <= Len(Trim(rsmov.Fields("F3COSTO"))) Then
'                If previo(X - 1) = Left(rsmov.Fields("F3COSTO"), Len(previo(X - 1))) Then
'                    Set nodX = TreeView1.Nodes.Add("a" & previo(X - 1), tvwChild, "a" & rsmov.Fields("F3COSTO"), rsmov.Fields("F3COSTO") & " " & rsmov.Fields("F3DESCRIP"))
'                    previo(X) = Trim(rsmov.Fields("F3COSTO"))
'                    X = X + 1
'                Else
'                    For w = 1 To X
'                        If Len(Trim(rsmov.Fields("F3COSTO"))) = Len(Trim(previo(X - w))) And Len(previo(X - w - 1)) < Len(Trim(rsmov.Fields("F3COSTO"))) Then Exit For 'And left(Trim(rsmov.Fields("F3COSTO")), longitud_previa) = Trim(previo(X - w))
'                    Next
'                    Set nodX = TreeView1.Nodes.Add("a" & previo(X - w - 1), tvwChild, "a" & rsmov.Fields("F3COSTO"), rsmov.Fields("F3COSTO") & " " & rsmov.Fields("F3DESCRIP"))
'                    previo(X) = Trim(rsmov.Fields("F3COSTO"))
'                    X = X + 1
'                End If
'            Else
'                For w = X To 1 Step -1
'                    If Len(Trim(rsmov.Fields("F3COSTO"))) = Len(Trim(previo(X - w))) Then Exit For
'                Next
                 w = Len(Trim(rsmov.Fields("F3COSTO"))) - 1
                Set nodX = TreeView1.Nodes.Add("a" & Left(rsmov.Fields("F3COSTO"), w), tvwChild, "a" & rsmov.Fields("F3COSTO"), rsmov.Fields("F3COSTO") & " " & rsmov.Fields("F3DESCRIP"))
'                previo(X) = Trim(rsmov.Fields("F3COSTO"))
'                X = X + 1
'            End If
        End If
        If rsmov.EOF Then Exit Do
'        longitud_previa = Len(Trim(rsmov.Fields("F3COSTO")))
        rsmov.MoveNext
    Loop
    nodX.EnsureVisible
    Expandir_Nodo_Seleccionado
    
Exit Sub
Errores:
    If Err.Number = 35601 Then
        w = w - 1
        Resume
    Else
        Resume Next
    End If
    
End Sub
Private Sub Expandir_Nodo_Seleccionado()
Dim d As Integer
    With TreeView1.Nodes
        For d = 1 To .Count
            If .ITEM(d).Index > 1 Then
            .ITEM(d).Expanded = False
            End If
        Next
'        For d = 1 To TreeView1.Nodes.Count
'            If TreeView1.Nodes.ITEM(d).Children <= 1 Then
'                TreeView1.Nodes(d).Expanded = True
'            End If
'        Next
    End With
'   Select Case TreeView1.Nodes.ITEM(d).Children
'    Case Is > 1
'    Node.Expanded = False
'    End Select
End Sub


Private Sub Form_Load()
wruta = "c:\sistemas\contawin\tecom12"
'csql = "Select * From CF5PLA ORDER BY F3COSTO"
'With dxDBGrid1
'    .Dataset.ADODataset.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wruta & "\DB_TABLA.MDB" & ";Persist Security Info=False"
'    .Dataset.Active = False
'    .Dataset.ADODataset.CommandText = csql
'    .Dataset.Active = True
'    .KeyField = "F3COSTO"
'End With

llena_treeview

End Sub

Private Sub Form_Resize()
On Error Resume Next
FraBusqueda.Move 0, 0 + Toolbar.Height, Me.ScaleWidth, 870
txtbusqueda.Width = FraBusqueda.Width - 400
'dxDBGrid1.Move 0, FraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (FraBusqueda.Height + Toolbar.Height)

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
'Case "Filtro"
'    If dxDBGrid1.Filter.FilterActive = True Then
'        dxDBGrid1.Filter.FilterActive = False
'        Me.Toolbar.Buttons.Item(2).Image = 2
'        Me.Toolbar.Buttons.Item(2).ToolTipText = "Activar Filtro"
'    Else
'        dxDBGrid1.Filter.FilterActive = True
'        Me.Toolbar.Buttons.Item(2).Image = 5
'        Me.Toolbar.Buttons.Item(2).ToolTipText = "Desactivar Filtro"
'    End If
'Case "Agrupar"
'    If Button.ToolTipText = "Desagrupar Columnas" Then
'        dxDBGrid1.Options.Unset (egoShowGroupPanel)
'        Me.Toolbar.Buttons.Item(3).Image = 3
'        Me.Toolbar.Buttons.Item(3).ToolTipText = "Agrupar Columnas"
'    Else
'        dxDBGrid1.Options.Set (egoShowGroupPanel)
'        Me.Toolbar.Buttons.Item(3).Image = 6
'        Me.Toolbar.Buttons.Item(3).ToolTipText = "Desagrupar Columnas"
'    End If
Case "Salir"
    Me.Hide
End Select
End Sub

Private Sub TreeView1_DblClick()
    Dim K As Integer
    K = InStr(TreeView1.SelectedItem.Text, Space(1))
    wcodcosto = CStr(Left(TreeView1.SelectedItem.Text, K - 1))
    wdescosto = CStr(Mid(TreeView1.SelectedItem.Text, K + 1, Len(TreeView1.SelectedItem.Text)))
    Me.Hide
    'MsgBox CStr(Left(TreeView1.SelectedItem.Text, K - 1))
End Sub

Private Sub txtbusqueda_Change()
'dxDBGrid1.Dataset.Filtered = True
'dxDBGrid1.Dataset.Filter = "F3COSTO LIKE '*" & txtbusqueda.Text & "*' OR " & " F3DESCRIP LIKE '*" & txtbusqueda.Text & "*' OR " & " F5CCOSTO LIKE '*" & txtbusqueda.Text & "*'"
'
'If Len(Trim(txtbusqueda.Text)) = 0 Then
'    dxDBGrid1.Dataset.Filtered = False
'End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then txtbusqueda_KeyPress 13

End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
