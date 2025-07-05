VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Lista_Centros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Centros de Costo"
   ClientHeight    =   6270
   ClientLeft      =   3345
   ClientTop       =   2100
   ClientWidth     =   9480
   Icon            =   "Lista_Centros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerSlide 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4140
      Top             =   3480
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      ButtonWidth     =   1757
      ButtonHeight    =   487
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo   "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir "
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro      "
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar "
            Object.ToolTipText     =   "Agrupar Columnas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Atrás      "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Siguiente"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir        "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   8520
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
               Picture         =   "Lista_Centros.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":0B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":1658
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":1BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":218C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":2726
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":2CC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":325A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Centros.frx":37F4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   8745
      Begin VB.TextBox txtbusqueda 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8280
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   4605
      Index           =   1
      Left            =   360
      OleObjectBlob   =   "Lista_Centros.frx":3D8E
      TabIndex        =   3
      Top             =   1620
      Width           =   8730
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   4605
      Index           =   2
      Left            =   0
      OleObjectBlob   =   "Lista_Centros.frx":5CE1
      TabIndex        =   4
      Top             =   0
      Width           =   8730
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   4605
      Index           =   3
      Left            =   0
      OleObjectBlob   =   "Lista_Centros.frx":7C34
      TabIndex        =   5
      Top             =   0
      Width           =   8730
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   4605
      Index           =   4
      Left            =   0
      OleObjectBlob   =   "Lista_Centros.frx":9B87
      TabIndex        =   6
      Top             =   0
      Width           =   8730
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
      Height          =   4605
      Index           =   5
      Left            =   0
      OleObjectBlob   =   "Lista_Centros.frx":BADA
      TabIndex        =   7
      Top             =   0
      Width           =   8730
   End
End
Attribute VB_Name = "Lista_Centros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SwFilter As Boolean
Dim strDireccion As String
Dim rstconsulta As New ADODB.Recordset
Dim wgraba          As String
Dim wtippro         As String



Private Sub dxDBGrid_OnChangeNode(Index As Integer, ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
If nGridActive = Index Then
    proceso dxDBGrid(nGridActive).Columns.ColumnByName("DxCodigo").value & ""
End If
End Sub

Private Sub dxDBGrid_OnClick(Index As Integer)
If nGridActive = Index Then
    proceso dxDBGrid(nGridActive).Columns.ColumnByName("DxCodigo").value & ""
End If
End Sub

Private Sub dxDBGrid_OnDblClick(Index As Integer)
Me.MousePointer = vbHourglass
FrmName = Me.Name
sw_nuevo_doc = False
Reg_Centros.Show 1
FILL
Me.MousePointer = 0
End Sub

Private Sub dxDBGrid_OnKeyDown(Index As Integer, KeyCode As Integer, ByVal Shift As Long)
dxDBGrid_OnClick (nGridActive)
Select Case KeyCode
Case 13
    dxDBGrid_OnDblClick (nGridActive)
Case 39
    If dxDBGrid(nGridActive).Columns.FocusedIndex = (dxDBGrid(nGridActive).Columns.Count - 1) And nGridActive < 5 Then
        strDireccion = "D"
        TimerSlide.Enabled = True
    End If
Case 37
    If dxDBGrid(nGridActive).Columns.FocusedIndex = 0 And nGridActive > 1 Then
        strDireccion = "I"
        TimerSlide.Enabled = True
    End If
End Select
End Sub

Private Sub Form_Load()
Dim csql    As String
Dim i       As Integer

nGridActive = 1

Me.MousePointer = vbHourglass
Me.left = 0
Me.top = 0


FillAll

Me.MousePointer = vbDefault
 End Sub

Private Sub Form_Resize()
On Error Resume Next
fraBusqueda.Move 0, 0 + Toolbar.Height, Me.ScaleWidth, 870
txtbusqueda.Width = fraBusqueda.Width - 400
dxDBGrid(0).Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)
dxDBGrid(1).Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)
dxDBGrid(2).Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)
dxDBGrid(3).Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)
dxDBGrid(4).Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)
dxDBGrid(5).Move 0, fraBusqueda.Height + Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - (fraBusqueda.Height + Toolbar.Height)

For i = 0 To 5
   ' dxDBGrid(i).Visible = False
    If i = nGridActive Then
        dxDBGrid(i).Visible = True
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dxDBGrid(1).Dataset.Close
    dxDBGrid(2).Dataset.Close
    dxDBGrid(3).Dataset.Close
    dxDBGrid(4).Dataset.Close
    dxDBGrid(5).Dataset.Close
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim csql As String

Select Case Tool.Id
    Case "ID_Nuevo":
        Me.MousePointer = vbHourglass
        sw_nuevo_doc = True
        frmcentros.Show 1
        Me.MousePointer = 0
    Case "ID_Imprimir"
        Me.MousePointer = vbHourglass
        RptCentroSxFecha.Show 1
        Me.MousePointer = 0
    Case "ID_Salir"
        Unload Me
End Select
    
End Sub

Private Sub FillAll()
Dim csql As String

For i = 1 To 5

    csql = "SELECT F3COSTO,F3ABREV,F3DESCRIP,PO, F3FECGRA "
    csql = csql & "From CENTROS "
    csql = csql & "WHERE F3COSTO<>'999' "
    csql = csql & "AND F3COSTO <>'998' "
    csql = csql & "and intCodigoNivel=" & i & " "
    csql = csql & "ORDER BY F3COSTO DESC;"
    
    With dxDBGrid(i)
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = cconex_dbbancos
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "F3COSTO"
    End With

Next
Me.Caption = "Lista de Centros - Nivel: " & nGridActive
End Sub

Private Sub FILL()
Dim csql As String

    csql = "SELECT F3COSTO,F3ABREV,F3DESCRIP,PO, F3FECGRA "
    csql = csql & "From CENTROS "
    csql = csql & "WHERE F3COSTO<>'999' "
    csql = csql & "AND F3COSTO <>'998' "
    csql = csql & "and intCodigoNivel=" & nGridActive & " "
    csql = csql & "ORDER BY F3COSTO DESC;"
    
    With dxDBGrid(nGridActive)
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = StrConexDbBancos
        .Dataset.ADODataset.CommandText = csql
        .Dataset.Active = True
        .KeyField = "F3COSTO"
    End With

End Sub

Private Sub TimerSlide_Timer()
On Error GoTo CapturaError
    Select Case strDireccion
    Case "D"
        If Not dxDBGrid(nGridActive).Width + dxDBGrid(nGridActive).left <= 0 Then
            dxDBGrid(nGridActive + 1).left = dxDBGrid(nGridActive).left + dxDBGrid(nGridActive).Width
            dxDBGrid(nGridActive).left = dxDBGrid(nGridActive).left - 150
            dxDBGrid(nGridActive + 1).left = dxDBGrid(nGridActive).left + dxDBGrid(nGridActive).Width
        Else
            TimerSlide.Enabled = False
            nGridActive = nGridActive + 1
            Me.Caption = "Lista de Centros - Nivel: " & nGridActive
            If nGridActive = 5 Then
                Toolbar.Buttons(8).Visible = True
                Toolbar.Buttons(9).Visible = False
            Else
                Toolbar.Buttons(8).Visible = True
                Toolbar.Buttons(9).Visible = True
            End If
            dxDBGrid(nGridActive).left = 0
            dxDBGrid(nGridActive).SetFocus
        End If
    Case "I"
        If Not dxDBGrid(nGridActive).Width - dxDBGrid(nGridActive).left <= 0 Then
            dxDBGrid(nGridActive - 1).left = dxDBGrid(nGridActive).left - dxDBGrid(nGridActive).Width
            dxDBGrid(nGridActive).left = dxDBGrid(nGridActive).left + 150
            dxDBGrid(nGridActive - 1).left = dxDBGrid(nGridActive).left - dxDBGrid(nGridActive).Width
        Else
            TimerSlide.Enabled = False
            nGridActive = nGridActive - 1
            Me.Caption = "Lista de Centros - Nivel: " & nGridActive
            If nGridActive = 1 Then
                Toolbar.Buttons(8).Visible = False
                Toolbar.Buttons(9).Visible = True
            Else
                Toolbar.Buttons(8).Visible = True
                Toolbar.Buttons(9).Visible = True
            End If
            dxDBGrid(nGridActive).left = 0
            dxDBGrid(nGridActive).SetFocus
        End If
    
    End Select
    Exit Sub
CapturaError:
    TimerSlide.Enabled = False
    MsgBox "No hay mas niveles programados.", vbCritical, wnomcia
    
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim csql As String

Select Case Trim(Button.Caption)
    Case "Nuevo":
        Me.MousePointer = vbHourglass
        sw_nuevo_doc = True
        wcodcosto = left(dxDBGrid(nGridActive).Bands.BandByName("Band").Caption & "", (nGridActive - 1) * 3)
        Reg_Centros.Show 1
        FILL
        Me.MousePointer = 0
    Case "Imprimir"
        Me.MousePointer = vbHourglass
        RptCentroSxFecha.Show 1
        Me.MousePointer = 0
    Case "Filtro"
        If dxDBGrid1.Filter.FilterActive = True Then
            dxDBGrid1.Filter.FilterActive = False
            Me.Toolbar.Buttons.ITEM(5).Image = 5
            Me.Toolbar.Buttons.ITEM(5).ToolTipText = "Activar Filtro"
        Else
            dxDBGrid1.Filter.FilterActive = True
            Me.Toolbar.Buttons.ITEM(5).Image = 7
            Me.Toolbar.Buttons.ITEM(5).ToolTipText = "Desactivar Filtro"
        End If
    Case "Agrupar"
        If Button.ToolTipText = "Desagrupar Columnas" Then
            dxDBGrid1.Options.Unset (egoShowGroupPanel)
            Me.Toolbar.Buttons.ITEM(6).Image = 3
            Me.Toolbar.Buttons.ITEM(6).ToolTipText = "Agrupar Columnas"
        Else
            dxDBGrid1.Options.Set (egoShowGroupPanel)
            Me.Toolbar.Buttons.ITEM(6).Image = 6
            Me.Toolbar.Buttons.ITEM(6).ToolTipText = "Desagrupar Columnas"
        End If
    Case "Atrás"
        dxDBGrid_OnClick (nGridActive)
        strDireccion = "I"
        TimerSlide.Enabled = True
    Case "Siguiente"
        dxDBGrid_OnClick (nGridActive)
        strDireccion = "D"
        TimerSlide.Enabled = True
    Case "Salir"
        Unload Me
End Select

End Sub

Private Sub txtbusqueda_Change()
dxDBGrid(nGridActive).Dataset.Filtered = True
dxDBGrid(nGridActive).Dataset.Filter = "F3COSTO LIKE '*" & txtbusqueda.Text & "*' " & _
                        "OR " & " F3DESCRIP LIKE '*" & txtbusqueda.Text & "*' " & _
                        "OR " & " F3ABREV LIKE '*" & txtbusqueda.Text & "*' OR " & " PO LIKE '*" & txtbusqueda.Text & "*' "


If Len(Trim(txtbusqueda.Text)) = 0 Then
        dxDBGrid(nGridActive).Dataset.Filtered = False
End If
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        dxDBGrid(nGridActive).Columns.FocusedIndex = 1
        dxDBGrid(nGridActive).SetFocus
    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbusqueda_Change
End If
End Sub

Private Sub proceso(ByVal centro As String)
On Error Resume Next
    

    MousePointer = 11
    dxDBGrid(nGridActive + 1).Bands.BandByName("Band").Caption = dxDBGrid(nGridActive).Columns.ColumnByName("dxCodigo").value & " - " & dxDBGrid(nGridActive).Columns.ColumnByName("dxDescripcion").value
    dxDBGrid(nGridActive + 1).Dataset.Filtered = True
    dxDBGrid(nGridActive + 1).Dataset.Filter = "f3costo like '" & centro & "*'"
    If dxDBGrid(nGridActive + 1).Bands.BandByName("Band").Caption = " - " Then
        dxDBGrid(nGridActive + 1).Bands.BandByName("Band").Caption = ""
    End If
    MousePointer = 0

End Sub
