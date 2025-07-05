VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Lista_Documentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Documentos"
   ClientHeight    =   7230
   ClientLeft      =   2925
   ClientTop       =   2190
   ClientWidth     =   7695
   Icon            =   "Lista_Documentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   635
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo  "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir      "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   4500
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
               Picture         =   "Lista_Documentos.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Documentos.frx":05A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Documentos.frx":0B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Lista_Documentos.frx":10DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid Grid 
      Height          =   5865
      Left            =   60
      OleObjectBlob   =   "Lista_Documentos.frx":1674
      TabIndex        =   1
      Top             =   1320
      Width           =   7590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   7605
      Begin VB.TextBox txtbusqueda 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   7140
      End
   End
End
Attribute VB_Name = "Lista_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Rs As New ADODB.Recordset
Private Sub DBGrid1_DblClick()

   DBGrid1_KeyPress 13

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

   DBGrid1_KeyPress KeyCode

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

   Select Case KeyAscii
      Case 13:
         If dc_prov.Recordset.RecordCount > 0 Then
            sw = False
            Cod_Prove = dbgrid1.Columns(0)
            Reg_Documentos.Show 1
            dc_prov.Refresh
            dbgrid1.Refresh
         End If
      Case 27:
         Unload Me
   End Select
 
End Sub

Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)



End Sub

Private Sub Form_Activate()
FILL
End Sub

Private Sub Form_Load()

If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
cnn_dbbancos.Open StrConexDbBancos

FILL
End Sub
Private Sub FILL()

csql = "SELECT F2CODDOC, UCASE(F2DESDOC) AS F2DESDOC, F2ABREV, F2TIPO From DOCUMENTOS ORDER BY F2CODDOC"

If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
With Grid
    .Dataset.Active = False
    .Dataset.ADODataset.ConnectionString = cnn_dbbancos
    .Dataset.ADODataset.CommandText = csql
    .Dataset.Active = True
    .KeyField = "F2CODDOC"
End With
End Sub

Private Sub Grid_OnDblClick()
            sw = False
            Cod_Prove = Grid.Columns(0).Value
            Reg_Documentos.Show 1
           
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Trim(Button.Caption)
        Case "Nuevo":
            sw = True
            Cod_Prove = ""
            Reg_Documentos.Show
            
            
        Case 2:
            If dbgrid1.Row < 0 Then
                MsgBox "Seleccione el documento a modificar.", 48, "CONTROL Plus!"
                dbgrid1.SetFocus
                Exit Sub
            Else
                sw = False
                Cod_Prove = dbgrid1.Columns(0)
                frmregdocs.Show 1
                dc_prov.Refresh
                dbgrid1.Refresh
            End If

        Case "Salir":
            Unload Me
    End Select
End Sub

Private Sub txtcta_GotFocus()

    txtcta.SelStart = 0
    txtcta.SelLength = Len(txtcta.Text)
    
End Sub

Private Sub txtcta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtnumcta.Text = ""
        dc_prov.Recordset.FindFirst "F2CODDOC like " & "'" & Trim(txtcta.Text) & "'"
        dbgrid1.SetFocus
    End If

End Sub

Private Sub txtnumcta_GotFocus()
    
    txtnumcta.SelStart = 0
    txtnumcta.SelLength = Len(txtnumcta.Text)
    
End Sub

Private Sub txtnumcta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtcta.Text = ""
        dc_prov.Recordset.FindFirst "F2DESDOC like " & "'" & "*" & Trim(txtnumcta.Text) & "*" & "'"
        dbgrid1.SetFocus
    End If
    
End Sub

Private Sub txtbusqueda_Change()
    Grid.Dataset.Filtered = True
    Grid.Dataset.Filter = "F2CODDOC LIKE '*" & txtbusqueda.Text & "*' " & _
    "OR " & " F2DESDOC LIKE '*" & txtbusqueda.Text & "*' " & _
    "or " & " F2ABREV like '*" & txtbusqueda.Text & "*' " & _
    "or " & " F2TIPO like '*" & txtbusqueda.Text & "*' "

    If Len(Trim(txtbusqueda.Text)) = 0 Then
            Grid.Dataset.Filtered = False
    End If
End Sub
