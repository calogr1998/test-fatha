VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ayuda_Productos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de Productos con Stock"
   ClientHeight    =   7785
   ClientLeft      =   1980
   ClientTop       =   3135
   ClientWidth     =   17550
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   17550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6420
      Left            =   60
      OleObjectBlob   =   "ayuda_productos.frx":058A
      TabIndex        =   2
      Top             =   1260
      Width           =   17265
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   17550
      _ExtentX        =   30956
      _ExtentY        =   635
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo   "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro     "
            Object.ToolTipText     =   "Activar Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agrupar"
            Object.ToolTipText     =   "Agrupar Columnas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compras"
            Object.ToolTipText     =   "Listar Ultimas Compras"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Picture         =   "ayuda_productos.frx":5444
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_productos.frx":59DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_productos.frx":5F78
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_productos.frx":6512
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_productos.frx":6AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_productos.frx":7046
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ayuda_productos.frx":75E0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7065
      Begin VB.TextBox txtbusqueda 
         Height          =   315
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   6645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   7200
      TabIndex        =   4
      Top             =   360
      Width           =   3555
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Incluir Productos sin Stock"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         TabIndex        =   6
         Top             =   180
         Width           =   2265
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar todos los productos"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         TabIndex        =   5
         Top             =   480
         Width           =   2625
      End
   End
End
Attribute VB_Name = "Ayuda_Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_mov     As New ADODB.Connection
Dim csql        As String
Dim Estado      As Boolean

Private Sub Check1_Click()
        If Check1.Value = 1 Then
            Me.Caption = "Ayuda de todos los Productos con movimiento"
            Check2.Value = 0
            FILL
        Else
            Me.Caption = "Ayuda de Productos con Stock"
            FILL
        End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Me.Caption = "Ayuda de todos los Productos"
    Check1.Value = 0
    FILL
Else
    Me.Caption = "Ayuda de Productos con Stock"
    FILL
End If
End Sub

Private Sub Checkagrupar_Click()
    If Checkagrupar.Value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.Value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub dxDBGrid1_OnDblClick()
    wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").Value
    wcodfab = dxDBGrid1.Columns.ColumnByFieldName("f5codfab").Value
    wmarca = dxDBGrid1.Columns.ColumnByFieldName("f2desmar").Value
    wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value
    wmedida = dxDBGrid1.Columns.ColumnByFieldName("f7codmed").Value
    wstockact = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f6stockact").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("f6stockact").Value)
    wprecos = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanet").Value)
    wprecosdol = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("f5vtanetdol").Value)
    wafecto = dxDBGrid1.Columns.ColumnByFieldName("f5afecto").Value
    wtipocc = IIf(IsNull(dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").Value), 0, dxDBGrid1.Columns.ColumnByFieldName("F5ULTTC").Value)
    Me.Hide
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
    End With

End Sub

Private Sub FILL()

 If Check2.Value = 1 Then
        csql = "SELECT DISTINCT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC, IF5PLA.F5VTANET, IF5PLA.F5VTANETDOL"
        csql = csql + " FROM ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '')]. AS Consulta3 RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) LEFT JOIN [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO"
        csql = csql + " GROUP BY Consulta3.CANTIDAD, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, IF5PLA.F5VALVTA, IF5PLA.F5FOB, IF5PLA.F5AFECTO, IF5PLA.F5FECUC, IF5PLA.F5ULTTC, IF5PLA.F5VTANET, IF5PLA.F5VTANETDOL, IF5PLA.F5STOCKACT"
        csql = csql + " ORDER BY IF5PLA.F5NOMPRO;"

 Else
    If wcod_alm = "" Then wcod_alm = "01"
        csql = ""
        csql = "SELECT Consulta3.CANTIDAD AS F6STOCKACT, IF5PLA.*, EF2MARCAS.F2DESMAR, Consulta2.F2CODALM"
        csql = csql + " FROM (EF2MARCAS INNER JOIN ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD FROM IF3VALES GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '" & wcod_alm & "')]. AS Consulta3"
        csql = csql + " RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) ON EF2MARCAS.F2CODMAR = IF5PLA.F5MARCA)"
        csql = csql + " INNER JOIN  [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES GROUP BY IF3VALES.F2CODALM, IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO"
        csql = csql + " WHERE (Consulta2.F2CODALM)='" & wcod_alm & "'"
        If Check1.Value <> 1 Then
            csql = csql + " and NOT(ISNULL(Consulta3.CANTIDAD))"
        End If
        csql = csql & " ORDER BY F5NOMPRO;"
        
 End If
    dxDBGrid1.Dataset.Active = False
    dxDBGrid1.Dataset.ADODataset.CommandText = csql
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "F5CODPRO"
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 13 Then
        dxDBGrid1_OnDblClick
    End If
End Sub

Private Sub Form_Load()
    Me.MousePointer = 11
    
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False

    If cnn_mov.State = adStateOpen Then cnn_mov.Close
    cnn_mov.ConnectionString = cnn_dbbancos
    cnn_mov.Open cconexion
    
    With dxDBGrid1
        .Dataset.ADODataset.ConnectionString = cnn_mov
    End With
Check2.Value = 1
Check2_Click

Me.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

    dxDBGrid1.Dataset.Close
    cnn_mov.Close
    Set Ayuda_Productos = Nothing
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            sw_nuevo_doc = True
            sw_mant_ayuda = True
            mant_productos.Show 1
            If sw_mant_ayuda = False Then Unload Me
        Case "ID_Compras":
            wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").Value
            wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value
            lista_compras.Show 1
            Unload Me
        Case "ID_Salir":
            Unload Me
    End Select

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
Case "Nuevo":
    sw_nuevo_doc = True
    sw_mant_ayuda = True
    mant_productos.Show 1
    If sw_mant_ayuda = False Then Unload Me
Case "Compras":
    wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("f5codpro").Value
    wdesproducto = dxDBGrid1.Columns.ColumnByFieldName("f5nompro").Value
    lista_compras.Show 1
Case "Filtro"
    If dxDBGrid1.Filter.FilterActive = True Then
        dxDBGrid1.Filter.FilterActive = False
        Me.Toolbar.Buttons.ITEM(4).Image = 2
        Me.Toolbar.Buttons.ITEM(4).ToolTipText = "Activar Filtro"
    Else
        dxDBGrid1.Filter.FilterActive = True
        Me.Toolbar.Buttons.ITEM(4).Image = 5
        Me.Toolbar.Buttons.ITEM(4).ToolTipText = "Desactivar Filtro"
    End If
Case "Agrupar"
    If Button.ToolTipText = "Desagrupar Columnas" Then
        dxDBGrid1.Options.Unset (egoShowGroupPanel)
        Me.Toolbar.Buttons.ITEM(5).Image = 3
        Me.Toolbar.Buttons.ITEM(5).ToolTipText = "Agrupar Columnas"
    Else
        dxDBGrid1.Options.Set (egoShowGroupPanel)
        Me.Toolbar.Buttons.ITEM(5).Image = 6
        Me.Toolbar.Buttons.ITEM(5).ToolTipText = "Desagrupar Columnas"
    End If
Case "Salir"
    Unload Me
End Select

End Sub

Private Sub txtbusqueda_Change()
'    dxDBGrid1.Dataset.Filtered = True
'    dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtbusqueda.Text & "*' " & _
'    "OR " & " F5CODFAB LIKE '*" & txtbusqueda.Text & "*' " & _
'    "or " & " F2DESMAR like '*" & txtbusqueda.Text & "*' " & _
'    "or " & " F5NOMPRO like '*" & txtbusqueda.Text & "*' " & _
'    "or " & " F7CODMED like '*" & txtbusqueda.Text & "*' "
''    "or " & " F5MARCA  like '*" & txtbusqueda.Text & "*' "
'
'    If Len(Trim(txtbusqueda.Text)) = 0 Then
'            dxDBGrid1.Dataset.Filtered = False
'    End If
End Sub

Private Sub txtbusqueda_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
        If Len(Trim(txtbusqueda.Text)) > 0 Then
            dxDBGrid1.Dataset.Filtered = True

        dxDBGrid1.Dataset.Filter = "F5CODPRO LIKE '*" & txtbusqueda.Text & "*' " & _
        "or " & " F5NOMPRO like '*" & txtbusqueda.Text & "*' "
'        "or " & " F5MARCA  like '*" & txtbusqueda.Text & "*' "
        Else
            dxDBGrid1.Dataset.Filtered = False
        End If
    'End If
    
End Sub



