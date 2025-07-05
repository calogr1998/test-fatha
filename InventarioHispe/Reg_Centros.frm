VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Reg_Centros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centros de Costos"
   ClientHeight    =   2880
   ClientLeft      =   3420
   ClientTop       =   1695
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Reg_Centros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   3795
      Begin VB.OptionButton OptEst 
         Caption         =   "Anulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   16
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton OptEst 
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox pnlcliente 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   8
      Top             =   2460
      Width           =   4335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   2235
      Begin VB.TextBox TxtCodigo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   60
         MaxLength       =   255
         TabIndex        =   7
         Top             =   240
         Width           =   2070
      End
   End
   Begin VB.TextBox TxtPoX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4620
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2100
      Width           =   1530
   End
   Begin VB.TextBox txtunidad 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   960
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtcliente 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   960
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2460
      Width           =   870
   End
   Begin VB.TextBox txtdescrip 
      Height          =   300
      Left            =   975
      TabIndex        =   1
      Top             =   1380
      Width           =   5190
   End
   Begin VB.TextBox TxtDireccion 
      Height          =   300
      Left            =   975
      TabIndex        =   2
      Top             =   1740
      Width           =   5190
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   635
      ButtonWidth     =   2011
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
            Object.Visible         =   0   'False
            Caption         =   "Nuevo   "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar   "
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar  "
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir        "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
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
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Centros.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Centros.frx":05A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Centros.frx":0B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Centros.frx":10DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Centros.frx":1674
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reg_Centros.frx":1C0E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label 
      Caption         =   "Abreviatura"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label 
      Caption         =   "Orden cliente"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label 
      Caption         =   "Cliente"
      Height          =   195
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Dirección"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   1035
   End
End
Attribute VB_Name = "Reg_Centros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Dim Af As New ADOFunctions
Dim RsDemo  As New ADODB.Recordset
Dim Codigo As String

Private Sub Ayuda1_Click()

End Sub

Private Sub Form_Activate()
Me.SetFocus
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Trim(Button.Caption)
        Case "Nuevo"
            sw_nuevo_doc = True
            nuevo_centro
'            txtunidad.SetFocus
        Case "Grabar"
        
'        If Len(Trim(txtcodigo.Text)) = 0 Then
'            MsgBox "Datos Incompletos", vbInformation, "Atención"
'            Exit Sub
'        End If
        If Len(Trim(txtunidad.Text)) = 0 Then
            MsgBox "Datos Incompletos", vbInformation, wnomcia
            txtunidad.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtdescrip.Text)) = 0 Then
            MsgBox "Datos Incompletos", vbInformation, wnomcia
            txtdescrip.SetFocus
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        grabar_centro
        Me.MousePointer = 0
        wcodcosto = TxtCodigo.Text & ""
        wdescosto = txtdescrip.Text & ""
        wunicosto = txtunidad.Text & ""
        wclicosto = txtcliente.Text & ""
        If sw_load_mant = True Then Unload Me
        Case "Eliminar"
            eliminar_centro
        
        Case "Salir"
'            lista_almacen.adoctasctes.Refresh
'            If FrmName = "Lista_Centros" Then
'                Lista_Centros.Show
'            End If
            Unload Me
    End Select
End Sub

Private Sub TxtCodExt_GotFocus()
TxtCodExt.SelStart = 0: TxtCodExt.SelLength = Len(TxtCodExt.Text)
End Sub

Private Sub TxtCodExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub TxtPox_GotFocus()
    TxtPoX.SelStart = 0: TxtPoX.SelLength = Len(TxtPoX.Text)

End Sub

Private Sub TxtPox_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If (KeyAscii > 46 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 45 Or (KeyAscii > 64 And KeyAscii < 91) Then
    KeyAscii = KeyAscii
ElseIf KeyAscii = 13 Then
    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
Else
    KeyAscii = 0
End If
End Sub
Private Sub genera_cod()
SqlCad = "select F3COSTO FROM CENTROS where f3costo<>'999' and f3costo<>'998' and IntCodigoNivel=" & nGridActive & " and f3costo like '" & wcodcosto & "%' order by F3COSTO desc"
Set RsDemo = Af.OpenSQLForwardOnly(SqlCad, cconex_dbbancos)
If RsDemo.RecordCount > 0 Then
    RsDemo.MoveFirst
    Codigo = RsDemo.Fields("F3COSTO") + 1
    Codigo = Format(Codigo, Repetir(nGridActive * 2, "0"))
Else
    Codigo = 1
    Codigo = Format(Codigo, "00")
    Codigo = wcodcosto & Codigo
End If
    
End Sub

Private Sub Form_Load()
  


If sw_nuevo_doc = True Then
    nuevo_centro
Else
    actualizacion_centro Lista_Centros.dxDBGrid(nGridActive).Columns(0).Value
End If

End Sub

Private Sub actualizacion_centro(cod)
SqlCad = "select * from CENTROS where F3COSTO='" & cod & "'"
Set RsDemo = Af.OpenSQLForwardOnly(SqlCad, cconex_dbbancos)
If Not RsDemo.EOF Then
    TxtCodigo.Text = "" & RsDemo.Fields("F3COSTO")
    txtunidad.Text = "" & RsDemo.Fields("F3ABREV")
    txtdescrip.Text = "" & RsDemo.Fields("F3DESCRIP")
    txtcliente.Text = RsDemo.Fields("f3codcli") & ""
    TxtUtil.Text = Format(Val(RsDemo.Fields("UTILIDAD") & ""), "###,##0.00")
    TxtPoX.Text = RsDemo!PO & ""
    TxtCodExt.Text = RsDemo!CCONCAR & ""
    If Trim(RsDemo!F3ESTNUL & "") = "N" Then
        OptEst(0).Value = True
    Else
        OptEst(1).Value = True
    End If
'    txtfecha.Value = "" & IIf(IsDate(rsdemo.Fields("F3FECHA")), Format(rsdemo.Fields("F3FECHA"), "DD/MM/YYYY"), CVDate(Format(Date, "DD/MM/YYYY")))
    
    pnlcliente.Text = ObtenerCampo("EF2CLIENTES", "F2NOMCLI", "F2CODCLI", txtcliente.Text & "", "T", cnn_dbbancos)
    TxtCodigo.Enabled = False
End If
RsDemo.Close

End Sub

Private Sub nuevo_centro()
If sw_load_mant = True Then
'    SSActiveToolBars1.Tools.Item("ID_Eliminar").Visible = False
'    SSActiveToolBars1.Tools.Item("ID_Lista").Visible = False
End If
genera_cod
TxtCodigo.Text = Codigo
TxtCodigo.Enabled = False
txtunidad.Text = ""
TxtPoX.Text = ""
'TxtUtil.Text = "20.00"
txtdescrip.Text = ""
txtcliente.Text = ""
pnlcliente.Text = ""
'txtfecha.Value = Date
'txtdescrip.SetFocus
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.ID
        Case "ID_Nuevo"
        sw_nuevo_doc = True
        nuevo_centro
        txtunidad.SetFocus
        Case "ID_Grabar"
        
'        If Len(Trim(txtcodigo.Text)) = 0 Then
'            MsgBox "Datos Incompletos", vbInformation, "Atención"
'            Exit Sub
'        End If
        If Len(Trim(txtunidad.Text)) = 0 Then
            MsgBox "Datos Incompletos", vbInformation, "Atención"
            txtunidad.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtdescrip.Text)) = 0 Then
            MsgBox "Datos Incompletos", vbInformation, "Atención"
            txtdescrip.SetFocus
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        grabar_centro
        Me.MousePointer = 0
        wcodcosto = TxtCodigo.Text & ""
        wdescosto = txtdescrip.Text & ""
        wunicosto = txtunidad.Text & ""
        wclicosto = txtcliente.Text & ""
        If sw_load_mant = True Then Unload Me
        Case "ID_Eliminar"
        eliminar_centro
        Case "ID_Lista"
'            lista_almacen.adoctasctes.Refresh
            Unload Me
        Case "ID_Salir"
'            lista_almacen.adoctasctes.Refresh
            Unload Me
    End Select
End Sub

Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)

'   Select Case Button.Index
'      Case Is = 1
'        sw_nuevo_doc = True
'        nuevo_centro
'        txtdescrip.SetFocus
'      Case Is = 2
'        Me.MousePointer = vbhourglass
'        grabar_centro
'        Me.MousePointer = 0
''         Limpia_Variables
''         txtcodigo.SetFocus
'      Case Is = 3
'        eliminar_centro
'      Case Is = 4
'         cryreporte.DataFiles(0) = wrutabanco & "\DB_BANCOS.MDB"
'         cryreporte.ReportFileName = wrutatemp & "centros.rpt"
'         cryreporte.Action = 1
'      Case Is = 5
'         Unload Me
'   End Select

End Sub

Private Sub eliminar_centro()
    Beep
    If MsgBox("¿Está seguro de eliminar el Centro de Costo?", 36, wnomcia) = 6 Then
    
            
            csql = "delete * from centros where F3COSTO='" & TxtCodigo.Text & "'"
            cnn_dbbancos.Execute csql
            Unload Me
    
    End If

End Sub

Private Sub txtcliente_DblClick()

   txtcliente_KeyDown 113, 0

End Sub

Private Sub txtcliente_GotFocus()
txtcliente.SelStart = 0: txtcliente.SelLength = Len(txtcliente.Text)
End Sub

Private Sub txtcliente_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcodcliprov = ""
        ayuda_clientes.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcliprov)) > 0 Then
            txtcliente.Text = wcodcliprov
            pnlcliente.Text = wnomcliprov
            txtCodigo_KeyPress 13
        End If
    End If

End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      If Len(Trim(txtcliente.Text)) > 0 Then
         SqlCad = "SELECT F2NOMCLI FROM EF2CLIENTES WHERE F2CODCLI='" & txtcliente.Text & "'"
         Set RsDemo = Af.OpenSQLForwardOnly(SqlCad, cconex_dbbancos)
         If RsDemo.RecordCount > 0 Then
            pnlcliente.Text = RsDemo.Fields("f2nomcli") & ""
            ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
         Else
            MsgBox "Código del cliente no existe. Verifique.", 48, wnomcia
            txtcliente.SetFocus
         End If
      Else
         pnlcliente.Text = ""
      End If
   End If

End Sub

Private Sub txtCodigo_DblClick()

   If LLAMADA <> "NSE" Then
      txtcodigo_KeyUp 113, 0
   End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

'   If KeyAscii = 13 Then
'      If Len(Trim("" & txtcodigo.Text)) = 0 Then
'         txtcodigo.SetFocus
'         Exit Sub
'      End If
'      actualiza
'      txtdescrip.SetFocus
'   End If

End Sub

Private Sub txtcodigo_KeyUp(KeyCode As Integer, Shift As Integer)

'   If KeyCode = 113 Then
'      Me.MousePointer = vbhourglass
'      wcentro = ""
'      FrmHlpcentros.Show 1
'      Me.MousePointer = vbdefault
'      If Len(Trim(wcentro)) > 0 Then
'         txtcodigo.Text = Trim("" & wcentro)
'         actualiza
'         txtdescrip.SetFocus
'      Else
'         txtcodigo.SetFocus
'      End If
'   End If

End Sub

Private Sub txtdescrip_GotFocus()
    txtdescrip.SelStart = 0: txtdescrip.SelLength = Len(txtdescrip.Text)
End Sub

Private Sub txtdescrip_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"

End Sub


Private Sub grabar_centro()
On Error GoTo graba
Dim ctipoalm        As String
Dim cestnul        As String
Dim amovs(0 To 15)  As a_grabacion

    SqlCad = "select * from centros where f3costo='" & TxtCodigo.Text & "'"
    Set RsDemo = Af.OpenSQLForwardOnly(SqlCad, cconex_dbbancos)
    If Not RsDemo.EOF Then
        FL = 0
        amovs(4).campo = "F3FECMOD": amovs(4).valor = Format(Date, "dd/mm/yyyy"): amovs(4).Tipo = "F"
        amovs(5).campo = "USEMOD": amovs(5).valor = wusuario: amovs(5).Tipo = "T"
    Else
        FL = 1
        amovs(4).campo = "F3FECGRA": amovs(4).valor = Format(Date, "dd/mm/yyyy"): amovs(4).Tipo = "F"
        amovs(5).campo = "USEGRA": amovs(5).valor = wusuario: amovs(5).Tipo = "T"
    End If
    amovs(0).campo = "F3COSTO": amovs(0).valor = TxtCodigo.Text: amovs(0).Tipo = "T"
    amovs(1).campo = "F3DESCRIP": amovs(1).valor = txtdescrip.Text: amovs(1).Tipo = "T"
    amovs(2).campo = "F3CODCLI": amovs(2).valor = txtcliente.Text: amovs(2).Tipo = "T"
    amovs(3).campo = "F3ABREV": amovs(3).valor = txtunidad.Text: amovs(3).Tipo = "T"
    amovs(6).campo = "PO": amovs(6).valor = TxtPoX.Text: amovs(6).Tipo = "T"
    amovs(7).campo = "UTILIDAD": amovs(7).valor = 20: amovs(7).Tipo = "N"
    
If OptEst(0).Value = True Then
    cestnul = "N"
Else
    cestnul = "S"
End If

    amovs(8).campo = "F3ESTNUL": amovs(8).valor = cestnul: amovs(8).Tipo = "T"
    amovs(9).campo = "CCONCAR": amovs(9).valor = "000": amovs(9).Tipo = "T"
    amovs(10).campo = "INTCODIGONIVEL": amovs(10).valor = nGridActive: amovs(10).Tipo = "N"
    amovs(11).campo = "F3DIRECCION": amovs(11).valor = TxtDireccion.Text: amovs(11).Tipo = "T"
    
    RsDemo.Close
    Set RsDemo = Nothing
    
    If FL = 1 Then
        GRABA_REGISTRO_logistica amovs(), "CENTROS", "A", 11, cnn_dbbancos, ""
        MsgBox "Registro grabado.", vbInformation, wnomcia
    Else
        GRABA_REGISTRO_logistica amovs(), "CENTROS", "M", 11, cnn_dbbancos, "F3COSTO='" & TxtCodigo.Text & "'"
        MsgBox "Registro Actualizado.", vbInformation, wnomcia
    End If
    TxtCodigo.Enabled = False
    'MsgBox "El Registro se ha Actualizado", vbInformation, "Sistema de Gerencial"
    
    Exit Sub
    Resume
graba:
    If Err = 3186 Then
        For I% = 1 To 10000
        Next I%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", vbCritical, wnomcia
        Resume
    Else
        'MsgBox "Se ha producido el sgte. error " & ERROR(Err), 48, "Atención"
        MsgBox "Se ha producido el sgte. error " & Error(Err), vbCritical, wnomcia
        Exit Sub
    End If

End Sub

'Private Sub txtfecha_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      txtdescrip.SetFocus
'   End If
'End Sub


'Private Sub infocentro()
'
'   If datacentro.Recordset.RecordCount > 0 Then
'      txtcodigo.Text = datacentro.Recordset.Fields("F3COSTO")
'      txtdescrip.Text = datacentro.Recordset.Fields("F3DESCRIP") & ""
'      txtcliente.Text = datacentro.Recordset.Fields("F3CODCLI") & ""
'      tbclientes.Seek "=", txtcliente.Text
'      If Not tbclientes.NoMatch Then
'         pnlcliente.text = Trim(tbclientes.Fields("f2nomcli") & "")
'      Else
'         pnlcliente.text = ""
'      End If
'   End If
'
'End Sub

Private Sub txtunidad_GotFocus()
    txtunidad.SelStart = 0: txtunidad.SelLength = Len(txtunidad.Text)
End Sub

Private Sub txtunidad_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"

End Sub

Private Sub txtunidad_LostFocus()

    If sw_nuevo_doc = True Then
        If Len(Trim(txtunidad.Text)) > 0 Then
            strSQL = "SELECT F3ABREV FROM CENTROS WHERE F3ABREV='" & Trim(txtunidad.Text) & "'"
            Set RsDemo = Af.OpenSQLForwardOnly(strSQL, cconex_dbbancos)
            If Not RsDemo.EOF Then
                txtunidad.Text = ""
                MsgBox "El Centro de Costo ya existe. Verifíque.", vbInformation, wnomcia
                txtunidad.SetFocus
'            Else
'                pnlcosto.Caption = "": txtcentro.Text = ""
'                MsgBox "Código del Centro de Costo no existe. Verifique.", vbInformation, "Atención"
'                txtcentro.SetFocus
            End If
            RsDemo.Close
            Set RsDemo = Nothing
        End If
    End If
End Sub

Private Sub TxtUtil_GotFocus()
    TxtUtil.SelStart = 0: TxtUtil.SelLength = Len(TxtUtil.Text)
End Sub

Private Sub TxtUtil_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub TxtUtil_LostFocus()
If IsNumeric(TxtUtil.Text) Then
    TxtUtil.Text = Format(TxtUtil, "#,#0.00")
Else
    TxtUtil.SetFocus
End If
End Sub
