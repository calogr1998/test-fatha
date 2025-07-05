VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmcentros 
   Caption         =   "Centros de Costos"
   ClientHeight    =   2775
   ClientLeft      =   2835
   ClientTop       =   1935
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6345
   Begin Threed.SSCommand cmdfases 
      Height          =   375
      Left            =   4815
      TabIndex        =   10
      Top             =   45
      Visible         =   0   'False
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Fases"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2520
      Left            =   15
      TabIndex        =   3
      Top             =   120
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   4445
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   1
      Begin VB.TextBox txtFOB 
         Height          =   285
         Left            =   1275
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtunidad 
         Height          =   300
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin Threed.SSPanel pnlcliente 
         Height          =   285
         Left            =   2205
         TabIndex        =   9
         Top             =   1440
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.TextBox txtcliente 
         Height          =   300
         Left            =   1275
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1440
         Width           =   870
      End
      Begin VB.CommandButton Ayuda1 
         Caption         =   "..."
         Height          =   250
         Left            =   2175
         TabIndex        =   6
         ToolTipText     =   "Ayuda de los Centros de Costos"
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Data datacentro 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   2565
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   300
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox txtdescrip 
         Height          =   300
         Left            =   1275
         TabIndex        =   2
         Top             =   1080
         Width           =   4830
      End
      Begin VB.TextBox txtcodigo 
         Height          =   300
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   0
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "FOB"
         Height          =   255
         Left            =   315
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Neg."
         Height          =   195
         Left            =   315
         TabIndex        =   11
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   315
         TabIndex        =   8
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   315
         TabIndex        =   5
         Top             =   1095
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "frmcentros.frx":0000
      ToolBars        =   "frmcentros.frx":7E74
   End
End
Attribute VB_Name = "frmcentros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbempresa As DAO.Database
Dim tbclientes As DAO.Recordset
Dim RsDemo  As New ADODB.Recordset
Dim Codigo As String

Private Sub genera_cod()
sql = "select F3COSTO FROM CENTROS WHERE F3COSTO<>'999' order by F3COSTO desc"
If RsDemo.State = adStateOpen Then RsDemo.Close
RsDemo.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsDemo.EOF Then
    Codigo = RsDemo.Fields("F3COSTO") + 1
    Codigo = Format(Codigo, "000")
Else
    Codigo = 1
    Codigo = Format(Codigo, "000")
End If
End Sub

Private Sub Form_Load()
Me.left = 1500
Me.top = 980
If sw_nuevo_doc = True Then
    nuevo_centro
Else
    actualizacion_centro Lista_Centros.dxDBGrid(nGridActive).Columns(0).value
End If
End Sub

Private Sub actualizacion_centro(cod)
sql = "select * from CENTROS where F3COSTO='" & cod & "'"
If RsDemo.State = adStateOpen Then RsDemo.Close
RsDemo.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsDemo.EOF Then
    txtCodigo.Text = "" & RsDemo.Fields("F3COSTO")
    txtunidad.Text = "" & RsDemo.Fields("F3ABREV")
    txtdescrip.Text = "" & RsDemo.Fields("F3DESCRIP")
    txtcliente.Text = RsDemo.Fields("f3codcli") & ""
    If RsDemo.Fields("FOB") <> "" Then
    txtFOB.Text = Format(CDbl(RsDemo.Fields("FOB")), "##,#0.00")
    Else
    txtFOB.Text = Format(0, "##,#0.00")
    End If
    pnlcliente.Caption = traerCampo("EF2CLIENTES", "F2NOMCLI", "F2CODCLI", txtcliente.Text & "")
    txtCodigo.Enabled = False
End If
RsDemo.Close
End Sub

Private Sub nuevo_centro()
If sw_load_mant = True Then
    SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Visible = False
    SSActiveToolBars1.Tools.ITEM("ID_Lista").Visible = False
End If
genera_cod
txtCodigo.Text = Codigo
txtCodigo.Enabled = False
txtunidad.Text = ""
txtdescrip.Text = ""
txtcliente.Text = ""
txtFOB.Text = ""
pnlcliente.Caption = ""
'txtdescrip.SetFocus
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
        Case "ID_Nuevo"
        sw_nuevo_doc = True
        nuevo_centro
        txtunidad.SetFocus
        Case "ID_Grabar"
        Me.MousePointer = vbHourglass
        grabar_centro
        Me.MousePointer = 0
        wcodcosto = txtCodigo.Text & ""
        wdescosto = txtdescrip.Text & ""
        wunicosto = txtunidad.Text & ""
        wclicosto = txtcliente.Text & ""
        If txtFOB.Text <> "" Then
        wFob = Format(CDbl(txtFOB.Text), "##,#0.00")
        Else
        wFob = 0
        End If
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

'Private Sub tblbar_ButtonClick(ByVal Button As ComctlLib.Button)
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
'End Sub
Private Sub eliminar_centro()
    Beep
    If MsgBox("¿Está seguro de eliminar el Centro de Costo?", 36, "Atención") = 6 Then
    sql = "select * from  centros where f3costo='" & txtCodigo.Text & "'"
    If RsDemo.State = adStateOpen Then RsDemo.Close
    RsDemo.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsDemo.EOF Then
            sql = "DELETE * from centros where f3costo='" & txtCodigo.Text & "'"
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            nuevo_centro
        Else
            Beep
        End If
    RsDemo.Close
    txtdescrip.SetFocus
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
        'wcodcli = ""
        wcodcliprov = ""
        wnomcliprov = ""
        ayuda_clientes.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcliprov)) > 0 Then
            txtcliente.Text = wcodcliprov
            pnlcliente.Caption = wnomcliprov
            txtCodigo_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      If Len(Trim(txtcliente.Text)) > 0 Then
'         tbclientes.Seek "=", txtcliente.Text
'         If Not tbclientes.NoMatch Then
'            pnlcliente.Caption = tbclientes.Fields("f2nomcli") & ""
'            txtcodigo.SetFocus
'         Else
'            MsgBox "Código del cliente no existe. Verifique.", 48, "Contawin"
'            txtcliente.SetFocus
'         End If
'      Else
'         txtcodigo.SetFocus
'      End If
'   End If
    If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub txtcodigo_DblClick()
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
   If KeyAscii = 13 Then
      txtcliente.SetFocus
   End If
End Sub

Private Sub grabar_centro()
On Error GoTo graba
Dim ctipoalm        As String
Dim amovs(0 To 5)  As a_grabacion

    sql = "select * from centros where f3costo='" & txtCodigo.Text & "'"
    If RsDemo.State = adStateOpen Then RsDemo.Close
    RsDemo.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsDemo.EOF Then
        sw = 0
    Else
        sw = 1
    End If
    amovs(0).campo = "F3COSTO": amovs(0).valor = txtCodigo.Text: amovs(0).TIPO = "T"
    amovs(1).campo = "F3DESCRIP": amovs(1).valor = txtdescrip.Text: amovs(1).TIPO = "T"
    amovs(2).campo = "F3CODCLI": amovs(2).valor = txtcliente.Text: amovs(2).TIPO = "T"
    amovs(3).campo = "F3ABREV": amovs(3).valor = txtunidad.Text: amovs(3).TIPO = "T"
    amovs(4).campo = "F3FECGRA": amovs(4).valor = Format(Date, "DD/MM/YYYY"): amovs(4).TIPO = "F"
    'amovs(5).campo = "USEGRA": amovs(5).valor = wusuario: amovs(5).tipo = "T"
    If txtFOB.Text <> "" Then
            amovs(5).campo = "FOB": amovs(5).valor = Format(CDbl(txtFOB.Text), "##,#0.00"): amovs(5).TIPO = "T"
        '    amovs(4).campo = "F3FECGRA": amovs(4).valor = Format(Date, "DD/MM/YYYY"): amovs(4).tipo = "F"
        '    amovs(5).campo = "USEGRA": amovs(5).valor = wusuario: amovs(5).tipo = "T"
    Else
            amovs(5).campo = "FOB": amovs(5).valor = 0: amovs(5).TIPO = "T"
    End If
    RsDemo.Close
    Set RsDemo = Nothing
    
    If sw = 1 Then
            GRABA_REGISTRO_logistica amovs(), "CENTROS", "A", 5, cnn_dbbancos, ""
    Else
        GRABA_REGISTRO_logistica amovs(), "CENTROS", "M", 5, cnn_dbbancos, "F3COSTO='" & txtCodigo.Text & "'"
    End If
    txtCodigo.Enabled = False
    MsgBox "El Registro se ha Actualizado", vbInformation, "Sistema de Gerencial"
    Exit Sub
    
graba:
    If Err = 3186 Then
        For i% = 1 To 10000
        Next i%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "Atención"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "Atención"
        Resume Next
    End If
End Sub

Private Sub txtFOB_KeyPress(KeyAscii As Integer)
KeyAscii = Verificar_Tecla(KeyAscii)
End Sub

'Private Sub grabar()
'
'   If Len(Trim("" & txtcodigo.Text)) = 0 Then
'      Limpia_Variables
'      Exit Sub
'   End If
'
'   tbcentros.Index = "CENTROS"
'   tbcentros.Seek "=", txtcodigo.Text
'   If tbcentros.NoMatch Then
'      tbcentros.AddNew
'      tbcentros.Fields("F3COSTO") = txtcodigo.Text
'   Else
'      tbcentros.Edit
'   End If
'   tbcentros.Fields("F3DESCRIP") = txtdescrip.Text
'   tbcentros.Fields("F3CODCLI") = txtcliente.Text
'   tbcentros.Update
'   datacentro.Refresh
'   wcentro = txtcodigo.Text
'
'End Sub

'Private Sub infocentro()
'
'   If datacentro.Recordset.RecordCount > 0 Then
'      txtcodigo.Text = datacentro.Recordset.Fields("F3COSTO")
'      txtdescrip.Text = datacentro.Recordset.Fields("F3DESCRIP") & ""
'      txtcliente.Text = datacentro.Recordset.Fields("F3CODCLI") & ""
'      tbclientes.Seek "=", txtcliente.Text
'      If Not tbclientes.NoMatch Then
'         pnlcliente.Caption = Trim(tbclientes.Fields("f2nomcli") & "")
'      Else
'         pnlcliente.Caption = ""
'      End If
'   End If
'
'End Sub
Private Sub txtunidad_Change()

End Sub

Private Sub txtunidad_GotFocus()
    txtunidad.SelStart = 0: txtunidad.SelLength = Len(txtunidad.Text)
End Sub

Private Sub txtunidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtdescrip.SetFocus
   End If
End Sub

Function Verificar_Tecla(Tecla_Presionada)
        Dim Teclas As String
        'Acepta todos los números, la tecla Backspace, _
        la tecla Enter, la coma y el punto
        Teclas = "1234567890.," & Chr(vbKeyBack)
        If InStr(1, Teclas, Chr(Tecla_Presionada)) Then
                Verificar_Tecla = Tecla_Presionada
        Else
              ' Si no es ninguna de las indicadas retorna 0
                Verificar_Tecla = 0
        End If
End Function
