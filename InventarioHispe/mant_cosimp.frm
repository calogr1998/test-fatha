VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form mant_cosimp 
   Caption         =   "Actualización de Costos de Importación"
   ClientHeight    =   1812
   ClientLeft      =   2328
   ClientTop       =   2616
   ClientWidth     =   5892
   LinkTopic       =   "Form1"
   ScaleHeight     =   1812
   ScaleWidth      =   5892
   Begin Threed.SSPanel SSPanel1 
      Height          =   1590
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   5730
      _Version        =   65536
      _ExtentX        =   10107
      _ExtentY        =   2805
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin Threed.SSPanel pnlcodasoc 
         Height          =   285
         Left            =   2385
         TabIndex        =   7
         Top             =   1035
         Visible         =   0   'False
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   503
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.TextBox txtcodasoc 
         Height          =   285
         Left            =   1584
         TabIndex        =   2
         Top             =   1035
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtdescrip 
         Height          =   285
         Left            =   1575
         TabIndex        =   1
         Top             =   585
         Width           =   3975
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Left            =   1575
         TabIndex        =   0
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Asociado"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   180
         Width           =   495
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      Tools           =   "mant_cosimp.frx":0000
      ToolBars        =   "mant_cosimp.frx":7E74
   End
End
Attribute VB_Name = "mant_cosimp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstbcosimp As New ADODB.Recordset

Private Sub Form_Load()
    Me.Height = 7890
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 980
    sw_lleno = False
    sw_hlp = False
    If sw_nuevo_doc = True Then
        nuevo_cosimp
    Else
        actualizacion_cosimp lista_cosimp.tdbcostimp.Columns(0)
    End If
    
End Sub

Private Sub nuevo_cosimp()
    sw_lleno = False
'    txtcodigo.Enabled = True
    sql = "select * from tb_costosimp ORDER BY f2codigo DESC"
    If rstbcosimp.State = adStateOpen Then rstbcosimp.Close
    rstbcosimp.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rstbcosimp.EOF Then
        rstbcosimp.MoveFirst
        txtcodigo.Text = Format(Val("" & rstbcosimp.Fields("f2codigo")) + 1, "000")
    End If
    rstbcosimp.Close
    'txtcodigo.Text = ""
    txtdescrip.Text = ""
    txtcodasoc.Text = ""
    pnlcodasoc.Caption = ""
End Sub

Private Sub actualizacion_cosimp(cod)
    sql = "select * from tb_costosimp where f2codigo='" & cod & "'"
    If rstbcosimp.State = adStateOpen Then rstbcosimp.Close
    rstbcosimp.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rstbcosimp.EOF Then
        txtcodigo.Text = "" & rstbcosimp.Fields("f2codigo")
        txtdescrip.Text = "" & rstbcosimp.Fields("f2descripcion")
        txtcodasoc.Text = "" & rstbcosimp.Fields("f2codcon")
        pnlcodasoc.Caption = "" & rstbcosimp.Fields("f2descon")
    End If
    rstbcosimp.Close
    sw_lleno = True
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.ID
        Case "ID_Nuevo"
            sw_nuevo_doc = True
            nuevo_cosimp
            txtdescrip.SetFocus
        Case "ID_Grabar"
            grabar_cosimp
        Case "ID_Eliminar"
            eliminar_cosimp
        Case "ID_Imprimir"
            With acr_costimp
                .DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cnn_dbbancos & ";Persist Security Info=False"
                .DataControl1.Source = "Select * from tb_costosimp"
                .fldfecha.Text = Format(Date, "DD/MM/YYYY")
                .lblempresa.Caption = wempresa
                .Show 1
            End With
        Case "ID_Lista"
            lista_cosimp.adoctasctes.Refresh
            Unload Me
        Case "ID_Salir"
            lista_cosimp.adoctasctes.Refresh
            Unload Me
            Unload lista_cosimp
    End Select
End Sub

Private Sub txtcodasoc_DblClick()
txtcodasoc_KeyDown 113, 0
End Sub

Private Sub txtcodasoc_GotFocus()
txtcodasoc.SelStart = 0
txtcodasoc.SelLength = Len(txtcodasoc.Text)
End Sub

Private Sub txtcodasoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    sw_hlp = True
    hlp_asoc.Show 1
    txtcodasoc.Text = wcodigo
    pnlcodasoc.Caption = wnombre

End If

End Sub

Private Sub txtcodasoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdescrip.SetFocus
End If
End Sub

Private Sub txtcodigo_GotFocus()

txtcodigo.SelStart = 0
txtcodigo.SelLength = Len(txtcodigo.Text)

End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdescrip.SetFocus
End If
End Sub

Private Sub txtcodasoc_LostFocus()

    If Len(Trim(txtcodasoc.Text)) > 0 Then
        If sw_lleno = True And sw_hlp = False Then
            txtdescrip.SetFocus
            sw_lleno = False
        End If
        If sw_hlp = False Then
            sql = "select codigo,nombre from bf9gin where codigo='" & txtcodasoc.Text & "'"
            If rstbcosimp.State = adStateOpen Then rstbcosimp.Close
            rstbcosimp.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rstbcosimp.EOF Then
                pnlcodasoc.Caption = rstbcosimp.Fields("nombre")
            Else
                MsgBox "Codigo no existe.", vbInformation, "Mensaje"
                txtcodasoc.Text = ""
                pnlcodasoc.Caption = ""
                'txtcodasoc.SetFocus
            End If
            rstbcosimp.Close
        End If
        If txtcodasoc.Text = "" Then pnlcodasoc.Caption = ""
        sw_hlp = False
    End If

End Sub

Private Sub txtcodigo_LostFocus()
If sw_lleno = True Then
    txtdescrip.SetFocus
Else
    sql = "select * from TB_COSTOSIMP"
    If rstbcosimp.State = adStateOpen Then rstbcosimp.Close
    rstbcosimp.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    rstbcosimp.Find "f2codigo='" & txtcodigo.Text & "'"
    If Not rstbcosimp.EOF Then
        MsgBox "Codigo ya existe.", vbInformation, "Atención"
        txtcodigo.Text = ""
        txtcodigo.SetFocus
    Else
        txtdescrip.SetFocus
    End If
End If

End Sub

Private Sub txtdescrip_GotFocus()
txtdescrip.SelStart = 0
txtdescrip.SelLength = Len(txtdescrip.Text)
End Sub

Private Sub eliminar_cosimp()
 Beep
    If MsgBox("Está seguro de eliminar el concepto", 36, "Atención") = 6 Then
        sql = "select f2codigo from TB_COSTOSIMP where f2codigo='" & txtcodigo.Text & "'"
    If rstbcosimp.State = adStateOpen Then rstbcosimp.Close
    rstbcosimp.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rstbcosimp.EOF Then
        sql = "DELETE * from TB_COSTOSIMP where f2codigo='" & txtcodigo.Text & "'"
        cnn_dbbancos.Execute sql
        txtcodigo.Enabled = True
        nuevo_cosimp
    Else
        Beep
    End If
    rstbcosimp.Close
    txtcodigo.SetFocus
    End If

End Sub


Private Sub grabar_cosimp()
On Error GoTo graba
Dim amovs(0 To 3) As a_grabacion

sql = "select * from tb_costosimp where f2codigo='" & txtcodigo.Text & "'"
If rstbcosimp.State = adStateOpen Then rstbcosimp.Close
rstbcosimp.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not rstbcosimp.EOF Then
    sw = 0
Else
    sw = 1
End If
rstbcosimp.Close
amovs(0).campo = "f2codigo": amovs(0).valor = txtcodigo.Text: amovs(0).TIPO = "T"
amovs(1).campo = "f2descripcion": amovs(1).valor = txtdescrip.Text: amovs(1).TIPO = "T"
amovs(2).campo = "f2codcon": amovs(2).valor = txtcodasoc.Text: amovs(2).TIPO = "T"
amovs(3).campo = "f2descon": amovs(3).valor = pnlcodasoc.Caption: amovs(3).TIPO = "T"

If sw = 1 Then
    GRABA_REGISTRO amovs(), "tb_costosimp", "A", 3, cnn_dbbancos, ""
Else
    GRABA_REGISTRO amovs(), "tb_costosimp", "M", 3, cnn_dbbancos, "f2codigo='" & txtcodigo.Text & "'"
End If
txtcodigo.Enabled = False
Exit Sub


graba:
    If Err = 3186 Then
        For I% = 1 To 10000
        Next I%
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "Atención"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "Atención"
        Resume Next
    End If

End Sub


Private Sub txtdescrip_KeyPress(KeyAscii As Integer)
    
    'If KeyAscii = 13 Then
    '    txtcodasoc.SetFocus
    'End If

End Sub
