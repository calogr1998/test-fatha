VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form actualiza_bancos 
   Appearance      =   0  'Flat
   Caption         =   "Actualización de Bancos"
   ClientHeight    =   2340
   ClientLeft      =   2850
   ClientTop       =   2385
   ClientWidth     =   5685
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2340
   ScaleWidth      =   5685
   Begin Threed.SSPanel SSPanel1 
      Height          =   1635
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   5550
      _Version        =   65536
      _ExtentX        =   9790
      _ExtentY        =   2884
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox TxtTipCam 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   3645
         TabIndex        =   19
         Text            =   "0.000"
         Top             =   1215
         Width           =   780
      End
      Begin VB.TextBox txtnumdoc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   2
         Top             =   855
         Width           =   1095
      End
      Begin VB.TextBox txtcodtip 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   1
         Top             =   495
         Width           =   465
      End
      Begin VB.TextBox txtcodcta 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   0
         Top             =   135
         Width           =   465
      End
      Begin VB.TextBox txtcaja 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4905
         MaxLength       =   4
         TabIndex        =   3
         Top             =   855
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtmesc 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3645
         MaxLength       =   2
         TabIndex        =   9
         Top             =   855
         Visible         =   0   'False
         Width           =   465
      End
      Begin Threed.SSPanel pncuenta 
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   135
         Width           =   3210
         _Version        =   65536
         _ExtentX        =   5662
         _ExtentY        =   503
         _StockProps     =   15
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
      Begin Threed.SSPanel pntipope 
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   495
         Width           =   3210
         _Version        =   65536
         _ExtentX        =   5662
         _ExtentY        =   503
         _StockProps     =   15
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
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   1620
         TabIndex        =   4
         Top             =   1215
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label lbdoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nro de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   900
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código de Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label lblcaja 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nº Caja : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4185
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblmescaja 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes Caja :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2835
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "T/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   1260
         Width           =   285
      End
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   420
      Left            =   855
      TabIndex        =   5
      Top             =   1800
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox cryreporte 
      Height          =   480
      Left            =   135
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   20
      Top             =   1755
      Width           =   1200
   End
   Begin Threed.SSCommand cmdimprimir 
      Height          =   420
      Left            =   2205
      TabIndex        =   6
      Top             =   1800
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   420
      Left            =   3555
      TabIndex        =   7
      Top             =   1800
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
End
Attribute VB_Name = "actualiza_bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''Dim xcta        As String
'''''Dim flag        As Integer
'''''
'''''Private Sub cmdaceptar_Click()
'''''
'''''    If Len(Trim(txtcodcta.Text & "")) > 0 And Len(Trim(txtcodtip.Text & "")) > 0 Then
'''''        NCTA = txtcodcta.Text
'''''        If xcta <> "1" And xcta <> "2" Then
'''''            grabar_caja
'''''        Else
'''''            grabar_bancos
'''''        End If
'''''        cmdaceptar.Enabled = False
'''''        cmdimprimir.Enabled = True
'''''        cmdimprimir.SetFocus
'''''    Else
'''''        MsgBox "Error en el ingreso de los datos. Verifique", 48, "Atención"
'''''        txtcodcta.SetFocus
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub CmdImprimir_Click()
'''''
'''''    Imprimir
'''''    cryreporte.DataFiles(0) = wrutatemp & "\" & "temp_com.mdb"
'''''    If wf1formatov = "2" Then
'''''        cryreporte.ReportFileName = wrutatemp & "\" & "vouctas2.rpt"
'''''    Else
'''''        cryreporte.ReportFileName = wrutatemp & "\" & "vouctasp.rpt"
'''''    End If
'''''    cryreporte.Action = 1
'''''
'''''End Sub
'''''
'''''Private Sub cmdsalir_Click()
'''''
'''''    Unload Me
'''''
'''''End Sub
'''''
'''''Private Sub Form_Load()
'''''
'''''    Set dbbancos = OpenDatabase(wrutabanco & "\db_tabla.mdb")
'''''    Set tbcta = dbbancos.OpenRecordset("bf5pla")
'''''    Set TBBANCO = dbbancos.OpenRecordset("bf8tmov")
'''''    Set tbcheques = dbbancos.OpenRecordset("BF6CHQ")
'''''    Set TBBF4TCO = dbbancos.OpenRecordset("bf4tco" & registro_compras.txtmesmov.Text)
'''''    flag = False
'''''
'''''    txtfecha.Text = registro_compras.txtfecha.Text
'''''    TxtTipCam.Text = registro_compras.TxtTipCam.Text
'''''
'''''End Sub
'''''
'''''Private Sub Form_Unload(Cancel As Integer)
'''''On Error GoTo error_bd
'''''
'''''    tbcta.Close
'''''    TBBANCO.Close
'''''    tbcheques.Close
'''''    TBBF4TCO.Close
'''''    dbbancos.Close
'''''
'''''    Exit Sub
'''''
'''''error_bd:
'''''    Resume Next
'''''
'''''End Sub
'''''
'''''Private Sub Imprimir()
'''''Dim dbtempcom As Database
'''''Dim tbcab As Recordset
'''''Dim tbdet As Recordset
'''''Dim tbcabecera As Recordset
'''''Dim tbdetalle As Recordset
'''''Dim tbbancos As Recordset
'''''
'''''    Set dbtempcom = OpenDatabase(wrutatemp & "\temp_com.mdb")
'''''    dbtempcom.Execute ("Delete * From timp_det")
'''''    dbtempcom.Execute ("Delete * From timp_cab")
'''''
'''''    Set tbcab = dbtempcom.OpenRecordset("timp_cab")
'''''    Set tbdet = dbtempcom.OpenRecordset("timp_det")
'''''
'''''    Set dbmovis = OpenDatabase(wrutabanco & "\db_tabla.mdb")
'''''    Set tbcabecera = dbmovis.OpenRecordset("BF4TCO" & Format(mes, "00"))
'''''    tbcabecera.Index = "BF4TCO" & Format(mes, "00")
'''''
'''''    Set tbdetalle = dbmovis.OpenRecordset("bf3mov" & Format(mes, "00"))
'''''    tbdetalle.Index = "bf3mov" & Format(mes, "00")
'''''
'''''    Set tbbancos = dbmovis.OpenRecordset("bancos")
'''''
'''''    tbcabecera.Seek "=", Val(txtcodcta.Text), Val(nmovbanco)
'''''    If Not tbcabecera.NoMatch Then
'''''        tbcab.AddNew
'''''        tbcab.Fields("tnummov") = nmovbanco
'''''        tbcab.Fields("tfecha") = Format(Date, "dd/mm/yyyy")
'''''        tbcab.Fields("tfechareg") = tbcabecera.Fields("f4fecgir")
'''''
'''''        tbcta.Index = "idcodigo"
'''''        tbcta.Seek "=", Val(txtcodcta.Text)
'''''        If Not tbcta.NoMatch Then
'''''            tbcab.Fields("tctacte") = tbcta.Fields("numcta")
'''''            tbcab.Fields("tmoneda") = UCase(tbcta.Fields("moneda"))
'''''            tbcab.Fields("tsimbolo") = IIf(tbcta.Fields("dolar") = "S", "S/.", "US$")
'''''            tbbancos.Index = "idcodigo"
'''''            tbbancos.Seek "=", tbcta.Fields("codban")
'''''            If Not tbbancos.NoMatch Then
'''''                tbcab.Fields("tbanco") = tbbancos.Fields("banco") & ""
'''''            End If
'''''        End If
'''''
'''''        tbcab.Fields("tcodcta") = txtcodcta.Text
'''''        tbcab.Fields("tcheque") = tbcabecera.Fields("f4numdoc")
'''''        tbcab.Fields("ttc") = tbcabecera.Fields("f4tipcamb")
'''''        tbcab.Fields("ttipope") = tbcabecera.Fields("f4tipmov")
'''''        tbcab.Fields("tnomope") = Trim(pntipope.Caption)
'''''        tbcab.Fields("tgiradoa") = tbcabecera.Fields("f4giradoa")
'''''        tbcab.Fields("tconcepto") = tbcabecera.Fields("f4detal")
'''''        tbcab.Fields("tobserva1") = tbcabecera.Fields("f4observa1")
'''''        tbcab.Fields("tobserva2") = tbcabecera.Fields("f4observa2")
'''''        tbcab.Fields("tempresa") = GCODEMP
'''''        tbcab.Fields("tproyecto") = ""
'''''        tbcab.Fields("tnomproye") = ""
'''''        tbcab.Fields("tusuario") = gcoduse
'''''        tbcab.Update
'''''        '------------------------------------------------
'''''        tbdet.AddNew
'''''        tbdet.Fields("tnummov") = nmovbanco
'''''        tbcta.Index = "idcodigo"
'''''        tbcta.Seek "=", Val(txtcodcta.Text)
'''''        If Not tbcta.NoMatch Then
'''''            tbdet.Fields("tcuenta") = tbcta.Fields("f5codcta")
'''''            tbdet.Fields("tconcepto") = tbcabecera.Fields("f4detal")
'''''            If tbcabecera.Fields("f4destino") = "I" Then
'''''                tbdet.Fields("tdebe") = tbcabecera.Fields("f4total")
'''''            Else
'''''                tbdet.Fields("thaber") = tbcabecera.Fields("f4total")
'''''            End If
'''''        End If
'''''        tbdet.Fields("tdocum") = tbcabecera.Fields("f4numdoc")
'''''        tbdet.Update
'''''        '------------------------------------------------
'''''        tbdetalle.Seek "=", Val(txtcodcta.Text), Val(nmovbanco)
'''''        If Not tbdetalle.NoMatch Then
'''''            Do While tbdetalle.Fields("codcta") = Val(txtcodcta.Text) And tbdetalle.Fields("nummov") = Val(nmovbanco) And Not tbdetalle.EOF
'''''                tbdet.AddNew
'''''                tbdet.Fields("tnummov") = nmovbanco
'''''                tbdet.Fields("tcuenta") = tbdetalle.Fields("codgto")
'''''                tbdet.Fields("treferencia") = tbdetalle.Fields("orden")
'''''                tbdet.Fields("tconcepto") = tbdetalle.Fields("concepto")
'''''                tbdet.Fields("ttipdocu") = tbdetalle.Fields("tipdocu")
'''''                tbdet.Fields("tserdoc") = Format(tbdetalle.Fields("ser_doc"), "000")
'''''                tbdet.Fields("tdocum") = Format(tbdetalle.Fields("docum"), "0000000")
'''''                If tbdetalle.Fields("f3debhab") = "D" Then
'''''                    tbdet.Fields("tdebe") = tbdetalle.Fields("parcial")
'''''                Else
'''''                    tbdet.Fields("thaber") = tbdetalle.Fields("parcial")
'''''                End If
'''''                tbdet.Fields("tcosto") = tbdetalle.Fields("f3costo")
'''''                tbdet.Update
'''''                tbdetalle.MoveNext
'''''                If tbdetalle.EOF Then Exit Do
'''''                If tbdetalle.Fields("codcta") <> Val(txtcodcta.Text) Or tbdetalle.Fields("nummov") <> Val(nmovbanco) Then Exit Do
'''''            Loop
'''''        End If
'''''    End If
'''''    tbbancos.Close
'''''    tbdetalle.Close
'''''    tbcabecera.Close
'''''    dbmovis.Close
'''''
'''''    tbdet.Close
'''''    tbcab.Close
'''''    dbtempcom.Close
'''''
'''''End Sub
'''''
'''''Private Sub txtcaja_KeyPress(KeyAscii As Integer)
'''''
'''''    If KeyAscii = 13 Then
'''''        If Val(txtcaja.Text & "") <> xcaja And xcaja <> 0 Then
'''''            MsgBox "El número de caja asignado no corresponde con el de Bancos.", 48, "Compras"
'''''            txtcaja.SetFocus
'''''        Else
'''''            cmdaceptar.SetFocus
'''''        End If
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtcaja_LostFocus()
'''''
'''''    If Len(Trim(txtcaja.Text)) > 0 Then
'''''        If Val(txtcaja.Text & "") <> xcaja And xcaja <> 0 Then
'''''            MsgBox "El número de caja asignado no corresponde con el de Bancos.", 48, "Compras"
'''''            txtcaja.SetFocus
'''''        Else
'''''            cmdaceptar.SetFocus
'''''        End If
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub TxtCodcta_DblClick()
'''''
'''''    txtcodcta_KeyDown 113, 0
'''''
'''''End Sub
'''''
'''''Private Sub txtcodcta_KeyDown(KeyCode As Integer, Shift As Integer)
'''''
'''''    If KeyCode = 113 Then
'''''        flag = True
'''''        frmhlpctas.Show 1
'''''        flag = False
'''''        txtcodcta.Text = gcodcta
'''''        txtcodcta_KeyPress 13
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtcodcta_KeyPress(KeyAscii As Integer)
'''''
'''''    If KeyAscii = 13 Then
'''''        tbcta.Index = "idcodigo"
'''''        tbcta.Seek "=", Val(txtcodcta.Text)
'''''        If Not tbcta.NoMatch Then
'''''            pncuenta.Caption = Trim(tbcta.Fields("numcta") & "")
'''''            xcta = tbcta.Fields("tipo") & ""
'''''            xmoneda = tbcta.Fields("dolar") & ""
'''''            If xcta <> "1" And xcta <> "2" Then
'''''                lblcaja.Visible = True
'''''                txtcaja.Visible = True
'''''                lblmescaja.Visible = True
'''''                txtmesc.Visible = True
'''''                txtcaja.Text = Val(tbcta.Fields("numcaja") & "")
'''''                txtmesc.Text = Format(Val(tbcta.Fields("mescaja") & ""), "00")
'''''                xcaja = Val(tbcta.Fields("numcaja") & "")
'''''                xmovcaja = Val(tbcta.Fields("nummovcaja") & "")
'''''                xmescaja = Val(tbcta.Fields("mescaja") & "")
'''''                xctacaja = Val(txtcodcta.Text)
'''''                If xmescaja = 0 Then
'''''                    Rem xmescaja = Month(registro_compras.TxtFecha.Text)
'''''                    xmescaja = Month(registro_compras.TxtFecVen.Text)
'''''                    txtmesc.Text = Format(xmescaja, "00")
'''''                End If
'''''            Else
'''''                lblcaja.Visible = False
'''''                txtcaja.Visible = False
'''''                lblmescaja.Visible = False
'''''                txtmesc.Visible = False
'''''            End If
'''''            txtcodtip.SetFocus
'''''        Else
'''''            MsgBox "Código de cuenta no existe. Verifique.", 48, "Compras"
'''''            txtcodcta.SetFocus
'''''        End If
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtcodcta_LostFocus()
'''''
'''''    If Len(Trim(txtcodcta.Text)) > 0 And flag = False Then
'''''        tbcta.Index = "idcodigo"
'''''        tbcta.Seek "=", Val(txtcodcta.Text)
'''''        If Not tbcta.NoMatch Then
'''''            pncuenta.Caption = tbcta.Fields("numcta") & ""
'''''            xcta = tbcta.Fields("tipo") & ""
'''''            xmoneda = tbcta.Fields("dolar") & ""
'''''            If xcta <> "1" And xcta <> "2" Then
'''''                lblcaja.Visible = True
'''''                txtcaja.Visible = True
'''''                lblmescaja.Visible = True
'''''                txtmesc.Visible = True
'''''                txtcaja.Text = Val(tbcta.Fields("numcaja") & "")
'''''                txtmesc.Text = Format(Val(tbcta.Fields("mescaja") & ""), "00")
'''''                xcaja = Val(tbcta.Fields("numcaja") & "")
'''''                xmovcaja = Val(tbcta.Fields("nummovcaja") & "")
'''''                xmescaja = Val(tbcta.Fields("mescaja") & "")
'''''                xctacaja = Val(txtcodcta.Text)
'''''                If xmescaja = 0 Then
'''''                    Rem xmescaja = Month(registro_compras.TxtFecha.Text)
'''''                    xmescaja = Month(registro_compras.TxtFecVen.Text)
'''''                    txtmesc.Text = Format(xmescaja, "00")
'''''                End If
'''''            Else
'''''                lblcaja.Visible = False
'''''                txtcaja.Visible = False
'''''                lblmescaja.Visible = False
'''''                txtmesc.Visible = False
'''''            End If
'''''            txtcodtip.SetFocus
'''''        Else
'''''            MsgBox "Código de cuenta no existe. Verifique.", 48, "Compras"
'''''            txtcodcta.SetFocus
'''''        End If
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtcodtip_Change()
'''''
'''''    TBBANCO.Index = "idcodigo"
'''''    TBBANCO.Seek "=", txtcodtip.Text
'''''    If Not TBBANCO.NoMatch Then
'''''      If TBBANCO.Fields("DESTINO") = "E" Then
'''''        If TBBANCO.Fields("marche") = "*" Then
'''''            lbdoc.Caption = "Cheque Nº"
'''''            pntipope.Caption = TBBANCO.Fields("tipo")
'''''        Else
'''''            lbdoc.Caption = "Nro de Documento"
'''''            pntipope.Caption = TBBANCO.Fields("tipo")
'''''        End If
'''''      End If
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtcodtip_dblclick()
'''''
'''''    txtcodtip_KeyDown 113, 0
'''''
'''''End Sub
'''''
'''''Private Sub txtcodtip_KeyDown(KeyCode As Integer, Shift As Integer)
'''''
'''''    If KeyCode = 113 Then
'''''        frmayuopera.Show 1
'''''        txtcodtip.Text = gcodtip
'''''        txtcodtip_KeyPress 13
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtcodtip_KeyPress(KeyAscii As Integer)
'''''
'''''    If KeyAscii = 13 Then
'''''        txtnumdoc.SetFocus
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub txtfecha_KeyPress(KeyAscii As Integer)
'''''
'''''    If KeyAscii = 13 Then
'''''        cmdaceptar.SetFocus
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub TxtFecha_LostFocus()
'''''
'''''    If Len(Trim(txtfecha.Text)) > 0 Then
'''''        tbcambios.Index = "cambios"
'''''        tbcambios.Seek "=", CVDate(txtfecha.Text)
'''''        If Not tbcambios.NoMatch Then
'''''            TxtTipCam.Text = Format(tbcambios.Fields("cambio"), "0.000")
'''''        End If
'''''    End If
'''''
'''''End Sub
'''''
'''''Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
'''''
'''''    If KeyAscii = 13 Then
'''''        If lbdoc.Caption = "Cheque Nº" Then
'''''            tbtalon.Index = "idcodcta"
'''''            tbtalon.Seek "=", Val(txtcodcta.Text)
'''''            xtalon = 0
'''''            If Not tbtalon.NoMatch Then
'''''                Do While tbtalon.Fields("codcta") = Val(txtcodcta.Text) And Not tbtalon.EOF
'''''                    If (Val(txtnumdoc.Text) >= tbtalon.Fields("cdesde")) And (Val(txtnumdoc.Text) <= tbtalon.Fields("chasta")) Then
'''''                        xtalon = tbtalon.Fields("codtal")
'''''                        Exit Do
'''''                    End If
'''''                    tbtalon.MoveNext
'''''                    If tbtalon.EOF Then Exit Do
'''''                    If tbtalon.Fields("codcta") <> Val(txtcodcta.Text) Then Exit Do
'''''                Loop
'''''                If xtalon = 0 Then
'''''                    MsgBox "El Nº del cheque no pertenece a la cuenta No. " & txtcodcta.Text, 64, "Compras"
'''''                    txtnumdoc.Text = ""
'''''                    txtnumdoc.SetFocus
'''''                Else
'''''                    tbcheques.Index = "CTATALCHQ"
'''''                    tbcheques.Seek "=", Val(gcodcta), tbtalon.Fields("codtal"), Val(txtnumdoc.Text)
'''''                    If Not tbcheques.NoMatch Then
'''''                        MsgBox "El cheque ya fue girado el " & tbcheques.Fields("fecgir") & " en el movimiento No. " & tbcheques.Fields("nummov"), 64, "Compras"
'''''                        txtnumdoc.Text = ""
'''''                        txtnumdoc.SetFocus
'''''                    Else
'''''                        If xcta <> "1" And xcta <> "2" Then
'''''                            txtcaja.SetFocus
'''''                        Else
'''''                            txtfecha.SetFocus
'''''                        End If
'''''                    End If
'''''                End If
'''''            Else
'''''               MsgBox "La cuenta no tiene talonarios. Verifique.", 48, "Compras"
'''''               txtnumdoc.SetFocus
'''''            End If
'''''        Else
'''''            If xcta <> "1" And xcta <> "2" Then
'''''                txtcaja.SetFocus
'''''            Else
'''''                txtfecha.SetFocus
'''''            End If
'''''        End If
'''''    End If
'''''
'''''
'''''End Sub
'''''
