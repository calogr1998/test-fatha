VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form cons_reg_cnt_ret 
   Caption         =   "Registro Control de Retenciones"
   ClientHeight    =   2475
   ClientLeft      =   4005
   ClientTop       =   2445
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   7455
   Begin Threed.SSCommand cmdaceptar 
      Height          =   465
      Left            =   2205
      TabIndex        =   3
      Top             =   1890
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Height          =   1680
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   7260
      _Version        =   65536
      _ExtentX        =   12806
      _ExtentY        =   2963
      _StockProps     =   15
      BackColor       =   13160660
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
      Begin VB.TextBox txtcodprov 
         Height          =   285
         Left            =   1170
         TabIndex        =   2
         Top             =   810
         Width           =   810
      End
      Begin VB.TextBox txtanno 
         Height          =   330
         Left            =   5220
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   825
      End
      Begin VB.ComboBox cmbmes 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1860
      End
      Begin Threed.SSPanel pnlnombre 
         Height          =   285
         Left            =   1170
         TabIndex        =   8
         Top             =   1170
         Width           =   5925
         _Version        =   65536
         _ExtentX        =   10451
         _ExtentY        =   503
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlruc 
         Height          =   285
         Left            =   5220
         TabIndex        =   10
         Top             =   810
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   503
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4590
         TabIndex        =   11
         Top             =   810
         Width           =   480
      End
      Begin VB.Label lbltipo 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   315
         TabIndex        =   9
         Top             =   810
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   4590
         TabIndex        =   7
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   315
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   465
      Left            =   3690
      TabIndex        =   4
      Top             =   1890
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
End
Attribute VB_Name = "cons_reg_cnt_ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cconex_form     As String
Dim cnn_form        As New ADODB.Connection
Dim sw_ayuda        As Boolean
Dim rsproveedores   As New ADODB.Recordset

Private Sub cmbmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtanno.SetFocus
    End If

End Sub

Private Sub cmdaceptar_Click()

    Me.MousePointer = 11
    PROCESA
    Me.MousePointer = 1

End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    If wf1agente = "*" Then
        lbltipo.Caption = "Proveedor"
    Else
        lbltipo.Caption = "Cliente"
    End If

    cmbmes.AddItem "Enero    " & Space(80) & "01"
    cmbmes.AddItem "Febrero  " & Space(80) & "02"
    cmbmes.AddItem "Marzo    " & Space(80) & "03"
    cmbmes.AddItem "Abril    " & Space(80) & "04"
    cmbmes.AddItem "Mayo     " & Space(80) & "05"
    cmbmes.AddItem "Junio    " & Space(80) & "06"
    cmbmes.AddItem "Julio    " & Space(80) & "07"
    cmbmes.AddItem "Agosto   " & Space(80) & "08"
    cmbmes.AddItem "Setiembre" & Space(80) & "09"
    cmbmes.AddItem "Octubre  " & Space(80) & "10"
    cmbmes.AddItem "Noviembre" & Space(80) & "11"
    cmbmes.AddItem "Diciembre" & Space(80) & "12"
    cmbmes.ListIndex = 0
    
    txtanno.Text = Year(Date)
    
    sw_ayuda = False
    
End Sub

Private Sub PROCESA()
Dim rsretencab      As New ADODB.Recordset
Dim rsretendet      As New ADODB.Recordset
Dim SQL             As String
Dim sql2            As String
Dim ctipdoc         As String
Dim dfecha          As Date
Dim cserie          As String
Dim cdocum          As String
Dim ctipo           As String
Dim nhaber          As Double
Dim ndebe           As Double
Dim nanno           As Integer
Dim nmes            As Integer
Dim nitem           As Integer
Dim cruc            As String
Dim cnomprov        As String

    cnombase = "TMP_REND.MDB"
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    cnn_form.Execute ("DELETE * FROM TBREGISTRO")

    nanno = txtanno.Text
    nmes = Right(cmbmes.Text, 2)
    nitem = 0
    cruc = Trim(pnlruc.Caption)

    If Len(Trim(txtcodprov.Text)) = 0 Then
        SQL = "SELECT * FROM RETENDOC WHERE YEAR(FECHA)=" & nanno & " AND MONTH(FECHA)=" & nmes & " ORDER BY NOMBRE,SERIE,NUM_DOCUMENTO"
    Else
        SQL = "SELECT * FROM RETENDOC WHERE YEAR(FECHA)=" & nanno & " AND MONTH(FECHA)=" & nmes & " AND RUC='" & cruc & "' ORDER BY SERIE,NUM_DOCUMENTO"
    End If
    If rsretencab.State = adStateOpen Then rsretencab.Close
    rsretencab.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsretencab.EOF Then
        rsretencab.MoveFirst
        Do While Not rsretencab.EOF
            If rsretencab.Fields("ANULADO") <> "S" Then
                If Len(Trim(txtcodprov.Text)) = 0 Then
                    If wf1agente = "*" Then
                        If rsproveedores.State = adStateOpen Then rsproveedores.Close
                        rsproveedores.Open "SELECT F2NOMPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & rsretencab.Fields("RUC") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsproveedores.EOF Then
                            cnomprov = rsproveedores.Fields("F2NOMPROV") & ""
                        End If
                        rsproveedores.Close
                    Else
                        If rsproveedores.State = adStateOpen Then rsproveedores.Close
                        rsproveedores.Open "SELECT F2NOMCLI FROM EF2CLIENTES WHERE F2NEWRUC='" & rsretencab.Fields("RUC") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsproveedores.EOF Then
                            cnomprov = rsproveedores.Fields("F2NOMCLI") & ""
                        End If
                        rsproveedores.Close
                    End If
                Else
                    cnomprov = ""
                End If
            
                '----- BUSCA EN EL DETALLE
                sql2 = "SELECT * FROM RETENMOV WHERE SERIE_D='" & rsretencab.Fields("SERIE") & "' AND NUM_DOCUMENTOS='" & rsretencab.Fields("NUM_DOCUMENTO") & "' ORDER BY FECHA_EMISION"
                If rsretendet.State = adStateOpen Then rsretendet.Close
                rsretendet.Open sql2, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsretendet.EOF Then
                    Do While Not rsretendet.EOF
                        
                        '----- BUSCA EN TIPO DE DOCUMENTO
                        ctipdoc = ""
                        If rsdocumentos.State = adStateOpen Then rsdocumentos.Close
                        rsdocumentos.Open "SELECT F2DESDOC FROM DOCUMENTOS WHERE F2CODDOC='" & rsretendet.Fields("TIPO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsdocumentos.EOF Then
                            ctipdoc = Left(rsdocumentos.Fields("F2DESDOC") & "", 25)
                        End If
                        rsdocumentos.Close
                        
                        '----- 1.- ITEM X LA COMPRA
                        dfecha = Format(rsretendet.Fields("FECHA_EMISION"), "DD/MM/YYYY")
                        cserie = rsretendet.Fields("SERIE") & ""
                        cdocum = rsretendet.Fields("NUMERO_CORRELA") & ""
                        If wf1agente = "*" Then
                            ctipo = "Compra"
                        Else
                            ctipo = "Venta"
                        End If
                        nhaber = 0#
                        If Trim(rsretendet.Fields("TIPO") & "") = "07" Then
                            ndebe = Val(rsretendet.Fields("MONTO_PAGO") & "") * -1
                        Else
                            ndebe = Val(rsretendet.Fields("MONTO_PAGO") & "")
                        End If
                        nitem = nitem + 1
                        
                        cnn_form.Execute ("INSERT INTO TBREGISTRO (FECHA,TIPDOC,SERIE,NUMDOC,TRANSACCION,DEBE,HABER,ITEM,NOMPROV,TIPO) " & _
                                          "VALUES(cvdate('" & dfecha & "'),'" & ctipdoc & "','" & cserie & "','" & _
                                          cdocum & "','" & ctipo & "'," & ndebe & "," & nhaber & "," & nitem & ",'" & cnomprov & "','1')")
                                          
                        '----- 2.- ITEM X EL PAGO
                        dfecha = Format(rsretencab.Fields("FECHA"), "DD/MM/YYYY")
                        If wf1agente = "*" Then
                            ctipo = "Pago"
                        Else
                            ctipo = "Cobrado"
                        End If
                        If Trim(rsretendet.Fields("TIPO") & "") = "07" Then
                            nhaber = Val(Format(Val(rsretendet.Fields("MONTO_PAGO") & "") - Val(rsretendet.Fields("IMPORTE_RETENIDO") & ""), "0.00")) * -1
                        Else
                            nhaber = Val(Format(Val(rsretendet.Fields("MONTO_PAGO") & "") - Val(rsretendet.Fields("IMPORTE_RETENIDO") & ""), "0.00"))
                        End If
                        ndebe = 0#
                        nitem = nitem + 1
                        
                        cnn_form.Execute ("INSERT INTO TBREGISTRO (FECHA,TIPDOC,SERIE,NUMDOC,TRANSACCION,DEBE,HABER,ITEM,NOMPROV,TIPO) " & _
                                          "VALUES(cvdate('" & dfecha & "'),'" & ctipdoc & "','" & cserie & "','" & _
                                          cdocum & "','" & ctipo & "'," & ndebe & "," & nhaber & "," & nitem & ",'" & cnomprov & "','2')")
                                          
                        '----- 3.- ITEM X LA RETENCION
                        ctipdoc = "Comp. Retención"
                        dfecha = Format(rsretencab.Fields("FECHA"), "DD/MM/YYYY")
                        cserie = rsretencab.Fields("SERIE") & ""
                        cdocum = rsretencab.Fields("NUM_DOCUMENTO") & ""
                        ctipo = "Retención"
                        If Trim(rsretendet.Fields("TIPO") & "") = "07" Then
                            nhaber = Val(rsretendet.Fields("IMPORTE_RETENIDO") & "") * -1
                        Else
                            nhaber = Val(rsretendet.Fields("IMPORTE_RETENIDO") & "")
                        End If
                        ndebe = 0#
                        nitem = nitem + 1
                        
                        cnn_form.Execute ("INSERT INTO TBREGISTRO (FECHA,TIPDOC,SERIE,NUMDOC,TRANSACCION,DEBE,HABER,ITEM,NOMPROV,TIPO) " & _
                                          "VALUES(cvdate('" & dfecha & "'),'" & ctipdoc & "','" & cserie & "','" & _
                                          cdocum & "','" & ctipo & "'," & ndebe & "," & nhaber & "," & nitem & ",'" & cnomprov & "','3')")
                                          
                        rsretendet.MoveNext
                    Loop
                End If
            End If
            rsretencab.MoveNext
        Loop
    End If
    rsretencab.Close

    cnn_form.Close

    If Len(Trim(txtcodprov.Text)) = 0 Then
        With acr_reg_ret_todos
            .DataControl1.ConnectionString = cnn_form
            .DataControl1.Source = "SELECT * FROM TBREGISTRO ORDER BY ITEM"
            .lblempresa = wnomcia
            .lbltitulo = txtanno.Text & " - " & UCase(Left(cmbmes.Text, 10))
            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
            .Show 1
        End With
    Else
        With acr_reg_cnt_ret
            .DataControl1.ConnectionString = cnn_form
            .DataControl1.Source = "SELECT * FROM TBREGISTRO ORDER BY ITEM"
            .lblempresa = wnomcia
            .lbltitulo = txtanno.Text & " - " & UCase(Left(cmbmes.Text, 10))
            .fldfecha.Text = Format(Date, "DD/MM/YYYY")
            .fldnombre.Text = pnlnombre.Caption
            .fldruc.Text = pnlruc.Caption
            .Show 1
        End With
    End If

End Sub

Private Sub txtanno_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtcodprov.SetFocus
    End If

End Sub

Private Sub txtcodprov_DblClick()

    txtcodprov_KeyDown 113, 0
    
End Sub

Private Sub txtcodprov_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        If wf1agente = "*" Then
            wcodprov = ""
            sw_ayuda = True
            hlp_proveedores.Show 1
            sw_ayuda = False
            If Len(wcodprov) > 0 Then
                txtcodprov.Text = wcodprov
                pnlnombre.Caption = wnomprov
                pnlruc.Caption = wrucprov
            End If
        Else
            wcodcli = ""
            sw_ayuda = True
            hlp_clientes.Show 1
            sw_ayuda = False
            If Len(wcodcli) > 0 Then
                txtcodprov.Text = wcodcli
                pnlnombre.Caption = wnomcli
                pnlruc.Caption = wruccli
            End If
        End If
        
    End If

End Sub

Private Sub Txtcodprov_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If

End Sub

Private Sub txtcodprov_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtcodprov.Text)) > 0 Then
            If wf1agente = "*" Then
                If rsproveedores.State = adStateOpen Then rsproveedores.Close
                rsproveedores.Open "SELECT F2NOMPROV,F2NEWRUC FROM EF2PROVEEDORES WHERE F2CODPROV='" & Trim(txtcodprov.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsproveedores.EOF Then
                    pnlnombre.Caption = "" & rsproveedores.Fields("F2NOMPROV")
                    pnlruc.Caption = "" & rsproveedores.Fields("F2NEWRUC")
                End If
                rsproveedores.Close
            Else
                If rsproveedores.State = adStateOpen Then rsproveedores.Close
                rsproveedores.Open "SELECT F2NOMCLI,F2NEWRUC FROM EF2CLIENTES WHERE F2CODCLI='" & Trim(txtcodprov.Text) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsproveedores.EOF Then
                    pnlnombre.Caption = "" & rsproveedores.Fields("F2NOMCLI")
                    pnlruc.Caption = "" & rsproveedores.Fields("F2NEWRUC")
                End If
                rsproveedores.Close
            End If
        Else
            pnlnombre.Caption = "TODOS"
            pnlruc.Caption = ""
        End If
    End If
    
End Sub
