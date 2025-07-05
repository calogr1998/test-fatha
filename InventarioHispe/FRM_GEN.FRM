VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_gen 
   Caption         =   "Generación de Asientos Contables"
   ClientHeight    =   1935
   ClientLeft      =   3525
   ClientTop       =   1935
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4215
   Begin Threed.SSPanel SSPanel1 
      Height          =   1230
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   4020
      _Version        =   65536
      _ExtentX        =   7091
      _ExtentY        =   2170
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
      Begin VB.TextBox txtmes 
         Height          =   285
         Left            =   1035
         MaxLength       =   2
         TabIndex        =   0
         Top             =   225
         Width           =   420
      End
      Begin VB.TextBox txtanno 
         Height          =   285
         Left            =   2790
         MaxLength       =   4
         TabIndex        =   1
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   270
         Width           =   300
      End
      Begin VB.Label lblcompro 
         Caption         =   "Nº Mov."
         Height          =   285
         Left            =   495
         TabIndex        =   6
         Top             =   810
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   2205
         TabIndex        =   5
         Top             =   270
         Width           =   285
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   765
      TabIndex        =   2
      Top             =   1440
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
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
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Salir"
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
End
Attribute VB_Name = "frm_gen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CONSULTA As DAO.Recordset
Dim nelemen As Integer
Dim cabrev As String
Dim xcta As String
Dim xcc  As String
Dim nsoles As Double, ndolar As Double, cdh As String
Dim g_fecven As Variant, g_ruc As String
Dim ntotdebs As Double
Dim ntothabs As Double
Dim ntotdebd As Double
Dim ntothabd    As Double
Dim ccomproba   As String
Dim celemento   As String
Dim wcosto      As Integer
Dim tbcont_a    As DAO.Recordset
Dim tbcont_a2   As DAO.Recordset
Dim tbcontacop  As DAO.Recordset
Dim dbconta     As DAO.Database
Dim tbconta     As DAO.Recordset
Dim dfechacomp  As Date
Dim wtrans_acum As String
Dim wdetalconta As String * 1
Dim wdebe       As Double
Dim whaber      As Double

Private Sub proceso_acum()

    dbconta.Execute ("delete * from cont_acum")
    Set tbcont_a = dbconta.OpenRecordset("cont_acum")
    tbcont_a.Index = "cont_acum"
    
    If tbconta.RecordCount > 0 Then
        tbconta.MoveFirst
        Do While Not tbconta.EOF
            tbcont_a.Seek "=", tbconta.Fields("f5codcta") & ""
            If tbcont_a.NoMatch Then
                tbcont_a.AddNew
                copia_datos
            Else
                tbcont_a.Edit
            End If
            If "" & tbconta.Fields("f3debhab") = "D" Then
                tbcont_a.Fields("f3importe") = Format(tbcont_a.Fields("f3importe") + tbconta.Fields("f3importe"), "0.00")
                tbcont_a.Fields("f3imported") = Format(tbcont_a.Fields("f3imported") + tbconta.Fields("f3imported"), "0.00")
            Else
                tbcont_a.Fields("f3importe2") = Format(tbcont_a.Fields("f3importe2") + tbconta.Fields("f3importe"), "0.00")
                tbcont_a.Fields("f3imported2") = Format(tbcont_a.Fields("f3imported2") + tbconta.Fields("f3imported"), "0.00")
            End If
            tbcont_a.Update
            tbconta.MoveNext
        Loop
        '------------------------------------------
        dbconta.Execute ("delete * from cont_acumcop")
        Set tbcont_a2 = dbconta.OpenRecordset("cont_acumcop")
        
        tbcont_a.MoveFirst
        Do While Not tbcont_a.EOF
            If tbcont_a.Fields("f3importe2") > 0 Then
                tbcont_a2.AddNew
                tbcont_a2.Fields("F3PROAME") = tbcont_a.Fields("F3PROAME") & ""
                tbcont_a2.Fields("F3COMPRO") = tbcont_a.Fields("F3COMPRO") & ""
                tbcont_a2.Fields("F3ELEMEN") = tbcont_a.Fields("F3ELEMEN") & ""
                tbcont_a2.Fields("F3ORIGEN") = tbcont_a.Fields("F3ORIGEN") & ""
                tbcont_a2.Fields("F3FCHOPR") = tbcont_a.Fields("F3FCHOPR")
                tbcont_a2.Fields("F5CODCTA") = tbcont_a.Fields("F5CODCTA") & ""
                tbcont_a2.Fields("F3DETALL") = tbcont_a.Fields("F3DETALL") & ""
                tbcont_a2.Fields("F3CODGAS") = tbcont_a.Fields("F3CODGAS") & ""
                tbcont_a2.Fields("F3CHEQUE") = tbcont_a.Fields("F3CHEQUE")
                tbcont_a2.Fields("F3NROREF") = tbcont_a.Fields("F3NROREF") & ""
                tbcont_a2.Fields("F3IMPORTE") = tbcont_a.Fields("F3IMPORTE2")
                tbcont_a2.Fields("F3IMPORTED") = tbcont_a.Fields("F3IMPORTED2")
                tbcont_a2.Fields("F3IMPORTE2") = 0#
                tbcont_a2.Fields("F3IMPORTED2") = 0#
                tbcont_a2.Fields("F3DEBHAB") = "H"
                tbcont_a2.Fields("F3MONEDA") = tbcont_a.Fields("F3MONEDA") & ""
                tbcont_a2.Fields("F3TIPCAMBD") = tbcont_a.Fields("F3TIPCAMBD")
                tbcont_a2.Fields("F3CTABANC") = tbcont_a.Fields("F3CTABANC")
                tbcont_a2.Fields("F3MOVBANC") = tbcont_a.Fields("F3MOVBANC")
                tbcont_a2.Fields("F3TIPDOC") = tbcont_a.Fields("F3TIPDOC") & ""
                tbcont_a2.Fields("F3OBRA") = tbcont_a.Fields("F3OBRA") & ""
                tbcont_a2.Fields("F3TRANS") = tbcont_a.Fields("F3TRANS") & ""
                tbcont_a2.Fields("F3DESTINO") = tbcont_a.Fields("F3DESTINO") & ""
                tbcont_a2.Update
                
                tbcont_a.Edit
                tbcont_a.Fields("f3importe2") = 0#
                tbcont_a.Fields("f3imported2") = 0#
                tbcont_a.Fields("f3debhab") = "D"
                tbcont_a.Update
                
            End If
            tbcont_a.MoveNext
        Loop
        If tbcont_a2.RecordCount > 0 Then
            tbcont_a2.MoveFirst
            Do While Not tbcont_a2.EOF
                tbcont_a.AddNew
                tbcont_a.Fields("F3PROAME") = tbcont_a2.Fields("F3PROAME") & ""
                tbcont_a.Fields("F3COMPRO") = tbcont_a2.Fields("F3COMPRO") & ""
                tbcont_a.Fields("F3ELEMEN") = tbcont_a2.Fields("F3ELEMEN") & ""
                tbcont_a.Fields("F3ORIGEN") = tbcont_a2.Fields("F3ORIGEN") & ""
                tbcont_a.Fields("F3FCHOPR") = tbcont_a2.Fields("F3FCHOPR")
                tbcont_a.Fields("F5CODCTA") = tbcont_a2.Fields("F5CODCTA") & ""
                tbcont_a.Fields("F3DETALL") = tbcont_a2.Fields("F3DETALL") & ""
                tbcont_a.Fields("F3CODGAS") = tbcont_a2.Fields("F3CODGAS") & ""
                tbcont_a.Fields("F3CHEQUE") = tbcont_a2.Fields("F3CHEQUE")
                tbcont_a.Fields("F3NROREF") = tbcont_a2.Fields("F3NROREF") & ""
                tbcont_a.Fields("F3IMPORTE") = tbcont_a2.Fields("F3IMPORTE")
                tbcont_a.Fields("F3IMPORTED") = tbcont_a2.Fields("F3IMPORTED")
                tbcont_a.Fields("F3IMPORTE2") = tbcont_a2.Fields("F3IMPORTE2")
                tbcont_a.Fields("F3IMPORTED2") = tbcont_a2.Fields("F3IMPORTED2")
                tbcont_a.Fields("F3DEBHAB") = tbcont_a2.Fields("f3debhab") & ""
                tbcont_a.Fields("F3MONEDA") = tbcont_a2.Fields("F3MONEDA") & ""
                tbcont_a.Fields("F3TIPCAMBD") = tbcont_a2.Fields("F3TIPCAMBD")
                tbcont_a.Fields("F3CTABANC") = tbcont_a2.Fields("F3CTABANC")
                tbcont_a.Fields("F3MOVBANC") = tbcont_a2.Fields("F3MOVBANC")
                tbcont_a.Fields("F3TIPDOC") = tbcont_a2.Fields("F3TIPDOC") & ""
                tbcont_a.Fields("F3OBRA") = tbcont_a2.Fields("F3OBRA") & ""
                tbcont_a.Fields("F3TRANS") = tbcont_a2.Fields("F3TRANS") & ""
                tbcont_a.Fields("F3DESTINO") = tbcont_a2.Fields("F3DESTINO") & ""
                tbcont_a.Update
                tbcont_a2.MoveNext
            Loop
        End If
        tbcont_a2.Close
        '------------------------------------------
        tbcont_a.MoveFirst
        Do While Not tbcont_a.EOF
            If tbcont_a.Fields("f3importe") = 0# And tbcont_a.Fields("f3imported") = 0# Then
                tbcont_a.Delete
            End If
            tbcont_a.MoveNext
        Loop
        '------------------------------------------
        tbcont_a.MoveFirst
        nelemen = 0
        wdebe = 0#: whaber = 0#
        Do While Not tbcont_a.EOF
            nelemen = nelemen + 1
            tbcont_a.Edit
            tbcont_a.Fields("f3origen") = worigen
            tbcont_a.Fields("f3moneda") = "S"
            tbcont_a.Fields("f3compro") = txtmes.Text & "00001"
            tbcont_a.Fields("f3elemen") = Format(nelemen, "0000")
            If tbcont_a.Fields("f3debhab") = "D" Then
                wdebe = wdebe + tbcont_a.Fields("f3importe")
            Else
                whaber = whaber + tbcont_a.Fields("f3importe")
            End If
            tbcont_a.Update
            tbcont_a.MoveNext
        Loop
        '------------------------------------------
        dbconta.Execute ("delete * from contable_cop")
        Set tbcontacop = dbconta.OpenRecordset("contable_cop")
        tbconta.MoveFirst
        Do While Not tbconta.EOF
            tbcontacop.AddNew
            tbcontacop.Fields("F3PROAME") = tbconta.Fields("F3PROAME")
            tbcontacop.Fields("F3COMPRO") = tbconta.Fields("F3COMPRO")
            tbcontacop.Fields("F3ELEMEN") = tbconta.Fields("F3ELEMEN")
            tbcontacop.Fields("F3ORIGEN") = tbconta.Fields("F3ORIGEN")
            tbcontacop.Fields("F3FCHOPR") = tbconta.Fields("F3FCHOPR")
            tbcontacop.Fields("F5CODCTA") = tbconta.Fields("F5CODCTA")
            tbcontacop.Fields("F3DETALL") = tbconta.Fields("F3DETALL")
            tbcontacop.Fields("F3CODGAS") = tbconta.Fields("F3CODGAS")
            tbcontacop.Fields("F3CHEQUE") = tbconta.Fields("F3CHEQUE")
            tbcontacop.Fields("F3NROREF") = tbconta.Fields("F3NROREF")
            tbcontacop.Fields("F3IMPORTE") = tbconta.Fields("F3IMPORTE")
            tbcontacop.Fields("F3IMPORTED") = tbconta.Fields("F3IMPORTED")
            tbcontacop.Fields("F3DEBHAB") = tbconta.Fields("F3DEBHAB")
            tbcontacop.Fields("F3MONEDA") = tbconta.Fields("F3MONEDA")
            tbcontacop.Fields("F3TIPCAMBD") = tbconta.Fields("F3TIPCAMBD")
            tbcontacop.Fields("F3CTABANC") = tbconta.Fields("F3CTABANC")
            tbcontacop.Fields("F3MOVBANC") = tbconta.Fields("F3MOVBANC")
            tbcontacop.Fields("F3TIPDOC") = tbconta.Fields("F3TIPDOC")
            tbcontacop.Fields("F3OBRA") = tbconta.Fields("F3OBRA")
            tbcontacop.Fields("F3TRANS") = tbconta.Fields("F3TRANS")
            tbcontacop.Fields("F3DESTINO") = tbconta.Fields("F3DESTINO")
            tbcontacop.Update
            tbconta.MoveNext
        Loop
        tbcontacop.Close
        '------------------------------------------
        dbconta.Execute ("delete * from contable")
        tbcont_a.MoveFirst
        Do While Not tbcont_a.EOF
            nelemen = nelemen + 1
            tbconta.AddNew
            tbconta.Fields("F3PROAME") = tbcont_a.Fields("F3PROAME")
            tbconta.Fields("F3COMPRO") = tbcont_a.Fields("F3COMPRO")
            tbconta.Fields("F3ELEMEN") = tbcont_a.Fields("F3ELEMEN")
            tbconta.Fields("F3ORIGEN") = tbcont_a.Fields("F3ORIGEN")
            tbconta.Fields("F3FCHOPR") = tbcont_a.Fields("F3FCHOPR")
            tbconta.Fields("F5CODCTA") = tbcont_a.Fields("F5CODCTA")
            tbconta.Fields("F3DETALL") = tbcont_a.Fields("F3DETALL")
            tbconta.Fields("F3CODGAS") = tbcont_a.Fields("F3CODGAS")
            tbconta.Fields("F3CHEQUE") = tbcont_a.Fields("F3CHEQUE")
            tbconta.Fields("F3NROREF") = tbcont_a.Fields("F3NROREF")
            tbconta.Fields("F3IMPORTE") = tbcont_a.Fields("F3IMPORTE")
            tbconta.Fields("F3IMPORTED") = tbcont_a.Fields("F3IMPORTED")
            tbconta.Fields("F3DEBHAB") = tbcont_a.Fields("F3DEBHAB")
            tbconta.Fields("F3MONEDA") = tbcont_a.Fields("F3MONEDA")
            tbconta.Fields("F3TIPCAMBD") = tbcont_a.Fields("F3TIPCAMBD")
            tbconta.Fields("F3CTABANC") = tbcont_a.Fields("F3CTABANC")
            tbconta.Fields("F3MOVBANC") = tbcont_a.Fields("F3MOVBANC")
            tbconta.Fields("F3TIPDOC") = tbcont_a.Fields("F3TIPDOC")
            tbconta.Fields("F3OBRA") = tbcont_a.Fields("F3OBRA")
            tbconta.Fields("F3TRANS") = tbcont_a.Fields("F3TRANS")
            tbconta.Fields("F3DESTINO") = tbcont_a.Fields("F3DESTINO")
            tbconta.Update
            tbcont_a.MoveNext
        Loop
    End If
    tbcont_a.Close
    
End Sub

Private Sub genera()
Dim nsolcop As Double
Dim ndolcop As Double

On Error GoTo Errores

    Me.MousePointer = 11
    
    Set dbcntcont = OpenDatabase(App.Path & "\ctrcom.mdb")
    Set tbcntcont = dbcntcont.OpenRecordset("param_com")
    tbcntcont.Index = "IDCODEMP"
    tbcntcont.Seek "=", wempresa
    If Not tbcntcont.NoMatch Then
        wtrans_acum = tbcntcont.Fields("f1agrupa") & ""
        wdetalconta = tbcntcont.Fields("F1DETALCONTA") & ""
    End If
    tbcntcont.Close
    dbcntcont.Close

    Set dbcntcont = OpenDatabase(wcontacnt & "\cnt_cont.mdb")
    Set tbcntcont = dbcntcont.OpenRecordset("cf1cnt")
    tbcntcont.Index = "cf1cnt"
    tbcntcont.Seek "=", wempresa
    If Not tbcntcont.NoMatch Then
        wcosto = Val(tbcntcont.Fields("f1costo") & "")
    End If
        
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    'Set TbOfiRegis = dbcompras.OpenRecordset("REGISMOV")
    If TbOfiRegis.State = 1 Then TbOfiRegis.Close
    TbOfiRegis.Open "SELECT * FROM REGISMOV", cnn_dbbancos
    
    'Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    If TbDetRegis.State = 1 Then TbDetRegis.Close
    TbDetRegis.Open "SELECT * FROM REGISMOV", cnn_dbbancos
    
    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")

    Set dbconta = OpenDatabase(wrutatemp & "\db_conta.mdb")
    dbconta.Execute ("delete from contable")
    Set tbconta = dbconta.OpenRecordset("contable")
    
    Set dbtabla = OpenDatabase(wrutaconta & "\db_tabla.mdb")
    Set tbcf5costo = dbtabla.OpenRecordset("cf5costo")
    tbcf5costo.Index = "cf5costo"
    Set tbcf5pla = dbtabla.OpenRecordset("cf5pla")
    tbcf5pla.Index = "cf5pla"

    Set CONSULTA = dbcompras.OpenRecordset("Select * From [REGISDOC] Where F4MESMOV = '" + txtanno.Text + txtmes.Text + "' order by F4NUMMOV")
    If CONSULTA.RecordCount > 0 Then
        CONSULTA.MoveFirst
        Do While Not CONSULTA.EOF
            If Len(Trim(CONSULTA.Fields("f4contable") & "")) = 0 Then
                TbDocumento.Index = "idcoddoc"
                TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
                If Not TbDocumento.NoMatch Then
                    If TbDocumento.Fields("f2transfer") & "" <> "N" Then
                        lblcompro.Caption = "Nº Movimiento : " & CONSULTA.Fields("f4nummov") & ""
                        lblcompro.Refresh
                        nelemen = 0
                        ntotdebs = 0#: ntothabs = 0#
                        ntotdebd = 0#: ntothabd = 0#
                        TbOfiRegis.Filter = ""
                        TbOfiRegis.Filter = "F4MESMOV='" & CONSULTA.Fields("f4mesmov") & "' AND F4NUMMOV='" & CONSULTA.Fields("f4nummov") & "'"
                        If Not TbOfiRegis.EOF Then
                            dfechacomp = Format(CONSULTA.Fields("f4fecha"), "dd/mm/yyyy")
                            If Format(Month(CONSULTA.Fields("f4fecha")), "00") <> Format(txtmes.Text, "00") Then
                                dfechacomp = CVDate("01/" & Format(txtmes.Text, "00") & "/" & wanno)
                            End If
                            gen_prov
                            TbDetRegis.Filter = ""
                            TbDetRegis.Filter = "F4MESMOV='" & CONSULTA.Fields("f4mesmov") & "' AND F4NUMMOV='" & CONSULTA.Fields("f4nummov") & "'"
                            If Not TbDetRegis.EOF Then
                                Do While TbDetRegis.Fields("f4mesmov") = CONSULTA.Fields("f4mesmov") And TbDetRegis.Fields("f4nummov") = CONSULTA.Fields("f4nummov") And Not TbDetRegis.EOF
                                    xcta = ""
                                    xcc = ""
                                    gen_det
                                    tbcf5costo.Index = "cf5costo"
                                    tbcf5costo.Seek "=", xcta
                                    If Not tbcf5costo.NoMatch Then
                                        Select Case wcosto
                                            '------- PRORATEO
                                            Case 2:
                                                nsolcop = nsoles
                                                ndolcop = ndolar
                                                Do While tbcf5costo.Fields("f5codcta") = xcta And Not tbcf5costo.EOF
                                                    nsoles = Val(Format(nsolcop * (tbcf5costo.Fields("f5porcen") / 100), "0.00"))
                                                    ndolar = Val(Format(ndolcop * (tbcf5costo.Fields("f5porcen") / 100), "0.00"))
                                                    gen_ccosto
                                                    tbcf5costo.MoveNext
                                                    If tbcf5costo.EOF Then Exit Do
                                                    If tbcf5costo.Fields("f5codcta") <> xcta Then Exit Do
                                                Loop
                                            Case 3:
                                                tbcf5costo.Index = "CF5CTACC"
                                                tbcf5costo.Seek "=", xcta, xcc
                                                If Not tbcf5costo.NoMatch Then
                                                    gen_ccosto
                                                End If
                                            Case Else
                                                gen_ccosto
                                        End Select
                                    End If
                                    TbDetRegis.MoveNext
                                    If TbDetRegis.EOF Then Exit Do
                                    If TbDetRegis.Fields("f4mesmov") <> CONSULTA.Fields("f4mesmov") Or TbDetRegis.Fields("f4nummov") <> CONSULTA.Fields("f4nummov") Then Exit Do
                                Loop
                            End If
                            gen_igv
                            gen_otroimp
                            GEN_FONAVI
                            gen_retencion
                            gen_redsuma
                            gen_redresta
                            gen_dcto
                            '----------------------------------------------
                            '----------- genera redondeo
                            If CONSULTA.Fields("f4moneda") = "S" Then
                                If Format(ntotdebd, "#0.00") <> Format(ntothabd, "#0.00") Then
                                    tbconta.Index = "conta"
                                    tbconta.Seek "=", ccomproba, celemento
                                    If Not tbconta.NoMatch Then
                                        tbconta.Edit
                                        If Format(ntotdebd, "#0.00") > Format(ntothabd, "#0.00") Then
                                            If tbconta.Fields("f3debhab") = "H" Then
                                                tbconta.Fields("f3imported") = tbconta.Fields("f3imported") + Format(ntotdebd - ntothabd, "#0.00")
                                            Else
                                                tbconta.Fields("f3imported") = tbconta.Fields("f3imported") - Format(ntotdebd - ntothabd, "#0.00")
                                            End If
                                        Else
                                            If tbconta.Fields("f3debhab") = "H" Then
                                                tbconta.Fields("f3imported") = tbconta.Fields("f3imported") - Abs(Format(ntotdebd - ntothabd, "#0.00"))
                                            Else
                                                tbconta.Fields("f3imported") = tbconta.Fields("f3imported") + Abs(Format(ntotdebd - ntothabd, "#0.00"))
                                            End If
                                        End If
                                        tbconta.Update
                                    End If
                                End If
                            Else
                                If Format(ntotdebs, "#0.00") <> Format(ntothabs, "#0.00") Then
                                    tbconta.Index = "conta"
                                    tbconta.Seek "=", ccomproba, celemento
                                    If Not tbconta.NoMatch Then
                                        tbconta.Edit
                                        If Format(ntotdebs, "#0.00") > Format(ntothabs, "#0.00") Then
                                            If tbconta.Fields("f3debhab") = "H" Then
                                                tbconta.Fields("f3importe") = tbconta.Fields("f3importe") + Format(ntotdebs - ntothabs, "#0.00")
                                            Else
                                                tbconta.Fields("f3importe") = tbconta.Fields("f3importe") - Format(ntotdebs - ntothabs, "#0.00")
                                            End If
                                        Else
                                            If tbconta.Fields("f3debhab") = "H" Then
                                                tbconta.Fields("f3importe") = tbconta.Fields("f3importe") - Abs(Format(ntotdebs - ntothabs, "#0.00"))
                                            Else
                                                tbconta.Fields("f3importe") = tbconta.Fields("f3importe") + Abs(Format(ntotdebs - ntothabs, "#0.00"))
                                            End If
                                        End If
                                        tbconta.Update
                                    End If
                                End If
                            End If
                            '----------------------------------------------
                        End If
                    End If
                End If
            End If
            CONSULTA.MoveNext
            If CONSULTA.EOF Then Exit Do
        Loop
        'If wtrans_acum = "*" Then
        '    If MsgBox("Desea proceder a acumular los asientos contables ?", 36, "Atención") = 6 Then
        '        proceso_acum
        '    End If
        'End If
        MsgBox "Fin del proceso.", 48, "Compras"
        frm_cons.Show 1
    Else
        MsgBox "No hay registros en el mes de proceso.", 48, "Compras"
    End If
    Me.MousePointer = 1

    tbcf5pla.Close
    tbcf5costo.Close
    dbtabla.Close

    tbconta.Close
    dbconta.Close

    TbDocumento.Close
    dbcomtabla.Close

    TbDetRegis.Close
    TbOfiRegis.Close
    TbCabRegis.Close
    CONSULTA.Close
    dbcompras.Close
    
    tbcntcont.Close
    dbcntcont.Close

    Exit Sub

Errores:
    If Err = 3219 Then
        Resume Next
    Else
        If Err = 3043 Then
            MsgBox "Presione Ctrl+Alt+Del y Enter, Enter"
        Else
            MsgBox "Ha ocurrido el error" & Err & " no se pueden generar los asientos. Llame a su proveedor de software.", 48, "Compras"
'            Resume
            Exit Sub
        End If
    End If

End Sub

Private Sub gen_retencion()
Dim cabrev As String
On Error GoTo error44

    If CONSULTA.Fields("f4montoret") > 0 Then
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
        
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "RETENCION 4TA CAT."
        End If
        
        tbconta.Fields("f5codcta") = wctaret
        tbconta.Fields("f3codgas") = "4TA"
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            ntothabs = ntothabs + Format(CONSULTA.Fields("f4montoret") * CONSULTA.Fields("f4tipcam"), "#0.00")
            tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4montoret") * CONSULTA.Fields("f4tipcam"), "#0.00")
            tbconta.Fields("f3imported") = CONSULTA.Fields("f4montoret")
        Else
            tbconta.Fields("f3importe") = CONSULTA.Fields("f4montoret")
            If CONSULTA.Fields("f4tipcam") <> 0 Then
                ntothabd = ntothabd + Format(CONSULTA.Fields("f4montoret") / CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4montoret") / CONSULTA.Fields("f4tipcam"), "#0.00")
            Else
                tbconta.Fields("f3imported") = 0
            End If
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        cabrev = ""
        TbDocumento.Index = "IDCODDOC"
        TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
        If Not TbDocumento.NoMatch Then
            cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
        End If
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(cabrev = "Cre", "D", "H")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Update
    End If
    Exit Sub

error44:
   MsgBox "Ha ocurrido el error: " & Err
   Exit Sub

End Sub

Private Sub gen_redsuma()
On Error GoTo error_suma
Dim cabrev As String

    cabrev = ""
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
    If Not TbDocumento.NoMatch Then
       cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
    End If
    
    If CONSULTA.Fields("f4redsuma") <> 0 Then
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
        tbconta.Fields("f3fECCOM") = right(Year(dfechacomp), 2) & Format(Month(dfechacomp), "00") & Format(Day(dfechacomp), "00")
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "Redondeo"
        End If
        
        tbconta.Fields("f5codcta") = wredsuma
        tbconta.Fields("f3codgas") = ""
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            If UCase(cabrev) = "CRE" Then
                ntothabs = ntothabs + Format(CONSULTA.Fields("f4redsuma") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4redsuma") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4redsuma")
                nsoles = Format(CONSULTA.Fields("f4redsuma") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4redsuma")
            Else
                ntotdebs = ntotdebs + Format(CONSULTA.Fields("f4redsuma") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4redsuma") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4redsuma")
                nsoles = Format(CONSULTA.Fields("f4redsuma") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4redsuma")
            End If
        Else
            If UCase(cabrev) = "CRE" Then
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4redsuma")
                nsoles = CONSULTA.Fields("f4redsuma")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntothabd = ntothabd + Format(CONSULTA.Fields("f4redsuma") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4redsuma") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4redsuma") / CONSULTA.Fields("f4tipcam"), "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                    ndolar = 0
                End If
            Else
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4redsuma")
                nsoles = CONSULTA.Fields("f4redsuma")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntotdebd = ntotdebd + Format(CONSULTA.Fields("f4redsuma") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4redsuma") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4redsuma") / CONSULTA.Fields("f4tipcam"), "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                    ndolar = 0
                End If
            End If
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "H", "D")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Update
        '--------------------------
        cdh = IIf(UCase(cabrev) = "CRE", "H", "D")
        xcta = wredsuma
        tbcf5costo.Index = "CF5COSTO"
        tbcf5costo.Seek "=", xcta
        If Not tbcf5costo.NoMatch Then
            gen_ccosto
        End If
        xcta = ""
        '--------------------------
    End If
   
    Exit Sub

error_suma:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub

End Sub

Private Sub gen_redresta()
On Error GoTo error_resta
Dim cabrev As String

    cabrev = ""
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
    If Not TbDocumento.NoMatch Then
       cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
    End If
    
    If CONSULTA.Fields("f4redresta") <> 0 Then
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
        tbconta.Fields("f3fECCOM") = right(Year(dfechacomp), 2) & Format(Month(dfechacomp), "00") & Format(Day(dfechacomp), "00")
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "Redondeo"
        End If
        
        tbconta.Fields("f5codcta") = wredresta
        tbconta.Fields("f3codgas") = ""
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            If UCase(cabrev) <> "CRE" Then
                ntothabs = ntothabs + Format(CONSULTA.Fields("f4redresta") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4redresta") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4redresta")
                nsoles = Format(CONSULTA.Fields("f4redresta") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4redresta")
            Else
                ntotdebs = ntotdebs + Format(CONSULTA.Fields("f4redresta") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4redresta") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4redresta")
                nsoles = Format(CONSULTA.Fields("f4redresta") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4redresta")
            End If
        Else
            If UCase(cabrev) <> "CRE" Then
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4redresta")
                nsoles = CONSULTA.Fields("f4redresta")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntothabd = ntothabd + Format(CONSULTA.Fields("f4redresta") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4redresta") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4redresta") / CONSULTA.Fields("f4tipcam"), "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                    ndolar = 0
                End If
            Else
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4redresta")
                nsoles = CONSULTA.Fields("f4redresta")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntotdebd = ntotdebd + Format(CONSULTA.Fields("f4redresta") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4redresta") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4redresta") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                    ndolar = 0
                End If
            End If
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "D", "H")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Update
        
        '--------------------------
        cdh = IIf(UCase(cabrev) = "CRE", "D", "H")
        xcta = wredresta
        tbcf5costo.Index = "CF5COSTO"
        tbcf5costo.Seek "=", xcta
        If Not tbcf5costo.NoMatch Then
            gen_ccosto
        End If
        xcta = ""
        '--------------------------
    End If

    Exit Sub

error_resta:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub

End Sub

Private Sub gen_prov()
On Error GoTo ERROR2
Dim cabrev As String

    g_ruc = CONSULTA.Fields("F4RUCPRV") & ""
    g_fecven = CONSULTA.Fields("F4FECVEN")
    
    cabrev = ""
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
    If Not TbDocumento.NoMatch Then
        cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
    End If
    
    nelemen = nelemen + 1
    tbconta.AddNew
    tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
    If wf1numera = "*" Then
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        ccomproba = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
    Else
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        ccomproba = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
    End If
    tbconta.Fields("f3elemen") = Format(nelemen, "0000")
    celemento = Format(nelemen, "0000")
    tbconta.Fields("f3origen") = Format(worigen, "00")
    tbconta.Fields("f3fchopr") = dfechacomp
    tbconta.Fields("f3fECCOM") = right(Year(dfechacomp), 2) & Format(Month(dfechacomp), "00") & Format(Day(dfechacomp), "00")
    tbconta.Fields("f3fchvcto") = g_fecven
    tbconta.Fields("f3fECven") = right(Year(g_fecven), 2) & Format(Month(g_fecven), "00") & Format(Day(g_fecven), "00")
    
    Rem NSE tbconta.Fields("f3detall") = CONSULTA.Fields("f4nomprv") & ""
    tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
    tbconta.Fields("f5codcta") = CONSULTA.Fields("f4ctacont") & ""
    tbconta.Fields("f3codgas") = CONSULTA.Fields("F4GRUPO") & ""
    tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
    tbconta.Fields("f3nroref") = CONSULTA.Fields("F4SERDOC") & "-" & Val(Trim(Str(Val(CONSULTA.Fields("f4numdoc") & ""))))
    If CONSULTA.Fields("f4moneda") = "D" Then
        If UCase(cabrev) = "CRE" Then
            ntotdebs = ntotdebs + Format(CONSULTA.Fields("f4total") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
            tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4total") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
            tbconta.Fields("f3imported") = CONSULTA.Fields("f4total") * -1
        Else
            ntothabs = ntothabs + Format(CONSULTA.Fields("f4total") * CONSULTA.Fields("f4tipcam"), "#0.00")
            tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4total") * CONSULTA.Fields("f4tipcam"), "#0.00")
            tbconta.Fields("f3imported") = CONSULTA.Fields("f4total")
        End If
    Else
        If UCase(cabrev) = "CRE" Then
            tbconta.Fields("f3importe") = CONSULTA.Fields("f4total") * -1
            If CONSULTA.Fields("f4tipcam") <> 0 Then
                ntotdebd = ntotdebd + Format(CONSULTA.Fields("f4total") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4total") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
            Else
                tbconta.Fields("f3imported") = 0
            End If
        Else
            tbconta.Fields("f3importe") = CONSULTA.Fields("f4total")
            If CONSULTA.Fields("f4tipcam") <> 0 Then
                ntothabd = ntothabd + Format(CONSULTA.Fields("f4total") / CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4total") / CONSULTA.Fields("f4tipcam"), "#0.00")
            Else
                tbconta.Fields("f3imported") = 0
            End If
        End If
    End If
    tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
    tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
    
    tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
    tbconta.Fields("f3tipdoc") = cabrev
    tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "D", "H")
    'NUEVO PARA GML----------------------
    tbconta.Fields("F3RUC") = g_ruc
    tbconta.Fields("F3AUTOMATICO") = False
    tbconta.Fields("F3ORIGAUTO") = False
    tbconta.Fields("F4FECVENC") = g_fecven
    tbconta.Fields("F3DESTINO") = "P"
    tbconta.Fields("f3serdoc") = Val(CONSULTA.Fields("f4serdoc") & "")
    'FIN DE NUEVO PARA GML-----------------------
    tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
    tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
    tbconta.Update
    Exit Sub

ERROR2:
    MsgBox "Ha ocurrido el error: " & Err
    Resume
    Exit Sub

End Sub

Private Sub gen_otroimp()
On Error GoTo error3
Dim cabrev As String
                    
    If CONSULTA.Fields("f4otrimp") <> 0 Then
        nsoles = 0#:  ndolar = 0#
        cabrev = ""
        TbDocumento.Index = "IDCODDOC"
        TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
        If Not TbDocumento.NoMatch Then
            cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
        End If
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
                
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "OTROS IMPUESTOS"
        End If
        
        tbconta.Fields("f5codcta") = wctaotros
        tbconta.Fields("f3codgas") = "OTR"
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            If UCase(cabrev) = "CRE" Then
                ntothabs = ntothabs + Format(CONSULTA.Fields("f4otrimp") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4otrimp") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4otrimp") * -1
                nsoles = Format(CONSULTA.Fields("f4otrimp") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                ndolar = CONSULTA.Fields("f4otrimp") * -1
            Else
                ntotdebs = ntotdebs + Format(CONSULTA.Fields("f4otrimp") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4otrimp") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4otrimp")
                nsoles = Format(CONSULTA.Fields("f4otrimp") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4otrimp")
            End If
        Else
            If UCase(cabrev) = "CRE" Then
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4otrimp") * -1
                nsoles = CONSULTA.Fields("f4otrimp") * -1
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntothabd = ntothabd + Format(CONSULTA.Fields("f4otrimp") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4otrimp") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4otrimp") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                End If
                Else
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4otrimp")
                nsoles = CONSULTA.Fields("f4otrimp")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntotdebd = ntotdebd + Format(CONSULTA.Fields("f4otrimp") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4otrimp") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4otrimp") / CONSULTA.Fields("f4tipcam"), "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                End If
            End If
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
        tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "H", "D")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Update
        '--------------------------
        cdh = IIf(UCase(cabrev) = "CRE", "H", "D")
        xcta = wctaotros
        tbcf5costo.Index = "CF5COSTO"
        tbcf5costo.Seek "=", xcta
        If Not tbcf5costo.NoMatch Then
            gen_ccosto
        End If
        xcta = ""
        '--------------------------
    End If
    Exit Sub

error3:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub

End Sub

Private Sub gen_igv()
On Error GoTo error5
Dim cabrev As String

    cabrev = ""
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
    If Not TbDocumento.NoMatch Then
       cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
    End If
    
    If CONSULTA.Fields("f4igv") <> 0 Then
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
        tbconta.Fields("f3fECCOM") = right(Year(dfechacomp), 2) & Format(Month(dfechacomp), "00") & Format(Day(dfechacomp), "00")
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "I.G.V."
        End If
            
        tbconta.Fields("f5codcta") = wctaigv
        tbconta.Fields("f3codgas") = "IGV"
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            If UCase(cabrev) = "CRE" Then
                ntothabs = ntothabs + Format(CONSULTA.Fields("f4igv") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4igv") * CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4igv") * -1
            Else
                ntotdebs = ntotdebs + Format(CONSULTA.Fields("f4igv") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4igv") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4igv")
            End If
        Else
            If UCase(cabrev) = "CRE" Then
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4igv") * -1
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntothabd = ntothabd + Format(CONSULTA.Fields("f4igv") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4igv") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                End If
            Else
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4igv")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntotdebd = ntotdebd + Format(CONSULTA.Fields("f4igv") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4igv") / CONSULTA.Fields("f4tipcam"), "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                End If
            End If
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
        tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "H", "D")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Fields("f3serdoc") = Val(CONSULTA.Fields("f4serdoc") & "")
        tbconta.Update
    End If
    Exit Sub

error5:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub

End Sub

Private Sub GEN_FONAVI()
Dim cabrev As String

On Error GoTo error33

    If CONSULTA.Fields("f4FONAVI") > 0 Then
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
        
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "IMP. EXTR. SOLIDARIDAD"
        End If
        
        tbconta.Fields("f5codcta") = wctafon
        tbconta.Fields("f3codgas") = "FON"
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            ntothabs = ntothabs + Format(CONSULTA.Fields("f4fonavi") * CONSULTA.Fields("f4tipcam"), "#0.00")
            tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4fonavi") * CONSULTA.Fields("f4tipcam"), "#0.00")
            tbconta.Fields("f3imported") = CONSULTA.Fields("f4fonavi")
        Else
            tbconta.Fields("f3importe") = CONSULTA.Fields("f4fonavi")
            tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4fonavi") / CONSULTA.Fields("f4tipcam"), "#0.00")
            ntothabd = ntothabd + Format(CONSULTA.Fields("f4fonavi") / CONSULTA.Fields("f4tipcam"), "#0.00")
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
        tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        cabrev = ""
        TbDocumento.Index = "IDCODDOC"
        TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
        If Not TbDocumento.NoMatch Then
            cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
        End If
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "D", "H")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Update
    End If
    Exit Sub

error33:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub
  
End Sub

Private Sub gen_det()
On Error GoTo error4

    cabrev = ""
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
    If Not TbDocumento.NoMatch Then
        cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
    End If
    
    nelemen = nelemen + 1
    tbconta.AddNew
    tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
    If wf1numera = "*" Then
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
    Else
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
    End If
    tbconta.Fields("f3afecto") = TbDetRegis.Fields("f3afecto") & ""
    tbconta.Fields("f3elemen") = Format(nelemen, "0000")
    tbconta.Fields("f3origen") = Format(worigen, "00")
    tbconta.Fields("f3fchopr") = dfechacomp
    tbconta.Fields("f3fECCOM") = right(Year(dfechacomp), 2) & Format(Month(dfechacomp), "00") & Format(Day(dfechacomp), "00")
    If wdetalconta = "*" Then
        tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
    Else
        tbconta.Fields("f3detall") = left(TbDetRegis.Fields("f3concepto") & "", 100)
    End If
    
    tbconta.Fields("f5codcta") = TbDetRegis.Fields("f3ctacon") & ""
    xcta = TbDetRegis.Fields("f3ctacon") & ""
    xcc = TbDetRegis.Fields("f3cencos") & ""
    tbconta.Fields("f3codgas") = TbDetRegis.Fields("f3gasto") & ""
    tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
    tbconta.Fields("f3nroref") = "" & TbDetRegis.Fields("F3ORDEN")
    If CONSULTA.Fields("f4moneda") = "S" Then
        tbconta.Fields("f3importe") = Abs(TbDetRegis.Fields("f3importe"))
        If CONSULTA.Fields("f4tipcam") <> 0 Then
            tbconta.Fields("f3imported") = Format(Abs(TbDetRegis.Fields("f3importe")) / CONSULTA.Fields("f4tipcam"), "#0.00")
        Else
            tbconta.Fields("f3imported") = 0
        End If
        nsoles = Abs(TbDetRegis.Fields("f3importe"))
        If CONSULTA.Fields("f4tipcam") <> 0 Then
            ndolar = Format(Abs(TbDetRegis.Fields("f3importe")) / CONSULTA.Fields("f4tipcam"), "#0.00")
        Else
            ndolar = 0
        End If
    Else
        tbconta.Fields("f3imported") = Abs(TbDetRegis.Fields("f3importe"))
        tbconta.Fields("f3importe") = Format(Abs(TbDetRegis.Fields("f3importe")) * CONSULTA.Fields("f4tipcam"), "#0.00")
        ndolar = Abs(TbDetRegis.Fields("f3importe"))
        nsoles = Format(Abs(TbDetRegis.Fields("f3importe")) * CONSULTA.Fields("f4tipcam"), "#0.00")
    End If
    tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
    tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
    tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
    tbconta.Fields("f3tipdoc") = cabrev
    tbconta.Fields("f3debhab") = IIf(cabrev = "Cre", "H", "D")
    cdh = IIf(UCase(cabrev) = "CRE", "H", "D")
    tbconta.Fields("f3costo") = ObtenerCampo("CENTROS", "CCONCAR", "F3COSTO", TbDetRegis.Fields("f3cencos") & "", "T", cnn_dbbancos)
    'NUEVO PARA GML----------------------
    tbconta.Fields("F3RUC") = g_ruc
    tbconta.Fields("F3AUTOMATICO") = False
    tbcf5costo.Index = "CF5COSTO"
    tbcf5costo.Seek "=", xcta
    If Not tbcf5costo.NoMatch Then
        tbconta.Fields("F3ORIGAUTO") = True
    Else
        tbconta.Fields("F3ORIGAUTO") = False
    End If
    tbconta.Fields("F4FECVENC") = g_fecven
    'FIN DE NUEVO PARA GML-----------------------
    tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
    tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
    tbconta.Update
    '------------ acumula para redondeo
    If CONSULTA.Fields("f4moneda") = "S" Then
        If cdh = "D" Then
            ntotdebd = ntotdebd + ndolar
        Else
            ntothabd = ntothabd + ndolar
        End If
    Else
        If cdh = "D" Then
            ntotdebs = ntotdebs + nsoles
        Else
            ntothabs = ntothabs + nsoles
        End If
    End If
    '---------------------------------------------
    Exit Sub

error4:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub

End Sub

Private Sub gen_dcto()
On Error GoTo error_dcto
Dim cabrev As String

    cabrev = ""
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", CONSULTA.Fields("f4tipdoc") & ""
    If Not TbDocumento.NoMatch Then
        cabrev = TbDocumento.Fields("F2ABREV_CONCAR") & ""
    End If
    
    If CONSULTA.Fields("f4dcto") <> 0 Then
        nelemen = nelemen + 1
        tbconta.AddNew
        tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
        If wf1numera = "*" Then
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
        Else
            tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
        End If
        tbconta.Fields("f3elemen") = Format(nelemen, "0000")
        tbconta.Fields("f3origen") = Format(worigen, "00")
        tbconta.Fields("f3fchopr") = dfechacomp
        If wdetalconta = "*" Then
            tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
        Else
            tbconta.Fields("f3detall") = "Descuentos"
        End If
        
        tbconta.Fields("f5codcta") = wdcto
        tbconta.Fields("f3codgas") = ""
        tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
        tbconta.Fields("f3nroref") = ""
        If CONSULTA.Fields("f4moneda") = "D" Then
            If UCase(cabrev) <> "CRE" Then
                ntothabs = ntothabs + Format(CONSULTA.Fields("f4dcto") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4dcto") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4dcto")
                nsoles = Format(CONSULTA.Fields("f4dcto") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4dcto")
            Else
                ntotdebs = ntotdebs + Format(CONSULTA.Fields("f4dcto") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3importe") = Format(CONSULTA.Fields("f4dcto") * CONSULTA.Fields("f4tipcam"), "#0.00")
                tbconta.Fields("f3imported") = CONSULTA.Fields("f4dcto")
                nsoles = Format(CONSULTA.Fields("f4dcto") * CONSULTA.Fields("f4tipcam"), "#0.00")
                ndolar = CONSULTA.Fields("f4dcto")
            End If
        Else
            If UCase(cabrev) <> "CRE" Then
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4dcto")
                nsoles = CONSULTA.Fields("f4dcto")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntothabd = ntothabd + Format(CONSULTA.Fields("f4dcto") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4dcto") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4dcto") / CONSULTA.Fields("f4tipcam"), "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                    ndolar = 0
                End If
            Else
                tbconta.Fields("f3importe") = CONSULTA.Fields("f4dcto")
                nsoles = CONSULTA.Fields("f4dcto")
                If CONSULTA.Fields("f4tipcam") <> 0 Then
                    ntotdebd = ntotdebd + Format(CONSULTA.Fields("f4dcto") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    tbconta.Fields("f3imported") = Format(CONSULTA.Fields("f4dcto") / CONSULTA.Fields("f4tipcam"), "#0.00")
                    ndolar = Format(CONSULTA.Fields("f4dcto") / CONSULTA.Fields("f4tipcam") * -1, "#0.00")
                Else
                    tbconta.Fields("f3imported") = 0
                    ndolar = 0
                End If
            End If
        End If
        tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda")
        tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
        tbconta.Fields("f3tipdoc") = cabrev
        tbconta.Fields("f3debhab") = IIf(UCase(cabrev) = "CRE", "D", "H")
        'NUEVO PARA GML----------------------
        tbconta.Fields("F3RUC") = g_ruc
        tbconta.Fields("F3AUTOMATICO") = False
        tbconta.Fields("F3ORIGAUTO") = False
        tbconta.Fields("F4FECVENC") = g_fecven
        'FIN DE NUEVO PARA GML-----------------------
        tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
        tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
        tbconta.Update
    End If

    Exit Sub

error_dcto:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub

End Sub

Private Sub gen_ccosto()
Dim cnomcta As String, cabrev As String
Dim ccosto As String

On Error GoTo error_cc

    tbcf5pla.Seek "=", tbcf5costo.Fields("f5ccosto1") & ""
    cnomcta = ""
    If Not tbcf5pla.NoMatch Then
        cnomcta = tbcf5pla.Fields("f5nomcta") & ""
    End If

    nelemen = nelemen + 1
    tbconta.AddNew
    tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
    If wf1numera = "*" Then
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
    Else
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
    End If
    tbconta.Fields("f3elemen") = Format(nelemen, "0000")
    tbconta.Fields("f3origen") = Format(worigen, "00")
    tbconta.Fields("f3fchopr") = dfechacomp
        
    If wdetalconta = "*" Then
        tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
    Else
        tbconta.Fields("f3detall") = cnomcta
    End If
    
    tbconta.Fields("f5codcta") = tbcf5costo.Fields("f5ccosto1") & ""
    tbconta.Fields("f3codgas") = ""
    tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
    tbconta.Fields("f3nroref") = ""
    tbconta.Fields("f3importe") = nsoles
    tbconta.Fields("f3imported") = ndolar
    tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
    tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
    tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
    tbconta.Fields("f3tipdoc") = cabrev
    tbconta.Fields("f3debhab") = cdh
    tbconta.Fields("f3costo") = ccosto
    'NUEVO PARA GML----------------------
    tbconta.Fields("F3RUC") = ""
    tbconta.Fields("F3AUTOMATICO") = True
    tbconta.Fields("F3ORIGAUTO") = False
    tbconta.Fields("F4FECVENC") = g_fecven
    'FIN DE NUEVO PARA GML-----------------------
    tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
    tbconta.Fields("F3DESTINO") = "*"
    tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
    tbconta.Update
    '----------- acumula para redondeo
    If CONSULTA.Fields("f4moneda") = "S" Then
        If cdh = "D" Then
            ntotdebd = ntotdebd + ndolar
        Else
            ntothabd = ntothabd + ndolar
        End If
    Else
        If cdh = "D" Then
            ntotdebs = ntotdebs + nsoles
        Else
            ntothabs = ntothabs + nsoles
        End If
    End If
    
    '---------------------------------------------------------------------
    tbcf5pla.Seek "=", tbcf5costo.Fields("f5ccosto2")
    cnomcta = ""
    If Not tbcf5pla.NoMatch Then
        cnomcta = tbcf5pla.Fields("f5nomcta") & ""
    End If
    
    nelemen = nelemen + 1
    tbconta.AddNew
    tbconta.Fields("f3proame") = txtanno.Text & txtmes.Text
    If wf1numera = "*" Then
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & left(CONSULTA.Fields("f4nummov"), 2) & right(CONSULTA.Fields("f4nummov"), 3)
    Else
        tbconta.Fields("f3compro") = Format(txtmes.Text, "00") & right(CONSULTA.Fields("f4nummov"), 5)
    End If
    tbconta.Fields("f3elemen") = Format(nelemen, "0000")
    tbconta.Fields("f3origen") = Format(worigen, "00")
    tbconta.Fields("f3fchopr") = dfechacomp
    
    If wdetalconta = "*" Then
        tbconta.Fields("f3detall") = CONSULTA.Fields("f4refere") & ""
    Else
        tbconta.Fields("f3detall") = cnomcta
    End If
    
    tbconta.Fields("f5codcta") = tbcf5costo.Fields("f5ccosto2") & ""
    tbconta.Fields("f3codgas") = ""
    tbconta.Fields("f3cheque") = Val(CONSULTA.Fields("f4numdoc") & "")
    tbconta.Fields("f3nroref") = ""
    tbconta.Fields("f3importe") = nsoles
    tbconta.Fields("f3imported") = ndolar
    tbconta.Fields("f3moneda") = CONSULTA.Fields("f4moneda") & ""
    tbconta.Fields("f3codmon") = IIf(CONSULTA.Fields("f4moneda") & "" = "S", "MN", "ME")
    tbconta.Fields("f3tipcambd") = CONSULTA.Fields("f4tipcam")
    tbconta.Fields("f3tipdoc") = cabrev
    tbconta.Fields("f3debhab") = IIf(cdh = "D", "H", "D")
    tbconta.Fields("f3costo") = ccosto
    'NUEVO PARA GML----------------------
    tbconta.Fields("F3RUC") = ""
    tbconta.Fields("F3AUTOMATICO") = True
    tbconta.Fields("F3ORIGAUTO") = False
    tbconta.Fields("F4FECVENC") = g_fecven
    'FIN DE NUEVO PARA GML-----------------------
    tbconta.Fields("f3nummov") = CONSULTA.Fields("f4mesmov") & CONSULTA.Fields("f4nummov")
    tbconta.Fields("F3DESTINO") = "*"
    tbconta.Fields("f2tipdoc") = CONSULTA.Fields("f4tipdoc") & ""
    tbconta.Update
    '---------------------------------------------------------------------
    '----------- acumula para redondeo
    If CONSULTA.Fields("f4moneda") = "S" Then
        If cdh = "D" Then
            ntothabd = ntothabd + ndolar
        Else
            ntotdebd = ntotdebd + ndolar
        End If
    Else
        If cdh = "D" Then
            ntothabs = ntothabs + nsoles
        Else
            ntotdebs = ntotdebs + nsoles
        End If
    End If
    '--------------------------------------------------
    
    Exit Sub

error_cc:
    MsgBox "Ha ocurrido el error: " & Err
    Exit Sub
    
End Sub

Private Sub copia_datos()

    tbcont_a.Fields("F3PROAME") = tbconta.Fields("F3PROAME") & ""
    tbcont_a.Fields("F3COMPRO") = tbconta.Fields("F3COMPRO") & ""
    tbcont_a.Fields("F3ELEMEN") = tbconta.Fields("F3ELEMEN") & ""
    tbcont_a.Fields("F3ORIGEN") = tbconta.Fields("F3ORIGEN") & ""
    tbcont_a.Fields("F3FCHOPR") = tbconta.Fields("F3FCHOPR")
    
    tbcont_a.Fields("F5CODCTA") = tbconta.Fields("F5CODCTA") & ""
    tbcf5pla.Index = "CF5PLA"
    tbcf5pla.Seek "=", tbconta.Fields("F5CODCTA") & ""
    If Not tbcf5pla.NoMatch Then
       tbcont_a.Fields("F3DETALL") = tbcf5pla.Fields("f5NOMCTA") & ""
    End If
    Rem tbcont_a.Fields("F3CODGAS") = tbconta.Fields("F3CODGAS") & ""
    Rem tbcont_a.Fields("F3CHEQUE") = tbconta.Fields("F3CHEQUE")
    Rem tbcont_a.Fields("F3NROREF") = tbconta.Fields("F3NROREF") & ""
    tbcont_a.Fields("F3DEBHAB") = tbconta.Fields("F3DEBHAB") & ""
    tbcont_a.Fields("F3MONEDA") = tbconta.Fields("F3MONEDA") & ""
    tbcont_a.Fields("F3TIPCAMBD") = tbconta.Fields("F3TIPCAMBD")
    tbcont_a.Fields("F3CTABANC") = tbconta.Fields("F3CTABANC")
    tbcont_a.Fields("F3MOVBANC") = tbconta.Fields("F3MOVBANC")
    Rem tbcont_a.Fields("F3TIPDOC") = tbconta.Fields("F3TIPDOC") & ""
    Rem tbcont_a.Fields("F3OBRA") = tbconta.Fields("F3OBRA") & ""
    tbcont_a.Fields("F3TRANS") = tbconta.Fields("F3TRANS") & ""
    tbcont_a.Fields("F3DESTINO") = tbconta.Fields("F3DESTINO") & ""

End Sub

Private Sub cmdresp_Click(Index As Integer)

    Select Case Index
        Case 0:
            genera
            cmdresp(1).SetFocus
        Case 1:
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

    txtmes.Text = Format(mes, "00")
    txtanno.Text = wanno

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtmes.Text) >= 1 And Val(txtmes.Text) <= 12 Then
            txtmes.Text = Format(txtmes.Text, "00")
            cmdresp(0).SetFocus
        End If
    End If

End Sub
