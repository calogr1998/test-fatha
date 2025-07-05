VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmtrans 
   Appearance      =   0  'Flat
   Caption         =   "Transferencia de Asientos Contables"
   ClientHeight    =   1875
   ClientLeft      =   3735
   ClientTop       =   2145
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1875
   ScaleWidth      =   4335
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   0
      Left            =   855
      TabIndex        =   2
      Top             =   1395
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      ScaleHeight     =   720
      ScaleWidth      =   4140
      TabIndex        =   4
      Top             =   90
      Width           =   4200
      Begin VB.TextBox txtanno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   600
      End
      Begin VB.TextBox txtmes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   0
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2295
         TabIndex        =   6
         Top             =   315
         Width           =   300
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   585
         TabIndex        =   5
         Top             =   315
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdresp 
      Height          =   420
      Index           =   1
      Left            =   2295
      TabIndex        =   3
      Top             =   1395
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.Label lblcuenta 
      Appearance      =   0  'Flat
      Caption         =   "Procesando Comprobante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   630
      TabIndex        =   7
      Top             =   945
      Width           =   3030
   End
End
Attribute VB_Name = "frmtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbconta     As DAO.Database
Dim tbconta     As DAO.Recordset
Dim CONSULTA    As DAO.Recordset
Dim nelemen     As Integer
Dim cabrev      As String
Dim ccuenta     As String
Dim nref        As String
Dim cdh         As String

Private Sub cmdresp_Click(Index As Integer)

   Select Case Index
      Case 0:
         transfiere
         MsgBox "Fin del proceso", 48, "Compras"
         cmdresp(1).SetFocus
      Case 1:
         Unload Me
   End Select

End Sub

Private Sub Form_Load()

   txtmes.Text = Format(mes, "00")
   txtanno.Text = wanno

End Sub

Private Sub reg_analisis()

   Do While Not tbconta.EOF
      lblcuenta.Caption = "Procesando Cuenta : " & tbconta.Fields("f5codcta")
      lblcuenta.Refresh
      If Len(Trim(tbconta.Fields("f3nroref"))) <> 0 Then
         ccuenta = tbconta.Fields("f5codcta")
         nref = tbconta.Fields("f3nroref")
         cdh = tbconta.Fields("f3debhab")
         tbcf9saldo.Seek "=", ccuenta, nref
         If Not tbcf9saldo.NoMatch Then
            tbcf5pla.Seek "=", ccuenta
            If Not tbcf5pla.NoMatch Then
               If tbcf5pla.Fields("f5moneda") = "D" Then
                  If tbcf9saldo.Fields("f9saldod") = 0 And "" & Len(Trim(tbcf9saldo.Fields("f9mesd") & "")) <> 0 And tbcf9saldo.Fields("f9mesd") < mes Then
                     tbcf9saldo.Edit
                     tbcf9saldo.Fields("f9mesd") = "  "
                     tbcf9saldo.Update
                  End If
               Else 'soles
                  If tbcf9saldo.Fields("f9saldo") = 0 And "" & Len(Trim(tbcf9saldo.Fields("f9mes") & "")) <> 0 And tbcf9saldo.Fields("f9mes") < mes Then
                     tbcf9saldo.Edit
                     tbcf9saldo.Fields("f9mes") = "  "
                     tbcf9saldo.Update
                  End If
               End If
            End If
         Else
            tbcf9saldo.AddNew
            tbcf9saldo.Fields("f9codcta") = tbconta.Fields("f5codcta")
            tbcf9saldo.Fields("f9nroref") = tbconta.Fields("f3nroref")
            tbcf9saldo.Fields("f9mesi") = txtmes.Text
            tbcf9saldo.Fields("f9saldo") = 0#
            tbcf9saldo.Fields("f9saldod") = 0#
            tbcf9saldo.Update
         End If
             
         tbcf9saldo.Seek "=", ccuenta, nref
         If Not tbcf9saldo.NoMatch Then
            tbcf9saldo.Edit
            If tbcf5pla.Fields("f5moneda") = "D" Then
               If tbcf9saldo.Fields("f9saldod") = 0 And "" & Len(Trim(tbcf9saldo.Fields("f9mesd") & "")) <> 0 Then
                  tbcf9saldo.Fields("f9mesd") = "  "
               End If
            Else
               If tbcf9saldo.Fields("f9saldo") = 0 And "" & Len(Trim(tbcf9saldo.Fields("f9mes") & "")) <> 0 Then
                  tbcf9saldo.Fields("f9mes") = "  "
               End If
            End If
               
            If tbconta.Fields("f3debhab") = "D" Then
               tbcf9saldo.Fields("f9saldo") = Val("" & tbcf9saldo.Fields("f9saldo")) + Val(tbconta.Fields("f3importe") & "")
               tbcf9saldo.Fields("f9saldod") = Val("" & tbcf9saldo.Fields("f9saldod")) + Val(tbconta.Fields("f3imported") & "")
            Else
               tbcf9saldo.Fields("f9saldo") = Val("" & tbcf9saldo.Fields("f9saldo")) - Val(tbconta.Fields("f3importe") & "")
               tbcf9saldo.Fields("f9saldod") = Val("" & tbcf9saldo.Fields("f9saldod")) - Val(tbconta.Fields("f3imported") & "")
            End If
            
            If tbcf5pla.Fields("f5moneda") = "D" Then
               If tbcf9saldo.Fields("f9saldod") = 0 And "" & Len(Trim(tbcf9saldo.Fields("f9mesd") & "")) = 0 Then
                  tbcf9saldo.Fields("f9mesd") = txtmes.Text
               End If
            Else
               If tbcf9saldo.Fields("f9saldo") = 0 And "" & Len(Trim(tbcf9saldo.Fields("f9mes") & "")) = 0 Then
                  tbcf9saldo.Fields("f9mes") = txtmes.Text
               End If
            End If
            tbcf9saldo.Update
         End If

         tbcf9cta.AddNew
         tbcf9cta.Fields("f9origen") = tbconta.Fields("f3origen")
         tbcf9cta.Fields("f9fecha") = tbconta.Fields("f3fchopr")
         tbcf9cta.Fields("f9nroref") = tbconta.Fields("f3nroref")
         tbcf9cta.Fields("f9compro") = tbconta.Fields("f3compro")
         tbcf9cta.Fields("f9debhab") = tbconta.Fields("f3debhab")
         tbcf9cta.Fields("f9detall") = tbconta.Fields("f3detall")
         tbcf9cta.Fields("f9elemen") = tbconta.Fields("f3elemen")
         tbcf9cta.Fields("f9cheque") = tbconta.Fields("f3cheque")
         tbcf9cta.Fields("f9codcta") = tbconta.Fields("f5codcta")
         tbcf9cta.Fields("f9importe") = tbconta.Fields("f3importe")
         tbcf9cta.Fields("f9imported") = tbconta.Fields("f3imported")
         tbcf9cta.Fields("f9tipdoc") = tbconta.Fields("f3tipdoc")
         tbcf9cta.Fields("f9tipcambd") = tbconta.Fields("f3tipcambd")
         tbcf9cta.Update
      End If
      tbconta.MoveNext
      If tbconta.EOF Then Exit Do
   Loop
   
End Sub

Private Sub REG_MAYOR()
Dim ccompro As String

   tbconta.Index = "contable"
   tbconta.MoveFirst
   Do While Not tbconta.EOF
      lblcuenta.Caption = "Procesando Comprobante : " & tbconta.Fields("f3compro")
      lblcuenta.Refresh
      ccompro = tbconta.Fields("f3compro")
      tbmovcf4.Seek "=", ccompro
      If Not tbmovcf4.NoMatch Then
         tbmovcf4.Edit
         tbmovcf4.Fields("f4compro") = tbconta.Fields("f3compro")
         tbmovcf4.Fields("f4origen") = tbconta.Fields("f3origen")
         tbmovcf4.Fields("f4obra") = tbconta.Fields("f3obra")
         tbmovcf4.Fields("f4moneda") = tbconta.Fields("f3moneda")
         tbmovcf4.Fields("f4fecha") = tbconta.Fields("f3fchopr")
         tbmovcf4.Fields("f4tipcambd") = tbconta.Fields("f3tipcambd")
      Else
         tbmovcf4.AddNew
         tbmovcf4.Fields("f4compro") = tbconta.Fields("f3compro")
         tbmovcf4.Fields("f4origen") = tbconta.Fields("f3origen")
         tbmovcf4.Fields("f4obra") = tbconta.Fields("f3obra")
         tbmovcf4.Fields("f4moneda") = tbconta.Fields("f3moneda")
         tbmovcf4.Fields("f4fecha") = tbconta.Fields("f3fchopr")
         tbmovcf4.Fields("f4tipcambd") = tbconta.Fields("f3tipcambd")
      End If
      If tbconta.Fields("f3debhab") = "D" Then
         tbmovcf4.Fields("f4totdeb") = Val(0 & tbmovcf4.Fields("f4totdeb")) + tbconta.Fields("f3importe")
         tbmovcf4.Fields("f4totdebd") = Val(0 & tbmovcf4.Fields("f4totdebd")) + tbconta.Fields("f3imported")
      Else
         tbmovcf4.Fields("f4tothab") = Val(0 & tbmovcf4.Fields("f4tothab")) + tbconta.Fields("f3importe")
         tbmovcf4.Fields("f4tothabd") = Val(0 & tbmovcf4.Fields("f4tothabd")) + tbconta.Fields("f3imported")
      End If
      tbmovcf4.Fields("f4numele") = Format(Val(tbmovcf4.Fields("f4numele") & "") + 1, "####")
      tbmovcf4.Update
      tbconta.MoveNext
      If tbconta.EOF Then Exit Do
   Loop
   
   tbconta.Index = "contable"
   tbconta.MoveFirst
   tbmovcf4.Index = "CF4TCO" & txtmes.Text
   tbmovcf4.Seek "=", tbconta.Fields("f3compro")
   If Not tbmovcf4.NoMatch Then
      Do While Not tbmovcf4.EOF
         lblcuenta.Caption = "Procesando Cuadre : " & tbmovcf4.Fields("f4compro")
         lblcuenta.Refresh
         tbmovcf4.Edit
         If Val(tbmovcf4.Fields("f4totdeb") & "") <> Val(tbmovcf4.Fields("f4tothab") & "") Then
            tbmovcf4.Fields("f4cuadres") = "1"
         Else
            tbmovcf4.Fields("f4cuadres") = "0"
         End If
         If Val(tbmovcf4.Fields("f4totdebd") & "") <> Val(tbmovcf4.Fields("f4tothabd") & "") Then
            tbmovcf4.Fields("f4cuadred") = "1"
         Else
            tbmovcf4.Fields("f4cuadred") = "0"
         End If
         tbmovcf4.Update
         tbmovcf4.MoveNext
         If tbmovcf4.EOF Then Exit Do
      Loop
   End If
      
End Sub

Private Sub REG_SALDOS()
Dim xmes As String
Dim gr      As Integer
Dim dg1     As Integer
Dim dg2     As Integer
Dim dg3     As Integer
Dim dg4     As Integer
Dim dg5     As Integer
Dim grado   As Integer
Dim ccta    As String

   xmes = txtmes.Text
   tbcntcont.Index = "cf1cnt"
   tbcntcont.Seek "=", wcontacnt
   If Not tbcntcont.NoMatch Then
      gr = 1
      Do While tbcontrol.Fields("f1grdmax") + 1 > gr
         Select Case gr
            Case 1: dg1 = tbcontrol.Fields("f1dgrad1")
            Case 2: dg2 = tbcontrol.Fields("f1dgrad2")
            Case 3: dg3 = tbcontrol.Fields("f1dgrad3")
            Case 4: dg4 = tbcontrol.Fields("f1dgrad4")
            Case 5: dg5 = tbcontrol.Fields("f1dgrad5")
         End Select
         gr = gr + 1
      Loop
   End If
                    
   tbconta.MoveFirst
   Do While Not tbconta.EOF
      lblcuenta.Caption = "Procesando Cuenta : " & tbconta.Fields("f5codcta")
      lblcuenta.Refresh
      ccuenta = tbconta.Fields("f5codcta")
      tbcf5pla.Seek "=", ccuenta
      If Not tbcf5pla.NoMatch Then
         grado = tbcf5pla.Fields("f5grdcta") + 1
         gr = 1
         Do While grado > gr
            Select Case gr
               Case 1: ccta = Mid(ccuenta, 1, dg1)
               Case 2: ccta = Mid(ccuenta, 1, dg2)
               Case 3: ccta = Mid(ccuenta, 1, dg3)
               Case 4: ccta = Mid(ccuenta, 1, dg4)
               Case 5: ccta = Mid(ccuenta, 1, dg5)
            End Select
            ccta = Trim(ccta) & Space(12 - Len(Trim(ccta)))
            tbcf5pla.Seek "=", ccta
            If Not tbcf5pla.NoMatch Then
               tbcf5pla.Edit
               If tbconta.Fields("f3debhab") = "D" Then
                  tbcf5pla.Fields("f5debm" & xmes) = Val("" & tbcf5pla.Fields("f5debm" & xmes)) + Val(tbconta.Fields("f3importe") & "")
                  tbcf5pla.Fields("f5debm99") = Val("" & tbcf5pla.Fields("f5debm99")) + Val(tbconta.Fields("f3importe") & "")
                  tbcf5pla.Fields("f5debdm" & xmes) = Val("" & tbcf5pla.Fields("f5debdm" & xmes)) + Val(tbconta.Fields("f3imported") & "")
                  tbcf5pla.Fields("f5debdm99") = Val("" & tbcf5pla.Fields("f5debdm99")) + Val(tbconta.Fields("f3imported") & "")
               Else  ' haber - "H"
                  tbcf5pla.Fields("f5habm" & xmes) = Val("" & tbcf5pla.Fields("f5habm" & xmes)) + Val(tbconta.Fields("f3importe") & "")
                  tbcf5pla.Fields("f5habm99") = Val("" & tbcf5pla.Fields("f5habm99")) + Val(tbconta.Fields("f3importe") & "")
                  tbcf5pla.Fields("f5habdm" & xmes) = Val("" & tbcf5pla.Fields("f5habdm" & xmes)) + Val(tbconta.Fields("f3imported") & "")
                  tbcf5pla.Fields("f5habdm99") = Val("" & tbcf5pla.Fields("f5habdm99")) + Val(tbconta.Fields("f3imported") & "")
               End If
               tbcf5pla.Update
            End If
            gr = gr + 1
         Loop
      End If
      tbconta.MoveNext
      If tbconta.EOF Then Exit Do
   Loop

End Sub

Private Sub transfiere()

    Me.MousePointer = 11
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    
    Set dbtemconta = OpenDatabase(wrutatemp & "\db_conta.mdb")
    Set tbconta = dbtemconta.OpenRecordset("contable")
    
    Set dbcntcont = OpenDatabase(wcontacnt & "\cnt_cont.mdb")
    Set tbcntcont = dbcntcont.OpenRecordset("cf1cnt")
    
    Set dbmovconta = OpenDatabase(wrutaconta & "\db_mov" & txtmes.Text & ".mdb")
    Set tbmovcf3 = dbmovconta.OpenRecordset("cf3mov" & txtmes.Text)
    Set tbmovcf4 = dbmovconta.OpenRecordset("cf4tco" & txtmes.Text)
    tbmovcf4.Index = "cf4tco" & txtmes.Text
    
    Set dbtabla = OpenDatabase(wrutaconta & "\db_tabla.mdb")
    Set tbcf5pla = dbtabla.OpenRecordset("cf5pla")
    tbcf5pla.Index = "cf5pla"
    
    Set dbanalisis = OpenDatabase(wrutaconta & "\db_anali.mdb")
    Set tbcf9saldo = dbanalisis.OpenRecordset("cf9saldo")
    tbcf9saldo.Index = "cf9saldo"
    Set tbcf9cta = dbanalisis.OpenRecordset("cf9cta")
    tbcf9cta.Index = "cf9cta"

    If tbconta.RecordCount > 0 Then
        tbconta.MoveFirst
        If Mid(tbconta.Fields("f3proame"), 5, 2) = txtmes.Text Then
            Do While Not tbconta.EOF
                tbmovcf3.AddNew
                tbmovcf3.Fields("f3compro") = tbconta.Fields("f3compro") & ""
                tbmovcf3.Fields("f3proame") = tbconta.Fields("f3proame") & ""
                tbmovcf3.Fields("f3elemen") = tbconta.Fields("f3elemen") & ""
                tbmovcf3.Fields("f3origen") = tbconta.Fields("f3origen") & ""
                tbmovcf3.Fields("f3fchopr") = tbconta.Fields("f3fchopr")
                tbmovcf3.Fields("f3detall") = tbconta.Fields("f3detall") & ""
                tbmovcf3.Fields("f5codcta") = tbconta.Fields("f5codcta") & ""
                tbmovcf3.Fields("f3codgas") = tbconta.Fields("f3codgas") & ""
                tbmovcf3.Fields("f3cheque") = tbconta.Fields("f3cheque")
                tbmovcf3.Fields("f3nroref") = tbconta.Fields("f3nroref") & ""
                tbmovcf3.Fields("f3importe") = tbconta.Fields("f3importe")
                tbmovcf3.Fields("f3imported") = tbconta.Fields("f3imported")
                tbmovcf3.Fields("f3moneda") = tbconta.Fields("f3moneda") & ""
                tbmovcf3.Fields("f3tipcambd") = tbconta.Fields("f3tipcambd")
                tbmovcf3.Fields("f3tipdoc") = tbconta.Fields("f3tipdoc") & ""
                tbmovcf3.Fields("f3debhab") = tbconta.Fields("f3debhab") & ""
                tbmovcf3.Fields("f3costo") = tbconta.Fields("f3costo") & ""
                tbmovcf3.Fields("f3destino") = IIf(tbconta.Fields("f3destino") & "" = "*", "*", " ")
                tbmovcf3.Fields("F3CTAAUX") = tbconta.Fields("F3RUC") & ""
                tbmovcf3.Update
                TbCabRegis.Seek "=", left(tbconta.Fields("f3nummov"), 2), Mid(tbconta.Fields("f3nummov"), 3, 7)
                If Not TbCabRegis.NoMatch Then
                    TbCabRegis.Edit
                    TbCabRegis.Fields("f4contable") = "*"
                    TbCabRegis.Update
                End If
                tbconta.MoveNext
                If tbconta.EOF Then Exit Do
            Loop
            REG_MAYOR
            REG_SALDOS

            dbtemconta.Execute ("delete * from contable")

        Else
            MsgBox "Los asientos generados no corresponden al mes que se desea transferir. Verifique", 48, "Compras"
        End If
    Else
        MsgBox "No se han generado los asientos contables. Verifique", 48, "Compras"
    End If

    TbCabRegis.Close
    dbcompras.Close
    tbconta.Close
    dbtemconta.Close
    tbcntcont.Close
    dbcntcont.Close
    tbmovcf3.Close
    tbmovcf4.Close
    dbmovconta.Close
    tbcf5pla.Close
    dbtabla.Close
    tbcf9cta.Close
    tbcf9saldo.Close
    dbanalisis.Close
    
    Me.MousePointer = 1

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      If Val(txtmes.Text) >= 1 And Val(txtmes.Text) <= 12 Then
         txtmes.Text = Format(txtmes.Text, "00")
         cmdresp(0).SetFocus
      End If
   End If

End Sub

