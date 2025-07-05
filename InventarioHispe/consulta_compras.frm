VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form consulta_compras 
   Caption         =   "Consulta - Registro de Compras"
   ClientHeight    =   2508
   ClientLeft      =   4932
   ClientTop       =   2460
   ClientWidth     =   5268
   LinkTopic       =   "Form1"
   ScaleHeight     =   2508
   ScaleWidth      =   5268
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5100
      Begin VB.Frame Frame2 
         Caption         =   " Ordenar "
         Height          =   825
         Left            =   135
         TabIndex        =   5
         Top             =   765
         Width           =   4875
         Begin VB.OptionButton optorden 
            Caption         =   "Fecha"
            Height          =   420
            Index           =   1
            Left            =   3555
            TabIndex        =   7
            Top             =   315
            Width           =   1005
         End
         Begin VB.OptionButton optorden 
            Caption         =   "Nro. Registro"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   6
            Top             =   405
            Value           =   -1  'True
            Width           =   2130
         End
      End
      Begin VB.TextBox txtmes 
         Height          =   330
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   1
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   210
         Left            =   1935
         TabIndex        =   2
         Top             =   405
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   465
      Left            =   2655
      TabIndex        =   3
      Top             =   1935
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   468
      Left            =   1332
      TabIndex        =   4
      Top             =   1944
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
End
Attribute VB_Name = "consulta_compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTITULO As String
Private Sub cmdaceptar_Click()

    wmesregcompras = txtmes.Text
    wordenfecha = IIf(optorden(0).Value = True, "P", "F")
    GENERA_TEMP
    With acr_regcompras
        .DataControl1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\temp_com.MDB;Persist Security Info=False"
        If wordenfecha = "P" Then
            .DataControl1.Source = "Select * from temp_gg Order By TEMP_MOV"
        Else
            .DataControl1.Source = "Select * from temp_gg Order By TEMP_FECHA"
        End If
        .fldtitulo.Text = CTITULO & " - Moneda: Soles"
        .fldempresa.Text = wf4empresa
        .fldfecha.Text = Format(Date, "dd/mm/yyyy")
        .Show vbModal
    End With
    
End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    txtmes.Text = mes

End Sub

Private Sub txtmes_GotFocus()

    txtmes.SelStart = 0: txtmes.SelLength = Len(txtmes.Text)

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdaceptar.SetFocus
    End If

End Sub
Private Sub GENERA_TEMP()
Dim dbtempo As Database
Dim tbtempo As dao.Recordset
Dim sw_cambio_igv   As Boolean
                    
    csql = "SELECT REGISDOC.F4PORC_IGV, REGISDOC.F4FECHADEPOSITO, REGISDOC.F4NUMDEPOSITO, REGISDOC.F4NUMMOV, REGISDOC.F4NOMPRV, REGISDOC.F4TIPDOC, REGISDOC.F4SERDOC, REGISDOC.F4NUMDOC, REGISDOC.F4FECHA, REGISDOC.F4MONEDA, VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4BASIMP],[REGISDOC].[F4BASIMP]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4BASIMP,  REGISDOC.F4REFERE, REGISDOC.F4CODPRV, REGISDOC.F4USUARIOING, REGISDOC.F4DCTO, REGISDOC.F4RUCPRV,  REGISDOC.F4CODIGV,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4MONINA],[REGISDOC].[F4MONINA]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4MONINA,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4IGV],[REGISDOC].[F4IGV]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4IGV,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4TOTAL],[REGISDOC].[F4TOTAL]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4TOTAL,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[f4redsuma],[REGISDOC].[f4redsuma]*[REGISDOC].[F4TIPCAM]),'0.00')) AS f4redsuma,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S' " & _
            ",[REGISDOC].[f4redresta],[REGISDOC].[f4redresta]*[REGISDOC].[F4TIPCAM]),'0.00')) AS f4redresta,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4DCTO],[REGISDOC].[F4DCTO]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4DCTO,VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',[REGISDOC].[F4OTRIMP],[REGISDOC].[F4OTRIMP]*[REGISDOC].[F4TIPCAM]),'0.00')) AS F4OTRIMP,REGISDOC.F4TIPCAM, VAL(FORMAT(IIf([REGISDOC].[F4MONEDA]='S',0,[REGISDOC].[F4TOTAL]),'0.00')) AS TOTDOL " & _
            "FROM REGISDOC WHERE REGISDOC.F4TIPDOC <> '02' AND REGISDOC.F4MESMOV = '" & txtmes.Text & "'" & _
            "ORDER BY F4NUMMOV DESC"
    If rsif4orden.State = 1 Then rsif4orden.Close
    rsif4orden.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    rsif4orden.MoveFirst
    If Val(rsif4orden.Fields(3)) > 0 Then
        CTITULO = "Mes de " & txtmes.Text & " de " & wanno
        Set dbtempo = OpenDatabase(wrutatemp & "\temp_com.mdb")
        dbtempo.Execute ("delete * from temp_gg")
        Set tbtempo = dbtempo.OpenRecordset("temp_gg")

        rsif4orden.MoveFirst
        Do While Not rsif4orden.EOF
            tbtempo.AddNew
            tbtempo.Fields("temp_mov") = rsif4orden.Fields("f4nummov")
            tbtempo.Fields("temp_fecha") = rsif4orden.Fields("f4fecha")
            tbtempo.Fields("TEMP_TDOC") = Format(rsif4orden.Fields("F4TIPDOC"), "00")
            tbtempo.Fields("temp_serie") = Format(rsif4orden.Fields("f4serdoc"), "000")
            tbtempo.Fields("temp_docum") = Format(rsif4orden.Fields("f4numdoc"), "0000000")
            tbtempo.Fields("temp_prov") = rsif4orden.Fields("f4nomprv") & ""
            tbtempo.Fields("temp_detal") = rsif4orden.Fields("f4refere") & ""
            If Format(rsif4orden.Fields("F4TIPDOC"), "00") = "02" Then
                tbtempo.Fields("TEMP_BIMP") = 0#
                tbtempo.Fields("TEMP_EXON") = Val("" & rsif4orden.Fields("F4BASIMP")) + Val("" & rsif4orden.Fields("F4MONINA"))
                tbtempo.Fields("temp_totals") = Val("" & rsif4orden.Fields("F4BASIMP")) + Val("" & rsif4orden.Fields("F4MONINA"))
            Else
                sw_cambio_igv = False
                'If wmesregcompras >= wf1mescambio_igv Then
                '    If Month(rsif4orden.fields("f4fecha")) < Val(wf1mescambio_igv) Then
                '        sw_cambio_igv = True
                '    Else
                '        sw_cambio_igv = False
                '    End If
                'Else
                '    sw_cambio_igv = False
                'End If
                
                If wmesregcompras >= wf1mescambio_igv Then
                    If Val(rsif4orden.Fields("F4PORC_IGV") & "") <> gigv Then
                        sw_cambio_igv = True
                    Else
                        sw_cambio_igv = False
                    End If
                Else
                    sw_cambio_igv = False
                End If
                
                'If sw_cambio_igv = False Then
                    'Select Case rsif4orden.Fields("F4CODIGV") & ""
                     '   Case "001":
                            tbtempo.Fields("TEMP_BIMP") = Val("" & rsif4orden.Fields("F4BASIMP"))
                            tbtempo.Fields("temp_igvs") = Val("" & rsif4orden.Fields("f4igv"))
                     '   Case "002":
                     '       tbtempo.Fields("TEMP_BIMP_GYNG") = Val("" & rsif4orden.Fields("F4BASIMP"))
                     '       tbtempo.Fields("TEMP_IGVS_GYNG") = Val("" & rsif4orden.Fields("f4igv"))
                     '   Case "003":
                     '       tbtempo.Fields("TEMP_BIMP_SIN") = Val("" & rsif4orden.Fields("F4BASIMP"))
                     '       tbtempo.Fields("TEMP_IGVS_SIN") = Val("" & rsif4orden.Fields("f4igv"))
                    'End Select
                'Else
                '    tbtempo.Fields("TEMP_BIMP_OTRO") = Val("" & rsif4orden.Fields("F4BASIMP"))
                '    tbtempo.Fields("TEMP_IGVS_OTRO") = Val("" & rsif4orden.Fields("f4igv"))
                'End If
                
                tbtempo.Fields("TEMP_EXON") = Val("" & rsif4orden.Fields("F4OTRIMP")) + Val("" & rsif4orden.Fields("F4MONINA")) + Val("" & rsif4orden.Fields("F4REDSUMA")) - Val("" & rsif4orden.Fields("F4REDRESTA")) - Val("" & rsif4orden.Fields("F4DCTO"))
                tbtempo.Fields("temp_totals") = Val("" & rsif4orden.Fields("f4total"))
                tbtempo.Fields("temp_totalD") = Val("" & rsif4orden.Fields("TOTDOL"))
                tbtempo.Fields("temp_IGVD") = Val("" & rsif4orden.Fields("F4TIPCAM"))
            End If
            tbtempo.Fields("empresa") = wf4empresa
            tbtempo.Fields("temp_moneda") = IIf(rsif4orden.Fields("f4moneda") & "" = "S", "S/.", "US$")
            tbtempo.Fields("TEMP_RUC") = rsif4orden.Fields("F4RUCPRV") & ""
            tbtempo.Fields("MES") = CTITULO
            tbtempo.Fields("TEMP_CODIGV") = rsif4orden.Fields("F4CODIGV") & ""
            tbtempo.Fields("TEMP_NUMDEPOSITO") = Trim(rsif4orden.Fields("F4NUMDEPOSITO") & "")
            tbtempo.Fields("TEMP_FECHADEPOSITO") = rsif4orden.Fields("F4FECHADEPOSITO")
            tbtempo.Update
            rsif4orden.MoveNext
        Loop
        tbtempo.Close
        dbtempo.Close
    Else
        MsgBox "No hay registros.", 48, "Atención"
    End If

End Sub
