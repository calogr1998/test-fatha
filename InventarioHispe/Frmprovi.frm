VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmprovi 
   Caption         =   "Resumen de Provisiones"
   ClientHeight    =   3780
   ClientLeft      =   1995
   ClientTop       =   1395
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   7635
   Begin Threed.SSPanel PanelCab 
      Height          =   465
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   820
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
      Begin Threed.SSCommand BtnExit 
         Height          =   330
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   90
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483640
         Picture         =   "Frmprovi.frx":0000
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483640
         Picture         =   "Frmprovi.frx":015A
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6660
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   3165
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   7530
      _Version        =   65536
      _ExtentX        =   13282
      _ExtentY        =   5583
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
      Begin VB.Data dataprovi 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5580
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2700
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSDBGrid.DBGrid grdprovi 
         Bindings        =   "Frmprovi.frx":069C
         Height          =   2985
         Left            =   45
         OleObjectBlob   =   "Frmprovi.frx":06B4
         TabIndex        =   1
         Top             =   90
         Width           =   7395
      End
   End
End
Attribute VB_Name = "frmprovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub llena_tempo()
Dim CONSULTA        As DAO.Recordset
Dim tabprovi        As DAO.Recordset
Dim tbcabecera      As DAO.Recordset
Dim cmoneda         As String
Dim ntc             As Double
Dim cabrev          As String
Dim dbcomtabla      As DAO.Database
Dim tbdocumentos    As DAO.Recordset
Dim dbtabla         As DAO.Database
Dim tbcuenta        As DAO.Recordset
Dim codcuent        As String

    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("Delete From temp_PROVI")
    Set tabprovi = dbtemp.OpenRecordset("TEMP_provi")
    
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbcabecera = dbcompras.OpenRecordset("regisdoc")
    
    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set tbdocumentos = dbcomtabla.OpenRecordset("documentos")

    Set dbtabla = OpenDatabase(wrutaconta & "\db_tabla")
    Set tbcuenta = dbtabla.OpenRecordset("CF5PLA")
       
    Set CONSULTA = dbcompras.OpenRecordset("Select * from regismov where f4mesmov='" & mes & "' ")
    If CONSULTA.RecordCount > 0 Then
        CONSULTA.MoveFirst
        Do While Not CONSULTA.EOF
            tbcabecera.Index = "IDMESNUM"
            tbcabecera.Seek "=", "" & CONSULTA.Fields("f4mesmov"), "" & CONSULTA.Fields("f4nummov")
            If Not tbcabecera.NoMatch Then
                cmoneda = ""
                ntc = 0#
                cmoneda = tbcabecera.Fields("f4moneda") & ""
                ntc = tbcabecera.Fields("f4tipcam")
                tabprovi.Index = "temp_cta"
                tabprovi.Seek "=", CONSULTA.Fields("f3ctacon")
                If Not tabprovi.NoMatch Then
                    tabprovi.Edit
                Else
                    tabprovi.AddNew
                End If
                codcuent = CONSULTA.Fields("f3ctacon")
                tabprovi.Fields("F3GASTO") = CONSULTA.Fields("F3GASTO")
                tabprovi.Fields("F3CTACON") = CONSULTA.Fields("F3CTACON")
                
                tbcuenta.Index = "CF5PLA"
                tbcuenta.Seek "=", codcuent
                If Not tbcuenta.NoMatch Then
                    tabprovi.Fields("F3NOMCTA") = tbcuenta.Fields("F5NOMCTA")
                Else
                    tabprovi.Fields("F3NOMCTA") = ""
                End If
                    
                If cmoneda = "S" Then
                    If CONSULTA.Fields("F3DEBHAB") = "D" Then
                        tabprovi.Fields("F3DEBE_S") = CONSULTA.Fields("F3IMPORTE") + tabprovi.Fields("F3DEBE_S")
                        If ntc > 0 Then
                            tabprovi.Fields("F3DEBE_D") = Val(Format(CONSULTA.Fields("F3IMPORTE") / ntc, "0.00")) + tabprovi.Fields("F3DEBE_D")
                        End If
                    Else
                        tabprovi.Fields("F3HABER_S") = tabprovi.Fields("F3HABER_S") + CONSULTA.Fields("F3IMPORTE")
                        If ntc > 0 Then
                            tabprovi.Fields("F3HABER_D") = tabprovi.Fields("F3HABER_D") + Val(Format(CONSULTA.Fields("F3IMPORTE") / ntc, "0.00"))
                        End If
                    End If
                Else
                    If CONSULTA.Fields("F3DEBHAB") = "D" Then
                        tabprovi.Fields("F3DEBE_D") = CONSULTA.Fields("F3IMPORTE") + tabprovi.Fields("F3DEBE_D")
                        If ntc > 0 Then
                           tabprovi.Fields("F3DEBE_S") = Val(Format(CONSULTA.Fields("F3IMPORTE") * ntc, "0.00")) + tabprovi.Fields("F3DEBE_S")
                        End If
                    Else
                        tabprovi.Fields("F3HABER_D") = tabprovi.Fields("F3HABER_D") + CONSULTA.Fields("F3IMPORTE")
                        If ntc > 0 Then
                            tabprovi.Fields("F3HABER_S") = tabprovi.Fields("F3HABER_S") + Val(Format(CONSULTA.Fields("F3IMPORTE") * ntc, "0.00"))
                        End If
                    End If
                End If
                tabprovi.Update
            End If
            CONSULTA.MoveNext
        Loop
        '-----------------------------------------------------------------------------
        '-----------------------------------------------------------------------------
        Set CONSULTA = dbcompras.OpenRecordset("Select * from regisdoc where f4mesmov='" & mes & "'")
        If CONSULTA.RecordCount > 0 Then
            CONSULTA.MoveFirst
            Do While Not CONSULTA.EOF
                ntc = CONSULTA.Fields("f4tipcam")
                cabrev = ""
                tbdocumentos.Index = "idcoddoc"
                tbdocumentos.Seek "=", CONSULTA.Fields("f4tipdoc")
                If Not tbdocumentos.NoMatch Then
                    cabrev = tbdocumentos.Fields("f2abrev")
                End If
                tabprovi.Index = "temp_cta"
                tabprovi.Seek "=", wctaigv
                If Not tabprovi.NoMatch Then
                    tabprovi.Edit
                Else
                    tabprovi.AddNew
                End If
                tabprovi.Fields("F3CTACON") = wctaigv
                If CONSULTA.Fields("f4moneda") = "S" Then
                    If CONSULTA.Fields("f4igv") > 0 Then
                        tabprovi.Fields("f3debe_s") = tabprovi.Fields("f3debe_s") + CONSULTA.Fields("f4igv")
                        If ntc > 0 Then
                            tabprovi.Fields("f3debe_d") = tabprovi.Fields("f3debe_d") + Format(CONSULTA.Fields("f4igv") / ntc, "#0.00")
                        End If
                    End If
                Else
                    If CONSULTA.Fields("f4igv") > 0 Then
                        tabprovi.Fields("f3debe_d") = tabprovi.Fields("f3debe_d") + CONSULTA.Fields("f4igv")
                        If ntc > 0 Then
                            tabprovi.Fields("f3debe_s") = tabprovi.Fields("f3debe_s") + Format(CONSULTA.Fields("f4igv") * ntc, "#0.00")
                        End If
                    End If
                End If
                tabprovi.Update
                '----------------------------------------------------------------------------
                tabprovi.Index = "temp_cta"
                If Len(Trim(CONSULTA.Fields("f4ctacont") & "")) > 4 Then
                    tabprovi.Seek "=", Left(Trim(CONSULTA.Fields("f4ctacont") & ""), Len(Trim(CONSULTA.Fields("f4ctacont") & "")) - 4)
                Else
                    tabprovi.Seek "=", Trim(CONSULTA.Fields("f4ctacont") & "")
                End If
                
                If Not tabprovi.NoMatch Then
                    tabprovi.Edit
                Else
                    tabprovi.AddNew
                End If
                If Len(Trim(CONSULTA.Fields("f4ctacont") & "")) > 4 Then
                    tabprovi.Fields("F3CTACON") = Left(Trim(CONSULTA.Fields("f4ctacont")), Len(Trim(CONSULTA.Fields("f4ctacont"))) - 4)
                Else
                    tabprovi.Fields("F3CTACON") = Trim(CONSULTA.Fields("f4ctacont") & "")
                End If
                If CONSULTA.Fields("f4moneda") = "S" Then
                    If CONSULTA.Fields("f4total") > 0 Then
                        tabprovi.Fields("f3debe_s") = tabprovi.Fields("f3debe_s") + CONSULTA.Fields("f4total")
                        If ntc > 0 Then
                            tabprovi.Fields("f3debe_d") = tabprovi.Fields("f3debe_d") + Format(CONSULTA.Fields("f4total") / ntc, "#0.00")
                        End If
                    End If
                Else
                    If CONSULTA.Fields("f4total") > 0 Then
                        tabprovi.Fields("f3debe_d") = tabprovi.Fields("f3debe_d") + CONSULTA.Fields("f4total")
                        If ntc > 0 Then
                            tabprovi.Fields("f3debe_s") = tabprovi.Fields("f3debe_s") + Format(CONSULTA.Fields("f4total") * ntc, "#0.00")
                        End If
                    End If
                End If
                tabprovi.Update
                '----------------------------------------------------------------------------
                CONSULTA.MoveNext
            Loop
        End If

    End If

    CONSULTA.Close
    tbcabecera.Close
    dbcompras.Close
    
    tbcuenta.Close
    dbtabla.Close
    
    tbdocumentos.Close
    dbcomtabla.Close
    
    tabprovi.Close
    dbtemp.Close

End Sub

Private Sub BtnExit_Click()
 Unload Me
End Sub


Private Sub BtnPrint_Click()
    CrystalReport1.DataFiles(0) = wrutatemp & "\temp_com.Mdb"
    CrystalReport1.ReportFileName = wrutatemp & "\res_provi1.RPT"
    CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
    
    llena_tempo
    dataprovi.DatabaseName = wrutatemp & "\temp_com.Mdb"
    dataprovi.RecordSource = "temp_provi"
    dataprovi.Refresh
    
End Sub
