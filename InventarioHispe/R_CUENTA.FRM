VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form R_CUENTA 
   Caption         =   "Consulta por Cuenta Contable"
   ClientHeight    =   5100
   ClientLeft      =   1665
   ClientTop       =   1365
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9735
   Begin Threed.SSPanel SSPanel1 
      Height          =   870
      Left            =   135
      TabIndex        =   5
      Top             =   675
      Width           =   9420
      _Version        =   65536
      _ExtentX        =   16616
      _ExtentY        =   1535
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
      Begin Threed.SSPanel Pnldescripcion 
         Height          =   285
         Left            =   3015
         TabIndex        =   6
         Top             =   315
         Width           =   6270
         _Version        =   65536
         _ExtentX        =   11060
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
      Begin Threed.SSPanel pnlcuenta 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   315
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Data dataconta 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2970
      Visible         =   0   'False
      Width           =   1185
   End
   Begin Threed.SSPanel PanelCab 
      Height          =   465
      Left            =   135
      TabIndex        =   2
      Top             =   45
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
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
      Begin Threed.SSCommand BtnExitO 
         Height          =   330
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "R_CUENTA.frx":0000
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   45
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "R_CUENTA.frx":015A
      End
   End
   Begin MSDBGrid.DBGrid grid1 
      Bindings        =   "R_CUENTA.frx":069C
      Height          =   3435
      Left            =   135
      OleObjectBlob   =   "R_CUENTA.frx":06B4
      TabIndex        =   1
      Top             =   1575
      Width           =   9465
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   8955
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Threed.SSPanel PnlMes 
      Height          =   330
      Left            =   1395
      TabIndex        =   0
      Top             =   135
      Width           =   7305
      _Version        =   65536
      _ExtentX        =   12885
      _ExtentY        =   582
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
End
Attribute VB_Name = "R_CUENTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbtemp_regismov As DAO.Recordset
Dim Fnummes         As String

Sub Proceso()
Dim TbConsulta      As DAO.Recordset
Dim totimporte      As Double
Dim tempdebe        As Double, temphaber As Double
Dim totdebe         As Double, tothaber As Double
Dim xtitulo         As String
Dim dbinvtemp       As DAO.Database
Dim Msql            As String, xdetall As String, xf4mes As String
    
    Set dbinvtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set tbtemp_regismov = dbinvtemp.OpenRecordset("temp_regismov")
    
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "IDMESNUM"

    xtitulo = PnlMes.Caption
    If c_cuenta.opcmodo(0) = True Then
        Msql = "Select * From REGISMOV WHERE F4MESMOV = '" + c_cuenta.txtmes.Text + "' order by F4NUMMOV"
    Else
        Msql = "Select * From  REGISMOV WHERE VAL(F4MESMOV) <=" & Val(c_cuenta.txtmes.Text) & " order by F4NUMMOV"
    End If
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
    totimporte = 0
    Do While Not TbConsulta.EOF
        If Trim(TbConsulta.Fields("f3ctacon")) >= Trim(c_cuenta.txtcuenta.Text) And Trim(TbConsulta.Fields("f3ctacon")) <= Trim(c_cuenta.txtctahasta.Text) Then
            xdetall = ""
            tbtemp_regismov.AddNew
            TbCabRegis.Seek "=", TbConsulta.Fields("f4MESMOV"), TbConsulta.Fields("f4NUMMOV")
            If Not TbCabRegis.NoMatch = True Then
                tbtemp_regismov.Fields("F4MONEDA") = TbCabRegis.Fields("F4MONEDA")
                tbtemp_regismov.Fields("F4FECHA") = TbCabRegis.Fields("F4FECHA")
                tbtemp_regismov.Fields("F4TIPCAM") = TbCabRegis.Fields("F4TIPCAM")
                xdetall = TbCabRegis.Fields("F4REFERE") & ""
            End If
            Select Case gmonedacta
            Case "S"
                If TbCabRegis.Fields("F4MONEDA") = "S" Then
                    If TbConsulta.Fields("F3debhab") = "D" Then
                       tbtemp_regismov.Fields("F3DEBE") = TbConsulta.Fields("F3IMPORTE")
                       tempdebe = TbConsulta.Fields("F3IMPORTE")
                       totdebe = totdebe + tempdebe
                    Else
                        tbtemp_regismov.Fields("F3HABER") = TbConsulta.Fields("F3IMPORTE")
                        temphaber = TbConsulta.Fields("F3IMPORTE")
                        tothaber = tothaber + temphaber
                    End If
                End If

                If TbCabRegis.Fields("F4MONEDA") = "D" Then
                    If TbConsulta.Fields("F3debhab") = "D" Then
                        tbtemp_regismov.Fields("F3DEBE") = Format(TbConsulta.Fields("F3IMPORTE") * TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        tempdebe = Format(TbConsulta.Fields("F3IMPORTE") * TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        totdebe = totdebe + tempdebe
                    Else
                        tbtemp_regismov.Fields("F3HABER") = Format(TbConsulta.Fields("F3IMPORTE") * TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        temphaber = Format(TbConsulta.Fields("F3IMPORTE") * TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        tothaber = tothaber + temphaber
                    End If
                End If

            Case "D"
                If TbCabRegis.Fields("F4MONEDA") = "D" Then
                    If TbConsulta.Fields("F3debhab") = "D" Then
                        tbtemp_regismov.Fields("F3DEBE") = TbConsulta.Fields("F3IMPORTE")
                        tempdebe = TbConsulta.Fields("F3IMPORTE")
                        totdebe = totdebe + tempdebe
                    Else
                        tbtemp_regismov.Fields("F3HABER") = TbConsulta.Fields("F3IMPORTE")
                        temphaber = TbConsulta.Fields("F3IMPORTE")
                        tothaber = tothaber + temphaber
                    End If
                End If

                If TbCabRegis.Fields("F4MONEDA") = "S" Then
                    If TbConsulta.Fields("F3debhab") = "D" Then
                        tbtemp_regismov.Fields("F3DEBE") = Format(TbConsulta.Fields("F3IMPORTE") / TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        tempdebe = Format(TbConsulta.Fields("F3IMPORTE") / TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        totdebe = totdebe + tempdebe
                    Else
                        tbtemp_regismov.Fields("F3HABER") = Format(TbConsulta.Fields("F3IMPORTE") / TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        temphaber = Format(TbConsulta.Fields("F3IMPORTE") / TbCabRegis.Fields("F4TIPCAM"), "#0.00")
                        tothaber = tothaber + temphaber
                    End If
                End If
                
            End Select

            tbtemp_regismov.Fields("F4MESMOV") = TbConsulta.Fields("F4MESMOV")
            tbtemp_regismov.Fields("F4NUMMOV") = TbConsulta.Fields("F4NUMMOV")
            tbtemp_regismov.Fields("F3GASTO") = TbConsulta.Fields("F3GASTO")
            tbtemp_regismov.Fields("F3CONCEPTO") = Trim(TbConsulta.Fields("F3CONCEPTO") & "")
            tbtemp_regismov.Fields("F3ANOORDEN") = TbConsulta.Fields("F3ANOORDEN")
            tbtemp_regismov.Fields("F3ORDEN") = TbConsulta.Fields("F3ORDEN")
            tbtemp_regismov.Fields("F3CENCOS") = TbConsulta.Fields("F3CENCOS")
            tbtemp_regismov.Fields("F3CTACON") = TbConsulta.Fields("F3CTACON")
            tbtemp_regismov.Fields("F3DEBHAB") = TbConsulta.Fields("F3DEBHAB")
            tbtemp_regismov.Fields("F3IMPORTE") = TbConsulta.Fields("F3IMPORTE")
            tbtemp_regismov.Fields("F4CTADES") = TbConsulta.Fields("F3CTACON") & ""
            tbtemp_regismov.Fields("F4MONDES") = gmondes
            tbtemp_regismov.Fields("F4TITULO") = xtitulo

            Select Case TbConsulta.Fields("F4MESMOV")
                Case 1: xf4mes = "ENERO"
                Case 2: xf4mes = "FEBRERO"
                Case 3: xf4mes = "MARZO"
                Case 4: xf4mes = "ABRIL"
                Case 5: xf4mes = "MAYO"
                Case 6: xf4mes = "JUNIO"
                Case 7: xf4mes = "JULIO"
                Case 8: xf4mes = "AGOSTO"
                Case 9: xf4mes = "SETIEMBRE"
                Case 10: xf4mes = "OCTUBRE"
                Case 11: xf4mes = "NOVIEMBRE"
                Case 12: xf4mes = "DICIEMBRE"
            End Select
            tbtemp_regismov.Fields("MES") = xf4mes

            tbplancta.Index = "cf5pla"
            tbplancta.Seek "=", TbConsulta.Fields("F3CTACON") & ""
            If Not tbplancta.NoMatch Then
                tbtemp_regismov.Fields("CUENTA") = Trim(tbplancta.Fields("f5nomcta") & "")
            End If
            
            tbtemp_regismov.Update
            totimporte = totimporte + TbConsulta.Fields("F3IMPORTE")

        End If
        TbConsulta.MoveNext
        If TbConsulta.EOF = True Then Exit Do
    Loop

End Sub

Private Sub BtnExitO_Click()
    Unload Me
End Sub

Private Sub BtnPrint_Click()
    
    Report1.DataFiles(0) = wrutatemp & "\temp_com.mdb"
    Report1.ReportFileName = wrutatemp & "\rcuenta.rpt"
    Report1.Action = 1
    
End Sub


Private Sub Form_Load()
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    Set TbOfiRegis = dbcompras.OpenRecordset("REGISMOV")

    Set dbtempcomp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    dbtempcomp.Execute ("Delete From temp_regismov")
    Set tbtemp_regismov = dbtempcomp.OpenRecordset("temp_regismov")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")

    Set dbplancta = OpenDatabase(wrutaconta & "\db_tabla.mdb")
    Set tbplancta = dbplancta.OpenRecordset("cf5pla")
    tbplancta.Index = "cf5pla"

    tbplancta.Seek "=", c_cuenta.txtcuenta.Text
    If Not tbplancta.NoMatch = True Then
        Pnldescripcion = tbplancta.Fields("f5nomcta")
    End If

    Fnummes = Val(c_cuenta.txtmes.Text)
    pnlcuenta.Caption = c_cuenta.txtcuenta.Text
    Select Case Fnummes
        Case 1: F4MES = "ENERO"
        Case 2: F4MES = "FEBRERO"
        Case 3: F4MES = "MARZO"
        Case 4: F4MES = "ABRIL"
        Case 5: F4MES = "MAYO"
        Case 6: F4MES = "JUNIO"
        Case 7: F4MES = "JULIO"
        Case 8: F4MES = "AGOSTO"
        Case 9: F4MES = "SETIEMBRE"
        Case 10: F4MES = "OCTUBRE"
        Case 11: F4MES = "NOVIEMBRE"
        Case 12: F4MES = "DICIEMBRE"
    End Select
    F4MES = F4MES & " - " & wanno
    PnlMes.Caption = F4MES
    If c_cuenta.opcmodo(0) = True Then
        PnlMes.Caption = "MES DE " & PnlMes.Caption
    Else
        PnlMes.Caption = "ACUMULADO AL MES DE " & PnlMes.Caption
    End If
    
    Proceso

    dataconta.DatabaseName = wrutatemp & "\temp_com.Mdb"
    dataconta.RecordSource = "temp_regismov"
    dataconta.Refresh
    If dataconta.Recordset.RecordCount = 0 Then
       MsgBox "No existen registros para procesar la consulta. Verifique.", 48, "Compras"
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo error_bd

    TbCabRegis.Close
    TbOfiRegis.Close
    dbcompras.Close

    tbtemp_regismov.Close
    dbtempcomp.Close

    TbDocumento.Close
    dbcomtabla.Close

    tbplancta.Close
    dbplancta.Close

    Exit Sub

error_bd:
    Resume Next

End Sub

