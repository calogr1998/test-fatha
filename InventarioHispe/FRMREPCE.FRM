VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmrepce 
   Caption         =   "Reporte por Centro de Costo"
   ClientHeight    =   4125
   ClientLeft      =   1995
   ClientTop       =   1380
   ClientWidth     =   6360
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6360
   Begin Threed.SSCommand cmdsalir 
      Height          =   375
      Left            =   3195
      TabIndex        =   17
      Top             =   3645
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   661
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
   Begin Threed.SSCommand cmdaceptar 
      Height          =   375
      Left            =   1845
      TabIndex        =   16
      Top             =   3645
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   661
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
   Begin Crystal.CrystalReport cryreporte 
      Left            =   495
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   3435
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   6225
      _Version        =   65536
      _ExtentX        =   10980
      _ExtentY        =   6059
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      Begin Threed.SSCheck resumido 
         Height          =   285
         Left            =   360
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Resumido"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkgasto 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   2430
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Por Gasto"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtmes 
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Top             =   90
         Width           =   375
      End
      Begin Threed.SSPanel pnlgasto 
         Height          =   285
         Left            =   2745
         TabIndex        =   13
         Top             =   2295
         Width           =   3300
         _Version        =   65536
         _ExtentX        =   5821
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483630
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
      Begin VB.TextBox txtgasto 
         Height          =   285
         Left            =   1980
         TabIndex        =   12
         Top             =   2295
         Width           =   690
      End
      Begin Threed.SSFrame frmmodo 
         Height          =   510
         Left            =   1980
         TabIndex        =   3
         Top             =   2655
         Visible         =   0   'False
         Width           =   4065
         _Version        =   65536
         _ExtentX        =   7170
         _ExtentY        =   900
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optmodo 
            Caption         =   "Acumulado"
            Height          =   240
            Index           =   1
            Left            =   2475
            TabIndex        =   15
            Top             =   180
            Width           =   1365
         End
         Begin VB.OptionButton optmodo 
            Caption         =   "Mensual"
            Height          =   240
            Index           =   0
            Left            =   585
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin Threed.SSFrame Frame3D1 
         Height          =   555
         Left            =   90
         TabIndex        =   2
         Top             =   1665
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   979
         _StockProps     =   14
         Caption         =   "Moneda"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin VB.OptionButton optmoneda 
            Caption         =   "Dólares"
            Height          =   240
            Index           =   1
            Left            =   4005
            TabIndex        =   11
            Top             =   225
            Width           =   1320
         End
         Begin VB.OptionButton optmoneda 
            Caption         =   "Soles"
            Height          =   195
            Index           =   0
            Left            =   1215
            TabIndex        =   10
            Top             =   225
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin Threed.SSFrame Frame3D2 
         Height          =   1185
         Left            =   90
         TabIndex        =   1
         Top             =   405
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   2090
         _StockProps     =   14
         Caption         =   "Centro de Costo"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin Threed.SSPanel pnlcentro2 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   675
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   503
            _StockProps     =   15
            ForeColor       =   -2147483630
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
            Enabled         =   0   'False
         End
         Begin VB.TextBox txtcentro2 
            Height          =   285
            Left            =   810
            TabIndex        =   8
            Top             =   675
            Width           =   870
         End
         Begin Threed.SSPanel pnlcentro 
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Top             =   315
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   503
            _StockProps     =   15
            ForeColor       =   -2147483630
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
            Enabled         =   0   'False
         End
         Begin VB.TextBox txtcentro 
            Height          =   285
            Left            =   810
            TabIndex        =   5
            Top             =   315
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hasta :"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   675
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde :"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   315
            Width           =   555
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   2700
         TabIndex        =   18
         Top             =   135
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmrepce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbtemp      As DAO.Recordset
Dim xnombre     As String
Dim xx          As String
Dim cabrev      As String
Dim Msql        As String
Dim X           As Integer

Sub PROCESA(xcentro)
Dim TbConsulta  As DAO.Recordset
Dim cmes        As String
Dim xf4mes      As String

    cmes = Trim(txtmes.Text)
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("delete * from temp_centros")
    Set tbtemp = dbtemp.OpenRecordset("temp_centros")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"

    If xx = "NSE" Then
        Msql = "Select * From REGISMOV WHERE F3CENCOS = '" + xcentro + "' Order by F4NUMMOV"
    Else
        Msql = "Select * From REGISMOV WHERE F4MESMOV = '" + txtmes.Text + "' AND F3CENCOS = '" + xcentro + "' Order by F4NUMMOV"
    End If
    Set TbConsulta = dbcompras.OpenRecordset(Msql)

    If TbConsulta.RecordCount > 0 Then
        TbConsulta.MoveFirst
        Do While Not TbConsulta.EOF
            TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
            If Not TbCabRegis.NoMatch Then
                tbtemp.AddNew
                tbtemp.Fields("f3centro") = xcentro
                tbtemp.Fields("f3nomcentro") = xnombre
                tbtemp.Fields("f3nummov") = TbConsulta.Fields("f4nummov")
                tbtemp.Fields("f3fecha") = TbCabRegis.Fields("f4fecha")
                cabrev = ""
                TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc")
                If Not TbDocumento.NoMatch Then
                    cabrev = TbDocumento.Fields("f2abrev") & ""
                End If
                tbtemp.Fields("f3docum") = cabrev & TbCabRegis.Fields("f4serdoc") & "/" & TbCabRegis.Fields("f4numdoc")
                tbtemp.Fields("f3codprov") = TbCabRegis.Fields("f4codprv") & ""
                tbtemp.Fields("f3nomprov") = TbCabRegis.Fields("f4nomprv") & ""
                tbtemp.Fields("f3tipcamb") = TbCabRegis.Fields("f4tipcam")
                If xx = "NSE" Then
                    If optmoneda(0).Value = True Then
                        If TbCabRegis.Fields("f4moneda") = "S" Then
                            tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                            tbtemp.Fields("f3igv") = Format(TbConsulta.Fields("f3importe") * 0.19, "0.00")
                        Else
                            tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                            tbtemp.Fields("f3igv") = Format((TbConsulta.Fields("f3importe") * 0.19) * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        End If
                    Else
                        If TbCabRegis.Fields("f4moneda") = "D" Then
                            tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                            tbtemp.Fields("f3igv") = Format(TbConsulta.Fields("f3importe") * 0.19, "0.00")
                        Else
                            tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                            tbtemp.Fields("f3igv") = Format((TbConsulta.Fields("f3importe") * 0.19) / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        End If
                    End If
                Else
                    If optmoneda(0).Value = True Then
                        If TbCabRegis.Fields("f4moneda") = "S" Then
                            tbtemp.Fields("f3total") = TbCabRegis.Fields("f4total")
                            tbtemp.Fields("f3igv") = TbCabRegis.Fields("f4igv")
                        Else
                            tbtemp.Fields("f3total") = Format(TbCabRegis.Fields("f4total") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                            tbtemp.Fields("f3igv") = Format(TbCabRegis.Fields("f4igv") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        End If
                    Else
                        If TbCabRegis.Fields("f4moneda") = "D" Then
                            tbtemp.Fields("f3total") = TbCabRegis.Fields("f4total")
                            tbtemp.Fields("f3igv") = TbCabRegis.Fields("f4igv")
                        Else
                            tbtemp.Fields("f3total") = Format(TbCabRegis.Fields("f4total") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                            tbtemp.Fields("f3igv") = Format(TbCabRegis.Fields("f4igv") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        End If
                    End If
                End If
                tbtemp.Fields("f3moneda") = IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3simbolo") = IIf(optmoneda(0).Value = True, "S/.", "US$")
                tbtemp.Fields("empresa") = wnomcia
                'selecciona mes
                Select Case cmes
                    Case "01": xf4mes = "ENERO"
                    Case "02": xf4mes = "FEBRERO"
                    Case "03": xf4mes = "MARZO"
                    Case "04": xf4mes = "ABRIL"
                    Case "05": xf4mes = "MAYO"
                    Case "06": xf4mes = "JUNIO"
                    Case "07": xf4mes = "JULIO"
                    Case "08": xf4mes = "AGOSTO"
                    Case "09": xf4mes = "SETIEMBRE"
                    Case "10": xf4mes = "OCTUBRE"
                    Case "11": xf4mes = "NOVIEMBRE"
                    Case "12": xf4mes = "DICIEMBRE"
                End Select
                cmes = xf4mes
                
                'tbtemp.Fields("titulo") = "MES DE " & UCase(dev_mes(cmes)) & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("titulo") = "MES DE " & cmes & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3detalle") = TbCabRegis.Fields("f4refere") & ""
                tbtemp.Update
            End If
            TbConsulta.MoveNext
        Loop
    End If

End Sub
Sub procesa2(xcentro)
Dim TbConsulta As DAO.Recordset
Dim cmes As String
Dim xf4mes As String

    cmes = Trim(txtmes.Text)
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("delete * from temp_centros")
    Set tbtemp = dbtemp.OpenRecordset("temp_centros")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
    tbgastos.Index = "bf9gin"

    Msql = "Select * From REGISMOV WHERE F4MESMOV = '" + txtmes.Text + "' AND F3CENCOS = '" + xcentro + "' Order by F4NUMMOV"
    Set TbConsulta = dbcompras.OpenRecordset(Msql)

    If TbConsulta.RecordCount > 0 Then
        TbConsulta.MoveFirst
        Do While Not TbConsulta.EOF
            TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
            If Not TbCabRegis.NoMatch Then
                tbtemp.AddNew
                tbtemp.Fields("f3centro") = xcentro
                tbtemp.Fields("f3nomcentro") = xnombre
                tbtemp.Fields("f3nummov") = TbConsulta.Fields("f4nummov")
                tbtemp.Fields("f3fecha") = TbCabRegis.Fields("f4fecha")
                tbtemp.Fields("f3nummov") = TbConsulta.Fields("f4nummov")
                cabrev = ""
                TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc")
                If Not TbDocumento.NoMatch Then
                    cabrev = TbDocumento.Fields("f2abrev") & ""
                End If
                tbtemp.Fields("f3docum") = cabrev & TbCabRegis.Fields("f4serdoc") & "/" & TbCabRegis.Fields("f4numdoc")
                tbtemp.Fields("f3codprov") = TbCabRegis.Fields("f4codprv") & ""
                tbtemp.Fields("f3nomprov") = TbCabRegis.Fields("f4nomprv") & ""
                tbtemp.Fields("f3tipcamb") = TbCabRegis.Fields("f4tipcam")
                If optmoneda(0).Value = True Then
                    If TbCabRegis.Fields("f4moneda") = "S" Then
                        tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                        tbtemp.Fields("f3igv") = TbCabRegis.Fields("f4igv")
                    Else
                        tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        tbtemp.Fields("f3igv") = Format(TbCabRegis.Fields("f4igv") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                    End If
                Else
                    If TbCabRegis.Fields("f4moneda") = "D" Then
                        tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                        tbtemp.Fields("f3igv") = TbCabRegis.Fields("f4igv")
                    Else
                        tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        tbtemp.Fields("f3igv") = Format(TbCabRegis.Fields("f4igv") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                    End If
                End If
                tbtemp.Fields("f3moneda") = IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3simbolo") = IIf(optmoneda(0).Value = True, "S/.", "US$")
                tbtemp.Fields("empresa") = wnomcia
                'selecciona mes
                Select Case cmes
                    Case "01": xf4mes = "ENERO"
                    Case "02": xf4mes = "FEBRERO"
                    Case "03": xf4mes = "MARZO"
                    Case "04": xf4mes = "ABRIL"
                    Case "05": xf4mes = "MAYO"
                    Case "06": xf4mes = "JUNIO"
                    Case "07": xf4mes = "JULIO"
                    Case "08": xf4mes = "AGOSTO"
                    Case "09": xf4mes = "SETIEMBRE"
                    Case "10": xf4mes = "OCTUBRE"
                    Case "11": xf4mes = "NOVIEMBRE"
                    Case "12": xf4mes = "DICIEMBRE"
                End Select
                cmes = xf4mes

                'tbtemp.Fields("titulo") = "MES DE " & UCase(dev_mes(cmes)) & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("titulo") = "MES DE " & cmes & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3detalle") = TbCabRegis.Fields("f4refere") & ""
                tbtemp.Fields("f3gasto") = TbConsulta.Fields("f3gasto") & ""
                tbgastos.Seek "=", "G", TbConsulta.Fields("f3gasto") & ""
                If Not tbgastos.NoMatch Then
                    tbtemp.Fields("f3desgto") = tbgastos.Fields("nombre") & ""
                End If
                tbtemp.Update
            End If
            TbConsulta.MoveNext
        Loop
    End If

End Sub
Sub procesa3(xcentro)
Dim TbConsulta As DAO.Recordset
Dim cmes As String
Dim xf4mes As String
    cmes = Trim(txtmes.Text)
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("delete * from temp_centros")
    Set tbtemp = dbtemp.OpenRecordset("temp_centros")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
    tbgastos.Index = "bf9gin"

    Msql$ = "Select * From REGISMOV WHERE F4MESMOV = '" + txtmes.Text + "' AND F3CENCOS = '" + xcentro + "' AND F3GASTO = '" + txtgasto.Text + "' Order by F4NUMMOV"
    Set TbConsulta = dbcompras.OpenRecordset(Msql$)

    If TbConsulta.RecordCount > 0 Then
        TbConsulta.MoveFirst
        Do While Not TbConsulta.EOF
            TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
            If Not TbCabRegis.NoMatch Then
                tbtemp.AddNew
                tbtemp.Fields("f3centro") = xcentro
                tbtemp.Fields("f3nomcentro") = xnombre
                tbtemp.Fields("f3nummov") = TbConsulta.Fields("f4nummov")
                tbtemp.Fields("f3fecha") = TbCabRegis.Fields("f4fecha")
                tbtemp.Fields("f3nummov") = TbConsulta.Fields("f4nummov")
                cabrev = ""
                TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc")
                If Not TbDocumento.NoMatch Then
                    cabrev = TbDocumento.Fields("f2abrev") & ""
                End If
                tbtemp.Fields("f3docum") = cabrev & TbCabRegis.Fields("f4serdoc") & "/" & TbCabRegis.Fields("f4numdoc")
                tbtemp.Fields("f3codprov") = TbCabRegis.Fields("f4codprv") & ""
                tbtemp.Fields("f3nomprov") = TbCabRegis.Fields("f4nomprv") & ""
                tbtemp.Fields("f3tipcamb") = TbCabRegis.Fields("f4tipcam")
                If optmoneda(0).Value = True Then
                    If TbCabRegis.Fields("f4moneda") = "S" Then
                        tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                        tbtemp.Fields("f3igv") = TbCabRegis.Fields("f4igv")
                    Else
                        tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        tbtemp.Fields("f3igv") = Format(TbCabRegis.Fields("f4igv") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                    End If
                Else
                    If TbCabRegis.Fields("f4moneda") = "D" Then
                        tbtemp.Fields("f3total") = TbConsulta.Fields("f3importe")
                        tbtemp.Fields("f3igv") = TbCabRegis.Fields("f4igv")
                    Else
                        tbtemp.Fields("f3total") = Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        tbtemp.Fields("f3igv") = Format(TbCabRegis.Fields("f4igv") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                    End If
                End If
                tbtemp.Fields("f3moneda") = IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3simbolo") = IIf(optmoneda(0).Value = True, "S/.", "US$")
                tbtemp.Fields("empresa") = wnomcia
                'selecciona mes
                Select Case cmes
                    Case "01": xf4mes = "ENERO"
                    Case "02": xf4mes = "FEBRERO"
                    Case "03": xf4mes = "MARZO"
                    Case "04": xf4mes = "ABRIL"
                    Case "05": xf4mes = "MAYO"
                    Case "06": xf4mes = "JUNIO"
                    Case "07": xf4mes = "JULIO"
                    Case "08": xf4mes = "AGOSTO"
                    Case "09": xf4mes = "SETIEMBRE"
                    Case "10": xf4mes = "OCTUBRE"
                    Case "11": xf4mes = "NOVIEMBRE"
                    Case "12": xf4mes = "DICIEMBRE"
                End Select
                cmes = xf4mes
                
                'tbtemp.Fields("titulo") = "MES DE " & UCase(dev_mes(cmes)) & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("titulo") = "MES DE " & cmes & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3detalle") = TbCabRegis.Fields("f4refere") & ""
                tbtemp.Fields("f3gasto") = TbConsulta.Fields("f3gasto") & ""
                tbgastos.Seek "=", TbConsulta.Fields("f3gasto") & ""
                If Not tbgastos.NoMatch Then
                    tbtemp.Fields("f3desgto") = tbgastos.Fields("nombre") & ""
                End If
                tbtemp.Update
            End If
            TbConsulta.MoveNext
        Loop
    End If

End Sub
Sub procesa4(xcentro)
Dim TbConsulta As DAO.Recordset
Dim cmes As String
Dim G As String
Dim TOTAL As Double
Dim xf4mes As String

    cmes = Trim(txtmes.Text)
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("delete * from temp_centros")
    Set tbtemp = dbtemp.OpenRecordset("temp_centros")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
    tbgastos.Index = "bf9gin"

    Msql$ = "Select * From REGISMOV WHERE F4MESMOV = '" + txtmes.Text + "' AND F3CENCOS = '" + xcentro + "' Order by F3GASTO"
    Set TbConsulta = dbcompras.OpenRecordset(Msql$)

    If TbConsulta.RecordCount > 0 Then
        TbConsulta.MoveFirst
        G = "" & Trim(TbConsulta.Fields("f3gasto") & "")
        X = 0
        Do While Not TbConsulta.EOF
            TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
            If Not TbCabRegis.NoMatch Then
                tbtemp.AddNew
                tbtemp.Fields("f3moneda") = IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3simbolo") = IIf(optmoneda(0).Value = True, "S/.", "US$")
                tbtemp.Fields("empresa") = wnomcia
                'selecciona mes
                Select Case cmes
                    Case "01": xf4mes = "ENERO"
                    Case "02": xf4mes = "FEBRERO"
                    Case "03": xf4mes = "MARZO"
                    Case "04": xf4mes = "ABRIL"
                    Case "05": xf4mes = "MAYO"
                    Case "06": xf4mes = "JUNIO"
                    Case "07": xf4mes = "JULIO"
                    Case "08": xf4mes = "AGOSTO"
                    Case "09": xf4mes = "SETIEMBRE"
                    Case "10": xf4mes = "OCTUBRE"
                    Case "11": xf4mes = "NOVIEMBRE"
                    Case "12": xf4mes = "DICIEMBRE"
                End Select
                cmes = xf4mes
                
                'tbtemp.Fields("titulo") = "MES DE " & UCase(dev_mes(cmes)) & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("titulo") = "MES DE " & cmes & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                tbtemp.Fields("f3gasto") = TbConsulta.Fields("f3gasto") & ""
                tbgastos.Seek "=", "G", TbConsulta.Fields("f3gasto") & ""
                If Not tbgastos.NoMatch Then
                    tbtemp.Fields("f3desgto") = tbgastos.Fields("nombre") & ""
                End If
                tbtemp.Fields("f3centro") = xcentro
                tbtemp.Fields("f3nomcentro") = xnombre
                cabrev = ""
                TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc") & ""
                If Not TbDocumento.NoMatch Then
                    cabrev = TbDocumento.Fields("f2abrev") & ""
                End If
                tbtemp.Fields("f3codprov") = TbCabRegis.Fields("f4codprv") & ""
                TOTAL = 0#
                Do While Trim(TbConsulta.Fields("f3gasto") & "") = Trim(G)
                    TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
                    If optmoneda(0).Value = True Then
                        If TbCabRegis.Fields("f4moneda") = "S" Then
                            TOTAL = TOTAL + TbConsulta.Fields("f3importe")
                        Else
                            TOTAL = TOTAL + Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                        End If
                    Else
                        If TbCabRegis.Fields("f4moneda") = "D" Then
                            TOTAL = TOTAL + TbConsulta.Fields("f3importe")
                        Else
                            TOTAL = TOTAL + Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                        End If
                    End If
                    TbConsulta.MoveNext
                    If TbConsulta.EOF Then
                        X = 1
                        Exit Do
                    End If
                Loop
                If X = 0 Then G = "" & Trim(TbConsulta.Fields("f3gasto") & "")
            Else
                TbConsulta.MoveNext
            End If
            tbtemp.Fields("f3total") = TOTAL
            tbtemp.Update
        Loop
    End If

End Sub
Sub procesa5(xcentro)
Dim TbConsulta As DAO.Recordset
Dim cmes As String
Dim G As String
Dim TOTAL As Double
Dim xmes As String
Dim I As Integer
Dim xf4mes As String

    cmes = Trim(txtmes.Text)
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbCabRegis = dbcompras.OpenRecordset("REGISDOC")
    TbCabRegis.Index = "idmesnum"
    Set TbDetRegis = dbcompras.OpenRecordset("REGISMOV")
    
    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.Mdb")
    dbtemp.Execute ("delete * from temp_centros")
    Set tbtemp = dbtemp.OpenRecordset("temp_centros")

    Set dbcomtabla = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbDocumento = dbcomtabla.OpenRecordset("DOCUMENTOS")
    TbDocumento.Index = "idcoddoc"
    Set tbgastos = dbcomtabla.OpenRecordset("BF9GIN")
    tbgastos.Index = "bf9gin"

    For I = 1 To 12
        xmes = Format(I, "00")
        Msql$ = "Select * From REGISMOV WHERE F4MESMOV = '" + xmes + "' AND F3CENCOS = '" + xcentro + "' Order by F3GASTO"
        Set TbConsulta = dbcompras.OpenRecordset(Msql$)
    
        If TbConsulta.RecordCount > 0 Then
            TbConsulta.MoveFirst
            G = "" & Trim(TbConsulta.Fields("f3gasto"))
            X = 0
            Do While Not TbConsulta.EOF
                TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
                If Not TbCabRegis.NoMatch Then
                    tbtemp.Index = "gasto"
                    tbtemp.Seek "=", TbConsulta.Fields("f3gasto")
                    If tbtemp.NoMatch Then
                        tbtemp.AddNew
                        tbtemp.Fields("f3moneda") = IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                        tbtemp.Fields("f3simbolo") = IIf(optmoneda(0).Value = True, "S/.", "US$")
                        tbtemp.Fields("empresa") = wnomcia
                        'selecciona mes
                        Select Case cmes
                            Case "01": xf4mes = "ENERO"
                            Case "02": xf4mes = "FEBRERO"
                            Case "03": xf4mes = "MARZO"
                            Case "04": xf4mes = "ABRIL"
                            Case "05": xf4mes = "MAYO"
                            Case "06": xf4mes = "JUNIO"
                            Case "07": xf4mes = "JULIO"
                            Case "08": xf4mes = "AGOSTO"
                            Case "09": xf4mes = "SETIEMBRE"
                            Case "10": xf4mes = "OCTUBRE"
                            Case "11": xf4mes = "NOVIEMBRE"
                            Case "12": xf4mes = "DICIEMBRE"
                        End Select
                        cmes = xf4mes
                        
                        'tbtemp.Fields("titulo") = "MES DE " & UCase(dev_mes(cmes)) & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                        tbtemp.Fields("titulo") = "MES DE " & cmes & " - " & IIf(optmoneda(0).Value = True, "SOLES", "DOLARES")
                        tbtemp.Fields("f3gasto") = TbConsulta.Fields("f3gasto") & ""
                        tbgastos.Seek "=", "G", TbConsulta.Fields("f3gasto") & ""
                        If Not tbgastos.NoMatch Then
                            tbtemp.Fields("f3desgto") = tbgastos.Fields("nombre") & ""
                        End If
                        tbtemp.Fields("f3centro") = xcentro
                        tbtemp.Fields("f3nomcentro") = xnombre
                        cabrev = ""
                        TbDocumento.Seek "=", TbCabRegis.Fields("f4tipdoc")
                        If Not TbDocumento.NoMatch Then
                            cabrev = TbDocumento.Fields("f2abrev") & ""
                        End If
                        tbtemp.Fields("f3codprov") = TbCabRegis.Fields("f4codprv") & ""
                    Else
                        tbtemp.Edit
                    End If
                    TOTAL = 0#
                    Do While Trim(TbConsulta.Fields("f3gasto") & "") = Trim(G)
                        TbCabRegis.Seek "=", TbConsulta.Fields("f4mesmov"), TbConsulta.Fields("f4nummov")
                        If optmoneda(0).Value = True Then
                            If TbCabRegis.Fields("f4moneda") = "S" Then
                                TOTAL = TOTAL + TbConsulta.Fields("f3importe")
                            Else
                                TOTAL = TOTAL + Format(TbConsulta.Fields("f3importe") * TbCabRegis.Fields("f4tipcam"), "#0.00")
                            End If
                        Else
                            If TbCabRegis.Fields("f4moneda") = "D" Then
                                TOTAL = TOTAL + TbConsulta.Fields("f3importe")
                            Else
                                TOTAL = TOTAL + Format(TbConsulta.Fields("f3importe") / TbCabRegis.Fields("f4tipcam"), "#0.00")
                            End If
                        End If
                        TbConsulta.MoveNext
                        If TbConsulta.EOF Then
                            X = 1
                            Exit Do
                        End If
                    Loop
                    If X = 0 Then G = "" & Trim(TbConsulta.Fields("f3gasto") & "")
                    tbtemp.Fields("f3total" & xmes) = TOTAL
                    tbtemp.Update
                Else
                    TbConsulta.MoveNext
                End If
            Loop
        End If
    Next

End Sub

Private Sub chkgasto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtgasto.SetFocus
    End If
    
End Sub

Private Sub cmdaceptar_Click()
Dim xccosto As String

    xccosto = txtcentro.Text
    If chkgasto.Value = False Then
        If resumido.Value = False Then
            If txtcentro.Text = txtcentro2.Text Then
                xnombre = pnlcentro.Caption
                PROCESA xccosto
                cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                cryreporte.ReportFileName = wrutatemp & "\rpt_cent.rpt"
                cryreporte.Action = 1
            Else
                tbcentros.Index = "centros"
                tbcentros.Seek "=", txtcentro.Text
                If Not tbcentros.NoMatch Then
                    Do While tbcentros.Fields("f3costo") >= txtcentro.Text And tbcentros.Fields("f3costo") <= txtcentro2.Text
                        xnombre = tbcentros.Fields("f3descrip")
                        PROCESA tbcentros.Fields("f3costo")
                        If tbtemp.RecordCount > 0 Then
                            cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                            cryreporte.ReportFileName = wrutatemp & "\rpt_cent.rpt"
                            'cryreporte.Destination = 1
                            cryreporte.Action = 1
                        End If
                        tbcentros.MoveNext
                        If tbcentros.EOF Then Exit Do
                    Loop
                End If
            End If
        Else
            If txtcentro.Text = txtcentro2.Text Then
                xx = "NSE"
                xnombre = pnlcentro.Caption
                PROCESA xccosto
                xx = ""
                cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                cryreporte.ReportFileName = wrutatemp & "\rpt_cent.rpt"
                cryreporte.Action = 1
            End If
        End If
    Else
        If Len(Trim(txtgasto.Text)) = 0 Then
            If resumido.Value = True Then
                If optmodo(0).Value = True Then
                    If txtcentro.Text = txtcentro2.Text Then
                        xnombre = pnlcentro.Caption
                        procesa4 xccosto
                        cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                        cryreporte.ReportFileName = wrutatemp & "\rpt_RES.rpt"
                        cryreporte.Action = 1
                    Else
                        tbcentros.Index = "centros"
                        tbcentros.Seek "=", txtcentro.Text
                        If Not tbcentros.NoMatch Then
                            Do While tbcentros.Fields("f3costo") >= txtcentro.Text And tbcentros.Fields("f3costo") <= txtcentro2.Text
                                xnombre = tbcentros.Fields("f3descrip")
                                procesa4 tbcentros.Fields("f3costo")
                                cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                                cryreporte.ReportFileName = wrutatemp & "\rpt_RES.rpt"
                                'cryreporte.Destination = 1
                                cryreporte.Action = 1
                                tbcentros.MoveNext
                                If tbcentros.EOF Then Exit Do
                            Loop
                        End If
                    End If
                Else
                    If txtcentro.Text = txtcentro2.Text Then
                        xnombre = pnlcentro.Caption
                        procesa5 xccosto
                        cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                        cryreporte.ReportFileName = wrutatemp & "\rpt_resa.rpt"
                        cryreporte.Action = 1
                    Else
                        tbcentros.Index = "centros"
                        tbcentros.Seek "=", txtcentro.Text
                        If Not tbcentros.NoMatch Then
                            Do While tbcentros.Fields("f3costo") >= txtcentro.Text And tbcentros.Fields("f3costo") <= txtcentro2.Text
                                xnombre = tbcentros.Fields("f3descrip")
                                procesa5 tbcentros.Fields("f3costo")
                                cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                                cryreporte.ReportFileName = wrutatemp & "\rpt_resa.rpt"
                                'cryreporte.Destination = 1
                                cryreporte.Action = 1
                                tbcentros.MoveNext
                                If tbcentros.EOF Then Exit Do
                            Loop
                        End If
                    End If
                End If
            Else
                If txtcentro.Text = txtcentro2.Text Then
                    xnombre = pnlcentro.Caption
                    procesa2 xccosto
                    cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                    cryreporte.ReportFileName = wrutatemp & "\cent_gto.rpt"
                    cryreporte.Action = 1
                Else
                    tbcentros.Index = "centros"
                    tbcentros.Seek "=", txtcentro.Text
                    If Not tbcentros.NoMatch Then
                        Do While tbcentros.Fields("f3costo") >= txtcentro.Text And tbcentros.Fields("f3costo") <= txtcentro2.Text
                            xnombre = tbcentros.Fields("f3descrip")
                            procesa2 tbcentros.Fields("f3costo")
                            cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                            cryreporte.ReportFileName = wrutatemp & "\cent_gto.rpt"
                            'cryreporte.Destination = 1
                            cryreporte.Action = 1
                            tbcentros.MoveNext
                            If tbcentros.EOF Then Exit Do
                        Loop
                    End If
                End If
            End If
        Else
            If txtcentro.Text = txtcentro2.Text Then
                xnombre = pnlcentro.Caption
                procesa3 xccosto
                cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                cryreporte.ReportFileName = wrutatemp & "\cent_cod.rpt"
                cryreporte.Action = 1
            Else
                tbcentros.Index = "centros"
                tbcentros.Seek "=", txtcentro.Text
                If Not tbcentros.NoMatch Then
                    Do While tbcentros.Fields("f3costo") >= txtcentro.Text And tbcentros.Fields("f3costo") <= txtcentro2.Text
                        xnombre = tbcentros.Fields("f3descrip")
                        procesa3 tbcentros.Fields("f3costo")
                        cryreporte.DataFiles(0) = wrutatemp & "\temp_com.mdb"
                        cryreporte.ReportFileName = wrutatemp & "\cent_cod.rpt"
                        'cryreporte.Destination = 1
                        cryreporte.Action = 1
                        tbcentros.MoveNext
                        If tbcentros.EOF Then Exit Do
                    Loop
                End If
            End If
        End If
    End If

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Set dbcentros = OpenDatabase(wrutaconta & "\db_costo.mdb")
    Set tbcentros = dbcentros.OpenRecordset("centros")
    txtmes.Text = Format(mes, "00")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set dbcompras = Nothing
    Set dbtemp = Nothing
    Set dbcomtabla = Nothing
    Set dbcentros = Nothing

End Sub

Private Sub optmoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkgasto.SetFocus
    End If

End Sub

Private Sub resumido_Click(Value As Integer)
    If resumido.Value = True Then
        frmmodo.Visible = True
    Else
        frmmodo.Visible = False
    End If
End Sub

Private Sub txtcentro_DblClick()
    txtcentro_KeyDown 113, 0
End Sub

Private Sub txtcentro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        gcodcen = txtcentro.Text
        frmayucen.Show 1
        If Len(Trim(gcodcen)) <> 0 Then
            txtcentro.Text = gcodcen
            pnlcentro.Caption = Trim(gnomcen)
        End If
        txtcentro_KeyPress 13
    End If

End Sub

Private Sub txtcentro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem txtcentro.Text = Format(txtcentro.Text, "000000")
        tbcentros.Index = "centros"
        tbcentros.Seek "=", txtcentro.Text
        If Not tbcentros.NoMatch Then
            pnlcentro.Caption = tbcentros.Fields("f3descrip") & ""
            txtcentro2.Text = txtcentro.Text
            txtcentro2.SetFocus
        Else
            MsgBox "Código del centro de costo no existe. Verifique.", 48, "Compras"
            txtcentro.SetFocus
        End If
    End If

End Sub

Private Sub txtcentro2_DblClick()
    txtcentro2_KeyDown 113, 0
End Sub

Private Sub txtcentro2_GotFocus()
    txtcentro2.SelStart = 0: txtcentro2.SelLength = Len(txtcentro2.Text)
End Sub

Private Sub txtcentro2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        gcodcen = txtcentro2.Text
        frmayucen.Show 1
        If Len(Trim(gcodcen)) <> 0 Then
            txtcentro2.Text = gcodcen
            pnlcentro2.Caption = Trim(gnomcen)
        End If
        txtcentro2_KeyPress 13
    End If

End Sub

Private Sub txtcentro2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem txtcentro2.Text = Format(txtcentro2.Text, "000000")
        tbcentros.Index = "centros"
        tbcentros.Seek "=", txtcentro2.Text
        If Not tbcentros.NoMatch Then
            pnlcentro2.Caption = tbcentros.Fields("f3descrip") & ""
            optmoneda(0).SetFocus
        Else
            MsgBox "Código del centro de costo no existe. Verifique.", 48, "Compras"
            txtcentro2.SetFocus
        End If
    End If

End Sub

Private Sub txtgasto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtgasto.Text)) = 0 Then
            pnlgasto.Caption = "TODOS"
        Else
        End If
        cmdaceptar.SetFocus
    End If

End Sub

Private Sub txtmes_GotFocus()
    txtmes.SelStart = 0: txtmes.SelLength = Len(txtmes.Text)
End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtmes.Text) >= 1 And Val(txtmes.Text) <= 12 Then
            txtmes.Text = Format(txtmes.Text, "00")
            txtcentro.SetFocus
        Else
            MsgBox "El mes es incorrecto. Verifique.", 48, "Compras"
            txtmes.SetFocus
        End If
    End If

End Sub
