VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmResInforme 
   ClientHeight    =   5310
   ClientLeft      =   1725
   ClientTop       =   1095
   ClientWidth     =   7305
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7305
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5670
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   630
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Threed.SSPanel PnlZona 
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   810
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3360
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin Threed.SSPanel PnlFecha 
      Height          =   240
      Left            =   2700
      TabIndex        =   8
      Top             =   495
      Width           =   2490
      _Version        =   65536
      _ExtentX        =   4392
      _ExtentY        =   423
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
      BevelOuter      =   0
   End
   Begin Threed.SSPanel PnlImporta 
      Height          =   3075
      Left            =   270
      TabIndex        =   5
      Top             =   1710
      Visible         =   0   'False
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   5424
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
      BevelInner      =   2
      Begin Threed.SSCommand SSCommand1 
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   135
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "&Salir"
         ForeColor       =   -2147483630
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   135
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid Grid2 
         Bindings        =   "FrmResInforme.frx":0000
         Height          =   2265
         Left            =   180
         OleObjectBlob   =   "FrmResInforme.frx":0014
         TabIndex        =   6
         Top             =   450
         Width           =   6450
      End
      Begin VB.Label LblImporta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         TabIndex        =   7
         Top             =   90
         Width           =   3255
      End
   End
   Begin MSDBGrid.DBGrid grid1 
      Bindings        =   "FrmResInforme.frx":1AB9
      Height          =   4065
      Left            =   90
      OleObjectBlob   =   "FrmResInforme.frx":1ACD
      TabIndex        =   4
      Top             =   1170
      Width           =   7080
   End
   Begin Threed.SSPanel PanelCab 
      Height          =   510
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   900
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      Begin VB.CommandButton BtnExit 
         Height          =   420
         Left            =   495
         Picture         =   "FrmResInforme.frx":29AA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   420
      End
      Begin Threed.SSCommand BtnPrint 
         Height          =   420
         Left            =   45
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   45
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmResInforme.frx":3044
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   6480
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Resumen de Importaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2475
      TabIndex        =   3
      Top             =   135
      Width           =   3345
   End
End
Attribute VB_Name = "FrmResInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TbZona      As DAO.Recordset
Dim TbTmpInf    As DAO.Recordset
Dim WMONEDA     As String
Dim dbinvtemp   As DAO.Database
Dim tbtempimp   As DAO.Recordset
Dim tbtempimdet As DAO.Recordset

Sub Command3D1_Click()
    Me.MousePointer = 11
    Proceso
    Me.MousePointer = 1
End Sub

Sub Proceso()
Dim TbConsulta  As DAO.Recordset
Dim wsaldotot   As Double
Dim BDTEMPO     As DAO.Database
Dim tbtemp      As DAO.Recordset
Dim Msql        As String
    
    Set BDTEMPO = OpenDatabase(wrutatemp & "\temp_com.mdb")
    BDTEMPO.Execute ("delete * from temp_imp")
    BDTEMPO.Close
    Msql = "Select * From Import_Cab Where F4FECHA>=Cvdate('" + FrmRepImporta.TxtDesde.Text + "') and F4FECHA<=Cvdate('" + FrmRepImporta.TxtHasta.Text + "') AND F4CODPRV ='" + FrmRepImporta.TxtCodPrv.Text + "' Order By F4NUMIMP"
    
    Set TbConsulta = dbcompras.OpenRecordset(Msql)
    If TbConsulta.RecordCount = 0 Then
       Exit Sub
    End If
    TbConsulta.MoveFirst
    wsaldotot = 0#
    Do While Not TbConsulta.EOF
       tbtempimp.AddNew
       tbtempimp.Fields("NUMIMP") = "" & TbConsulta.Fields("F4NUMIMP")
       tbtempimp.Fields("FECHA") = "" & TbConsulta.Fields("F4FECHA")
       tbtempimp.Fields("REFERE") = "" & TbConsulta.Fields("F4REFERE")
       tbtempimp.Fields("DOCUM") = "" & TbConsulta.Fields("F4SERIE") & "/" & TbConsulta.Fields("F4NUMFAC")
       tbtempimp.Fields("TOTAL") = Format(TbConsulta.Fields("F4TOTAL"), "###,##0.000")
       wsaldotot = wsaldotot + Val(Format(TbConsulta.Fields("F4TOTAL"), "0.000"))
       TbConsulta.MoveNext
       'GRABA FILA
       tbtempimp.Update
       
       If TbConsulta.EOF Then Exit Do
    Loop
    
    tbtempimp.AddNew
    tbtempimp.Fields("TOTAL") = "========="
    tbtempimp.Update
    
    tbtempimp.AddNew
    tbtempimp.Fields("FECHA") = "Total General"
    tbtempimp.Fields("TOTAL") = Format(wsaldotot, "###,##0.000")
    tbtempimp.Update

End Sub

Sub Vacea_Importacion()
Dim BDTEMP As DAO.Database
Dim tbtemp As DAO.Recordset
    
    Set BDTEMP = OpenDatabase(wrutatemp & "\temp_com.mdb")
    BDTEMP.Execute ("delete * from temp_imdet")
    BDTEMP.Close
  
    TbDetImport.Seek "=", Format(Grid1.Columns(0), "0000000")
    
    If Not TbDetImport.NoMatch Then

       tbtempimdet.AddNew
       Do While TbDetImport.Fields("F4NUMIMP") = Grid1.Columns(0)
          tbtempimdet.Fields("F3NUMORD") = "" & TbDetImport.Fields("F3NUMORD")
          tbtempimdet.Fields("F3CODFAB") = "" & TbDetImport.Fields("F3CODFAB")
          tbtempimdet.Fields("F5CODPRO") = "" & TbDetImport.Fields("F5CODPRO")
          
          TBPRODUCTO.Seek "=", "" & TbDetImport.Fields("F5CODPRO")
          If Not TBPRODUCTO.NoMatch Then
             tbtempimdet.Fields("F5NOMPRO") = "" & TBPRODUCTO.Fields("F5NOMPRO")
          End If
          tbtempimdet.Fields("F3CANTIDAD") = Format(TbDetImport.Fields("F3CANTIDAD"), "###,##0.000")
          tbtempimdet.Fields("F3PREUNI") = Format(TbDetImport.Fields("F3PREUNI"), "###,##0.000")
          tbtempimdet.Fields("F3TOTAL") = Format(TbDetImport.Fields("F3TOTAL"), "###,##0.000")
          tbtempimdet.Fields("F3PRECOS") = Format(TbDetImport.Fields("F3PRECOS"), "###,##0.000")
          tbtempimdet.Fields("F3MARGEN") = Format(TbDetImport.Fields("F3MARGEN"), "###,##0.000")
          tbtempimdet.Fields("F3VALVTA") = Format(TbDetImport.Fields("F3VALVTA"), "###,##0.000")
          tbtempimdet.Fields("F3DSCTO") = Format(TbDetImport.Fields("F3DSCTO"), "###,##0.000")
          tbtempimdet.Fields("F3VTANET") = Format(TbDetImport.Fields("F3VTANET"), "###,##0.000")
          TbDetImport.MoveNext
          tbtempimdet.Update
          If TbDetImport.EOF Then Exit Do
          
          If TbDetImport.Fields("F4NUMIMP") = Grid1.Columns(0) Then

          End If

       Loop
    End If

End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub


Private Sub BtnPrint_Click()
    Report1.DataFiles(0) = wrutatemp & "\temp_com.mdb"
    Report1.ReportFileName = wrutatemp & "\rinfor.rpt"
    Report1.Action = 1
'    Me.MousePointer = 11
'    IMPRIME
    Me.MousePointer = 1
End Sub

Private Sub Form_Activate()
  Grid1.SetFocus
End Sub

Private Sub Form_Load()
    Set dbcompras = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set DbInventa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")

    Set TbCabImport = dbcompras.OpenRecordset("IMPORT_CAB")
    Set TbDetImport = dbcompras.OpenRecordset("IMPORT_DET")

    Set TBPRODUCTO = DbInventa.OpenRecordset("IF5PLA")
    
    Set dvinvtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    Set tbtempimp = dvinvtemp.OpenRecordset("temp_imp")
    
    Set tbtempimdet = dvinvtemp.OpenRecordset("temp_imdet")

    TbDetImport.Index = "IDNUMIMP"
    TBPRODUCTO.Index = "IDCODPRO"
    
    PnlFecha.Caption = "Del  " & FrmRepImporta.TxtDesde.Text & "  AL  " & FrmRepImporta.TxtHasta.Text
    
    Proceso     'lista las importaciones
    
    Data1.DatabaseName = wrutatemp & "\temp_com.mdb"
    Data1.RecordSource = "Select * from temp_imp"
    Data1.Refresh
End Sub

Private Sub Grid1_DblClick()
    If Len(Trim(Grid1.Columns(0))) > 0 Then
       PnlImporta.Visible = True
       LblImporta.Caption = "IMPORTACION  Nº " & Grid1.Columns(0)
       
       Vacea_Importacion    'lista los productos
       
       Data2.DatabaseName = wrutatemp & "\temp_com.mdb"
       Data2.RecordSource = "Select * from temp_imdet"
       Data2.Refresh
    End If
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
    Grid1.Refresh
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
       PnlImporta.Visible = False
    End If
End Sub

Private Sub SSCommand1_Click()
  PnlImporta.Visible = False
End Sub
