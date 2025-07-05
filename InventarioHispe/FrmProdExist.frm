VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form FrmProdExist 
   Caption         =   "Productos Existentes"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   2955
   ClientWidth     =   5340
   LinkTopic       =   "Form3"
   ScaleHeight     =   2310
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   90
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   2
      Tools           =   "FrmProdExist.frx":0000
      ToolBars        =   "FrmProdExist.frx":1984
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1500
      Left            =   225
      TabIndex        =   0
      Top             =   315
      Width           =   4740
      _Version        =   65536
      _ExtentX        =   8361
      _ExtentY        =   2646
      _StockProps     =   15
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.Frame Frame1 
         Caption         =   "Seleccion de Productos"
         ForeColor       =   &H80000006&
         Height          =   870
         Left            =   360
         TabIndex        =   1
         Top             =   315
         Width           =   3840
         Begin VB.OptionButton opcion2 
            Caption         =   "Con Stock"
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton opcion2 
            Caption         =   "Todos"
            Height          =   195
            Index           =   1
            Left            =   2475
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "FrmProdExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SwOpcionProd As Integer
Private Sub Form_Load()
Me.Height = 2800
Me.Width = 5460
End Sub

Private Sub opcion2_Click(Index As Integer)
Select Case Index
    Case 0: SwOpcionProd = 0
    Case 1: SwOpcionProd = 1
End Select
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Aceptar"
        Mostrar_Stock
           
    Case "ID_Salir"
        Unload Me
End Select

End Sub

Private Sub Mostrar_Stock()
Set RsProducto = New ADODB.Recordset
Set rsconsulta = New ADODB.Recordset

With Acr_ProdExist

    If rsconsulta.State = adStateOpen Then rsconsulta.Close
    rsconsulta.Open "SELECT * FROM SF1PARAM WHERE F1CODEMP = '" & wempresa & "'", cnn_control, adOpenKeyset, adLockOptimistic
    If Not rsconsulta.EOF Then
       .FldEmpresa.Text = rsconsulta.Fields("F1NOMEMP")
       rsconsulta.Close
    Else
       .FldEmpresa.Text = wempresa
       rsconsulta.Close
    End If
    .FldFecha.Text = Format(Now, "dd/mm/yyyy")
    .LblTitulo.Caption = " REPORTE DE PRODUCTOS EXISTENTES SEGÚN STOCK "
    
    .DataControl1.ConnectionString = DBBANCOWIN
    
    If SwOpcionProd = 0 Then
        SQL = "SELECT * FROM IF5PLA WHERE F5STOCKACT > 0 ORDER BY F5CODPRO"
    ElseIf SwOpcionProd = 1 Then
        SQL = "SELECT * FROM IF5PLA ORDER BY F5CODPRO"
    End If
    
    .DataControl1.Source = SQL

End With
Acr_ProdExist.Show vbModal
End Sub

