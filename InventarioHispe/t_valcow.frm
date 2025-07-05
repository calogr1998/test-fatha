VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form FrmValeCompra 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso por Compra"
   ClientHeight    =   5595
   ClientLeft      =   1095
   ClientTop       =   1035
   ClientWidth     =   9660
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   9660
   Begin VB.Data DataTipCam 
      Caption         =   "DataDetalle"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7155
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1125
      Visible         =   0   'False
      Width           =   1185
   End
   Begin Threed.SSPanel PanelCab 
      Height          =   420
      Left            =   -90
      TabIndex        =   11
      Top             =   0
      Width           =   3240
      _Version        =   65536
      _ExtentX        =   5715
      _ExtentY        =   741
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
      Begin VB.CommandButton BtnDel 
         Height          =   320
         Left            =   1710
         Picture         =   "t_valcow.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Eliminar"
         Top             =   45
         Width           =   360
      End
      Begin VB.CommandButton Command3D1 
         Height          =   320
         Left            =   2115
         Picture         =   "t_valcow.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Calculadora"
         Top             =   45
         Width           =   360
      End
      Begin VB.CommandButton BtnExit 
         Height          =   320
         Left            =   2655
         Picture         =   "t_valcow.frx":0494
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir"
         Top             =   45
         Width           =   360
      End
      Begin VB.CommandButton BtnConsul 
         Height          =   320
         Left            =   1305
         Picture         =   "t_valcow.frx":05DE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Consulta"
         Top             =   45
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton BtnPrint 
         Height          =   320
         Left            =   900
         Picture         =   "t_valcow.frx":06C8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir"
         Top             =   45
         Width           =   360
      End
      Begin VB.CommandButton BtnGraba 
         Height          =   320
         Left            =   495
         Picture         =   "t_valcow.frx":0BFA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Grabar"
         Top             =   45
         Width           =   360
      End
      Begin VB.CommandButton BtnAdd 
         Height          =   320
         Left            =   90
         Picture         =   "t_valcow.frx":112C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Nuevo"
         Top             =   45
         Width           =   360
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5010
      Left            =   0
      TabIndex        =   19
      Top             =   495
      Width           =   9645
      _Version        =   65536
      _ExtentX        =   17013
      _ExtentY        =   8837
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
      Begin MSMask.MaskEdBox txtfecmov 
         Height          =   285
         Left            =   7560
         TabIndex        =   1
         Top             =   180
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txtsergui 
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
         Left            =   4995
         MaxLength       =   11
         TabIndex        =   6
         Top             =   1350
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Data Datadetalle 
         Caption         =   "DataDetalle"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4230
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3915
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Txtnumord 
         Height          =   285
         Left            =   4005
         TabIndex        =   3
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox Txttipcam 
         Alignment       =   1  'Right Justify
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
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1755
         Width           =   600
      End
      Begin VB.PictureBox PanelFac 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6750
         ScaleHeight     =   510
         ScaleWidth      =   2445
         TabIndex        =   20
         Top             =   1215
         Width           =   2445
         Begin VB.TextBox Txtserfac 
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
            Left            =   855
            MaxLength       =   11
            TabIndex        =   33
            Top             =   135
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Txtnumfac 
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
            Left            =   1260
            MaxLength       =   11
            TabIndex        =   21
            Top             =   135
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Factura:"
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
            Left            =   90
            TabIndex        =   22
            Top             =   135
            Visible         =   0   'False
            Width           =   630
         End
      End
      Begin VB.ComboBox Cmbmoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1755
         Width           =   1815
      End
      Begin VB.TextBox Txtcodprov 
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
         Left            =   1260
         MaxLength       =   11
         TabIndex        =   4
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox Txtcodalm 
         Height          =   285
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   0
         Top             =   270
         Width           =   465
      End
      Begin VB.TextBox Txtnumgui 
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
         Left            =   5400
         MaxLength       =   11
         TabIndex        =   7
         Top             =   1350
         Width           =   1005
      End
      Begin VB.ComboBox CmbDocum 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1350
         Width           =   1815
      End
      Begin VB.TextBox Txtnummov 
         Height          =   285
         Left            =   1305
         MaxLength       =   8
         TabIndex        =   2
         Top             =   630
         Width           =   1050
      End
      Begin MSComDlg.CommonDialog caja 
         Left            =   8640
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueDBGrid50.TDBGrid GridDetalle 
         Bindings        =   "t_valcow.frx":165E
         Height          =   2670
         Left            =   90
         OleObjectBlob   =   "t_valcow.frx":1678
         TabIndex        =   10
         Top             =   2250
         Width           =   9510
      End
      Begin Threed.SSPanel Txtnomalm 
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Top             =   270
         Width           =   4605
         _Version        =   65536
         _ExtentX        =   8123
         _ExtentY        =   503
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
      End
      Begin Threed.SSPanel Txtnomprov 
         Height          =   285
         Left            =   2655
         TabIndex        =   24
         Top             =   990
         Width           =   3750
         _Version        =   65536
         _ExtentX        =   6615
         _ExtentY        =   503
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
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Left            =   6930
         TabIndex        =   34
         Top             =   225
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Importación :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   32
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
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
         Index           =   2
         Left            =   360
         TabIndex        =   31
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   7
         Left            =   360
         TabIndex        =   30
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Index           =   9
         Left            =   405
         TabIndex        =   29
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "T.Cambio:"
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
         Index           =   3
         Left            =   4140
         TabIndex        =   28
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label LblDocum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Guía:"
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
         Left            =   4140
         TabIndex        =   27
         Top             =   1395
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Documento"
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
         Index           =   1
         Left            =   360
         TabIndex        =   26
         Top             =   1395
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nro. Vale"
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
         Index           =   0
         Left            =   405
         TabIndex        =   25
         Top             =   675
         Width           =   660
      End
   End
   Begin VB.Menu menuregis 
      Caption         =   "&Registro"
      Begin VB.Menu menuadd 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu menugraba 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu menumodi 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu menudel 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu raya01 
         Caption         =   "-"
      End
      Begin VB.Menu menuprint 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu raya02 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu menuedit 
      Caption         =   "&Edición"
      Begin VB.Menu menubox 
         Caption         =   "&Ver ToolBox"
         Checked         =   -1  'True
      End
      Begin VB.Menu raya03 
         Caption         =   "-"
      End
      Begin VB.Menu menucontrol 
         Caption         =   "&Consultas"
      End
   End
   Begin VB.Menu menu300 
      Caption         =   "&Items"
      Visible         =   0   'False
      Begin VB.Menu menu310 
         Caption         =   "&Lista de Productos"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu menu310 
         Caption         =   "&Agrega Items"
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu menu310 
         Caption         =   "&Elinima Items"
         Index           =   2
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "FrmValeCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DbTempFac           As DAO.Database
Dim tbl                 As New TableDef
Dim Fld                 As Field
Dim Tabla               As DAO.Recordset
Dim TbOrigen            As DAO.Recordset
Dim TbProdProv          As DAO.Recordset
Dim wgraba              As Integer
Dim wvalor              As Integer
Dim wmes                As String
Dim wedit               As Integer
Dim wadic               As Integer
Dim wultinv             As Variant
Dim TbTempfac           As DAO.Recordset

Private Sub Actualiza_Datos()
    
    Tbproveedor.Index = "IDCODPROV"

    Txtnummov.Enabled = True
    Txtnummov.Text = "" & TbStockCab.Fields("f4numval")
    txtfecmov.Text = Format(TbStockCab.Fields("F4Fecval"), "dd/mm/yyyy")
    Txtcodalm.Text = "" & TbStockCab.Fields("F2CODALM")
    Txtcodprov.Text = "" & TbStockCab.Fields("F2CODPROV")
    CmbDocum.ListIndex = Val(Format(TbStockCab.Fields("F1coddoc"), "0"))

    Txtsergui.Text = Left("" & TbStockCab.Fields("F4NUMDOC"), 3)
    Txtnumgui.Text = Right("" & TbStockCab.Fields("F4NUMDOC"), 7)
    Txtserfac.Text = Left("" & TbStockCab.Fields("F4NUMFAC"), 3)
    Txtnumfac.Text = Right("" & TbStockCab.Fields("F4NUMFAC"), 7)
    Txtnumord.Text = Format(TbStockCab.Fields("F4NUMIMP"), "0000000")
    Cmbmoneda.ListIndex = IIf(TbStockCab.Fields("f4moneda") = "S", 0, 1)
    Txttipcam.Text = Val(Format(TbStockCab.Fields("F4tipcam"), "#0.000"))
        
    DbTempFac.Execute ("Delete From TempStock")
    Datadetalle.Refresh
    
    TbStockDet.Index = "IDALMVAL"
    TbStockDet.Seek "=", Txtcodalm.Text, Txtnummov.Text
    If Not TbStockDet.NoMatch Then
        Do While TbStockDet.Fields("F2codalm") = Txtcodalm.Text And TbStockDet.Fields("F4numval") = Txtnummov.Text
            Datadetalle.Recordset.AddNew
            Datadetalle.Recordset.Fields("f5codpro") = "" & TbStockDet.Fields("f5codpro")
            TBPRODUCTO.Index = "IDCODPRO"
            TBPRODUCTO.Seek "=", TbStockDet.Fields("f5codpro")
            If Not TBPRODUCTO.NoMatch Then
                Datadetalle.Recordset.Fields("f5codfab") = "" & TBPRODUCTO.Fields("f5modelo")
                Datadetalle.Recordset.Fields("f5nompro") = "" & TBPRODUCTO.Fields("f5nompro")
                Datadetalle.Recordset.Fields("f7sigmed") = "" & TBPRODUCTO.Fields("F7codmed")
                TbMedida.Seek "=", TBPRODUCTO.Fields("F7codmed")
                If Not TbMedida.NoMatch Then Datadetalle.Recordset.Fields("f7sigmed") = TbMedida.Fields("F7sigmed")
            End If
            Datadetalle.Recordset.Fields("f6canmov") = Val(Format(TbStockDet.Fields("f3canpro"), "#0.000"))
            If Cmbmoneda.ListIndex = 0 Then
                Datadetalle.Recordset.Fields("f6valpro") = Val(Format(TbStockDet.Fields("f3valvta"), "#0.000"))
                Datadetalle.Recordset.Fields("f6total") = Val(Format(TbStockDet.Fields("f3totite"), "#0.000"))
            Else
                Datadetalle.Recordset.Fields("f6valpro") = Val(Format(TbStockDet.Fields("f3valdol"), "#0.000"))
                Datadetalle.Recordset.Fields("f6total") = Val(Format(TbStockDet.Fields("f3totdol"), "#0.000"))
            End If
            Datadetalle.Recordset.Update
            TbStockDet.MoveNext
            If TbStockDet.EOF Then Exit Do
        Loop
    Else
        Nuevo_Item
    End If
    Datadetalle.Refresh
    GridDetalle.Enabled = False: Txtcodalm.Enabled = False
    txtfecmov.Enabled = False: Cmbmoneda.Enabled = False
    Txttipcam.Enabled = True: 'TxtNumMov.Enabled = False
    wgraba = 1
  
End Sub

Private Sub Agrega_Items()

    If GridDetalle.VisibleRows = 1 Then
      If Len(Trim(GridDetalle.Columns(0))) = 0 Then
        DbTempFac.Execute ("Delete From TempStock")
        Datadetalle.Refresh
      End If
    End If
    'If FrmHelpPrvPro.DataGrid.SelCount > 0 Then
    '    For i% = 0 To FrmHelpPrvPro.DataGrid.SelCount - 1
    '        FrmHelpPrvPro.DataGrid.Row = FrmHelpPrvPro.DataGrid.SelRow(i%)
            TBPRODUCTO.Seek "=", Trim(FrmHelpPrvPro.DataGrid.Columns(0))
            If Not TBPRODUCTO.NoMatch Then
                Datadetalle.Recordset.AddNew
                Datadetalle.Recordset.Fields("f5codpro") = "" & FrmHelpPrvPro.DataGrid.Columns(0)
                Datadetalle.Recordset.Fields("f5codfab") = "" & TBPRODUCTO.Fields("F5modelo")
                Datadetalle.Recordset.Fields("f5nompro") = "" & FrmHelpPrvPro.DataGrid.Columns(2)
                Datadetalle.Recordset.Fields("f6canmov") = 0#
                Datadetalle.Recordset.Fields("f7sigmed") = TBPRODUCTO.Fields("F7codmed")
                Datadetalle.Recordset.Fields("f6valpro") = Val(Format(FrmHelpPrvPro.DataGrid.Columns(3), "#0.00"))
                Datadetalle.Recordset.Fields("f6total") = 0#
                Datadetalle.Recordset.Update
            End If
        'Next i%
        Datadetalle.Refresh
    'Else
    '    GridDetalle.Refresh
    '    Nuevo_Item
    'End If
    'If GridDetalle.Rows > 0 Then
    '    GridDetalle.Row = 0: GridDetalle.Col = 2
    'End If

End Sub


Private Sub BtnAdd_Click()
    
    If wgraba = 0 Then
        Select Case MsgBox("El registro no ha sido grabado. Grabar ahora?..", 35, "Inventario")
        Case Is = 6: BtnGraba_Click
        Case Is = 2: Exit Sub
        End Select
    End If
    Txtcodprov.Enabled = True
    Txtnumord.Enabled = True
    wadic = 1
    Me.MousePointer = 11
    Nuevo_Datos
    Me.MousePointer = 1
    
End Sub

Private Sub BtnConsul_Click()

    Me.MousePointer = 11
    Gtipval = "0"
    gcodalm = Txtcodalm.Text
    gnumval = Txtnummov.Text
    FrmAyudaVale.Show 1
    Txtcodalm.Text = gcodalm
    Txtnummov.Text = gnumval
    Txtnummov_Keypress 13
    Me.MousePointer = 1

End Sub

Private Sub BtnDel_Click()

    Me.MousePointer = 11
    Elimina_Movimientos
    LIMPIA_DATOS
    Me.MousePointer = 1

End Sub

Private Sub Btnedit_Click()

    wedit = 0
    wgraba = 1
    
    
    If CVDate(wultinv) >= CVDate(txtfecmov.Text) Then
        Exit Sub
    End If

    TbStockCab.Index = "IDALMVAL"
    TbStockCab.Seek "=", Txtcodalm.Text, Txtnummov.Text
    If Not TbStockCab.NoMatch Then
        If MsgBox("Desea modificar el detalle del vale?..", 36, "Inventario") = 6 Then
            GridDetalle.Enabled = True: Txtcodalm.Enabled = True
            txtfecmov.Enabled = True: Cmbmoneda.Enabled = True
            Txttipcam.Enabled = True
            wedit = 1
            wgraba = 0
        End If
    End If

End Sub

Private Sub BtnExit_Click()

    If wgraba = 0 Then
       Select Case MsgBox("El registro no ha sido grabado. Grabar ahora?..", 35, "Inventario")
       Case Is = 6: BtnGraba_Click
       Case Is = 2: Exit Sub
       End Select
    End If
    Unload Me

End Sub

Private Sub BtnGraba_Click()
    
        Me.MousePointer = 11
        wedit = 1
        If wedit = 1 Then
            TbStockDet.Index = "IDALMVAL"
            TbStockDet.Seek "=", Txtcodalm.Text, Txtnummov.Text
            If Not TbStockDet.NoMatch Then
                Do While TbStockDet.Fields("F2codalm") = Txtcodalm.Text And TbStockDet.Fields("F4numval") = Txtnummov.Text
                    Reactualiza_Almacenes Trim(TbStockDet.Fields("F2codalm")), Trim(TbStockDet.Fields("F5codpro")), Val(Format(TbStockDet.Fields("F3canpro"), "#0.000")), CDate(TbStockDet.Fields("F4fecval")), Val(Format(TbStockDet.Fields("F3totite"), "#0.000")), Val(Format(TbStockDet.Fields("F3totdol"), "#0.000")), "S"
                    TbStockDet.MoveNext
                    If TbStockDet.EOF Then Exit Do
                Loop
                DbInventa.Execute "Delete From IF3VALES where F2codalm = '" + Txtcodalm.Text + "' and F4numval='" + Txtnummov.Text + "'"
            End If
        End If
        GRABA_DATOS

        Me.MousePointer = 1
        'For i% = 0 To GridDetalle.Rows - 1
           'GridDetalle.Row = i%
        TbTempfac.MoveFirst
        Do While Not TbTempfac.EOF
           TBPRODUCTO.Index = "idcodpro"
           TBPRODUCTO.Seek "=", TbTempfac!F5CODPRO
           If Not TBPRODUCTO.NoMatch Then
               If TBPRODUCTO.Fields("f5series") = "1" Then
                    gcodpro = TbTempfac!F5CODPRO
                    frmingseries.Show 1
               End If
           End If
           TbTempfac.MoveNext
        Loop
        'Next i%
   
End Sub

Private Sub BtnPrint_Click()

'On Error GoTo fin
    caja.CancelError = True
    caja.Action = 5
    Me.MousePointer = 11
    Imprime_Vale Trim(Txtcodalm.Text), Trim(Txtnummov.Text), 1
    Me.MousePointer = 1
'fin: Exit Sub

End Sub

Private Sub Calcula_Numero()
    
    Dim wcont   As String
    Dim WNUMERO As String
    
    TBALMACEN.Seek "=", Txtcodalm.Text
    If Not TBALMACEN.NoMatch Then
        wmes = Format(Month(txtfecmov.Text), "00")
        wcont = Format(Val(Mid(TBALMACEN.Fields("F1valing" & wmes), 5, 4)) + 1, "0000")
        WNUMERO = Mid(TBALMACEN.Fields("F1valing" & wmes), 1, 4) & wcont
        Txtnummov.Text = WNUMERO
    End If
    
End Sub

Private Sub CmbDocum_Change()
    wgraba = 0
End Sub

Private Sub CmbDocum_Click()
    
    Select Case CmbDocum.ListIndex
    Case 0: LblDocum.Caption = "Guía:"
            Txtsergui.Visible = True
            Txtsergui.Text = ""
            PanelFac.Visible = True
    Case 1, 2: LblDocum.Caption = Left(CmbDocum, 7) & ":"
            Txtsergui.Visible = True
            Txtsergui.Text = ""
            PanelFac.Visible = False
    Case 3: LblDocum.Caption = "Vale Int.:"
            Txtsergui.Visible = False
            Txtsergui.Text = "V.I"
            Txtnumgui.Text = Right(Txtnummov.Text, 6)
            PanelFac.Visible = False
    Case 4: LblDocum.Caption = "Orden Prod.:"
            Txtsergui.Visible = False
            Txtsergui.Text = ""
            Txtnumgui.Text = ""
            PanelFac.Visible = False
    End Select
    
End Sub

Private Sub cmbdocum_KeyPress(keyascii As Integer)
    
    If keyascii = 13 Then
        CmbDocum_Click
        If keyascii = 13 Then SendKeys "{TAB}"
    End If

End Sub

Private Sub Cmbmoneda_Change()
    wgraba = 0
End Sub

Private Sub Cmbmoneda_KeyPress(keyascii As Integer)
   
   If keyascii = 13 Then SendKeys "{tab}"
    
End Sub

Private Sub Datos_Importa()

    Dim TBCONSULTA As Recordset
    DbTempFac.Execute ("Delete From TempStock")

    Msql$ = "Select * From IMPORT_DET WHERE F4NUMIMP = '" + Txtnumord.Text + "' "
    Set TBCONSULTA = dbcompras.OpenRecordset(Msql$)

    Cmbmoneda.ListIndex = 1
    txtfecmov.Text = FrmImporta.TxtFecha.Text
    Txtcodprov.Text = FrmImporta.TxtRucPrv.Text
    Txtnomprov.Caption = FrmImporta.TxtNomPrv.Text

    Do While Not TBCONSULTA.EOF
       Datadetalle.Recordset.AddNew
       Datadetalle.Recordset.Fields("F5CODPRO") = Trim(TBCONSULTA.Fields("F5CODPRO"))
       TBPRODUCTO.Index = "IDCODPRO"
       TBPRODUCTO.Seek "=", TBCONSULTA.Fields("F5CODPRO")
       If Not TBPRODUCTO.NoMatch Then
          Datadetalle.Recordset.Fields("F5CODFAB") = "" & TBPRODUCTO.Fields("F5MODELO")
          Datadetalle.Recordset.Fields("F5NOMPRO") = "" & TBPRODUCTO.Fields("F5NOMPRO")
          Datadetalle.Recordset.Fields("F7SIGMED") = "" & TBPRODUCTO.Fields("F7CODMED")
       End If
       Datadetalle.Recordset.Fields("F6CANMOV") = Val(Format(TBCONSULTA.Fields("F3CANTIDAD"), "#0.00"))
       Datadetalle.Recordset.Fields("F6VALPRO") = Val(Format(TBCONSULTA.Fields("F3PRECOS"), "#0.000"))
       Datadetalle.Recordset.Fields("F6TOTAL") = Val(Format(TBCONSULTA.Fields("F3CANTIDAD"), "##0.00")) * Val(Format(TBCONSULTA.Fields("F3PRECOS"), "#0.000"))
       Datadetalle.Recordset.Update
       TBCONSULTA.MoveNext
    Loop
    Datadetalle.Refresh
    TBCONSULTA.Close

End Sub

Private Sub Elimina_Movimientos()
    
    Beep
    TbMovAlmacen.Index = "IDALMPROTRA"
    TbStockCab.Index = "IDALMVAL"
    TbStockCab.Seek "=", Txtcodalm.Text, Txtnummov.Text
    If Not TbStockCab.NoMatch Then
        If MsgBox("Está seguro de eliminar los movimientos registrados", 36, "Atención") = 6 Then
            TbStockCab.Delete
            TbStockDet.Index = "IDALMVAL"
            TbStockDet.Seek "=", Txtcodalm.Text, Txtnummov.Text
            If Not TbStockDet.NoMatch Then
                Do While TbStockDet.Fields("F2codalm") = Txtcodalm.Text And TbStockDet.Fields("F4numval") = Txtnummov.Text
                    Reactualiza_Almacenes Trim(Txtcodalm.Text), Trim(TbStockDet.Fields("F5codpro")), Val(Format(TbStockDet.Fields("F3canpro"), "#0.000")), CVDate(TbStockDet.Fields("F4fecval")), Val(Format(TbStockDet.Fields("F3totite"), "#0.000")), Val(Format(TbStockDet.Fields("F3totdol"), "#0.000")), "S"
                    TbStockDet.MoveNext
                    If TbStockDet.NoMatch Then Exit Do
                Loop
                DbInventa.Execute "Delete From IF3VALES where F2codalm='" + Txtcodalm.Text + "' and F4numval='" + Txtnummov.Text + "'"
                DbInventa.Execute "Delete From IF3SERIES where F2codalm='" + Txtcodalm.Text + "' and F4numval='" + Txtnummov.Text + "'"
            End If
            LIMPIA_DATOS
            Txtcodalm.SetFocus
        End If
    End If

End Sub

Private Sub Eliminar_Item()

    Beep
    If MsgBox("Esta Seguro de eliminar el Item", 36, "Atención") = 6 Then
        TbStockDet.Index = "IDALMVALPRO"
        TbStockDet.Seek "=", Txtcodalm.Text, Txtnummov.Text, Trim(GridDetalle.Columns(0))
        If Not TbStockDet.NoMatch Then
            TbStockDet.Delete
        End If
        Datadetalle.Recordset.Delete
        Datadetalle.Refresh
    End If
    If GridDetalle.VisibleRows = 0 Then
      Nuevo_Item

    End If

End Sub

Private Sub Form_Load()
    
    Set DbInventa = OpenDatabase(wrutabancos & "\DB_BANCOS.MDB")
    Set TbStockCab = DbInventa.OpenRecordset("IF4VALES")
    Set TbStockDet = DbInventa.OpenRecordset("IF3VALES")
    Set TBALMACEN = DbInventa.OpenRecordset("EF2ALMACENES")
    Set TbMovAlmacen = DbInventa.OpenRecordset("IF6ALMA")
    Set TBPRODUCTO = DbInventa.OpenRecordset("IF5PLA")
    Set Tbproveedor = DbInventa.OpenRecordset("EF2PROVEEDORES")
    
    Set TbProdProv = DbInventa.OpenRecordset("EF2PROD_PROV")

    DataTipCam.DatabaseName = wrutabancos & "\FACTURAS.MDB"
    DataTipCam.RecordSource = "Select * From IF6TIPCAM Order By F6FECHA"
    DataTipCam.Refresh

    Set DbTempFac = OpenDatabase(wrutatemp & "\TempFac.mdb")
    Set TbTempfac = DbTempFac.OpenRecordset("tempstock")
       
    Datadetalle.DatabaseName = wrutatemp & "\TempFac.Mdb"
    Datadetalle.RecordSource = "TempStock"
    Datadetalle.Refresh
    Nuevo_Item

    TBPRODUCTO.Index = "IDCODPRO"
    TBALMACEN.Index = "IDCODALM"
    Tbproveedor.Index = "IDCODPROV"
    
    TbProdProv.Index = "IDPROPRV"
    
    CmbDocum.Clear
    CmbDocum.AddItem "Guía de Remisión"
    CmbDocum.AddItem "Factura"
    CmbDocum.AddItem "Boleta de Venta"
    CmbDocum.AddItem "Vale Interno"
    CmbDocum.AddItem "Producción"
    CmbDocum.ListIndex = 3

    Cmbmoneda.Clear
    Cmbmoneda.AddItem "Soles"
    Cmbmoneda.AddItem "Dólares"
    Cmbmoneda.ListIndex = 0

    'txtfecmov.Text = Format(gfecha, "dd/mm/yyyy")
    Txttipcam.Text = Format(gtipcam, "###,##0.00")
    gcodori = "XC0"
    wgraba = 1
    wvalor = 0
    wedit = 0

    Txtnumord.Text = Importaciones.txtnumero.Text

    TbStockCab.Index = "IDNUMIMP"
    TbStockCab.Seek "=", Txtnumord.Text
    If Not TbStockCab.NoMatch Then
       If Left(TbStockCab.Fields("F4NUMVAL"), 1) = "I" Then
          Actualiza_Datos
       End If
    Else
       Datos_Importa
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    COMPRAS = 0
    DbTempFac.Execute ("Delete From TempStock")
    Datadetalle.Refresh
    TbTempfac.Close
    DbTempFac.Close
    'Tbproducto.Close
    TbStockCab.Close
    TbStockDet.Close
    'TbMedida.Close
    'TbProveedor.Close
    TBALMACEN.Close
    TbMovAlmacen.Close
    
    
    TbProdProv.Close
    'DBCOMPRAS.Close
    Unload FrmAyudaalm

End Sub

Private Sub GRABA_DATOS()

    Dim wcont       As Integer
    Dim wmes        As String
    Dim wadd        As Integer

    TBALMACEN.Seek "=", Txtcodalm.Text
    wmes = Format(Month(CVDate(txtfecmov.Text)), "00")
    TbStockCab.Index = "IDALMVAL"
    TbStockCab.Seek "=", Txtcodalm.Text, Trim(Txtnummov.Text)
    If TbStockCab.NoMatch Then
       TBALMACEN.Edit
       TBALMACEN.Fields("F1valing" & wmes) = Txtnummov.Text
       TBALMACEN.Update
       wadd = 1
       TbStockCab.AddNew
    Else
       wadd = 0
       TbStockCab.Edit
    End If
    TbStockCab.Fields("f4numval") = "" & Txtnummov.Text
    TbStockCab.Fields("f2codprov") = "" & Txtcodprov.Text
    TbStockCab.Fields("f4numdoc") = "" & Txtsergui.Text & "/" & Txtnumgui.Text
    TbStockCab.Fields("f4numfac") = "" & Txtserfac.Text & "/" & Txtnumfac.Text
    TbStockCab.Fields("f4fecval") = Format(txtfecmov.Text, "dd/mm/yyyy")
    TbStockCab.Fields("f2codalm") = "" & Txtcodalm.Text
    TbStockCab.Fields("f1codori") = "XC0"
    TbStockCab.Fields("f1coddoc") = Format(CmbDocum.ListIndex, "00")
    TbStockCab.Fields("f4moneda") = IIf(Cmbmoneda.ListIndex = 0, "S", "D")
    TbStockCab.Fields("f4tipcam") = "" & Txttipcam.Text
    TbStockCab.Fields("f4numimp") = Format$(Txtnumord.Text, "0000000")
    TbStockCab.Update

    i% = 0
    TbCabOrden.Index = "IDNUMORD"
    Do While i% <= gcanti
       TbCabOrden.Seek "=", Val(Format(wimporta(i%).Orden, "0000000"))
       If Not TbCabOrden.NoMatch Then
          If gfalta = "0" And wimporta(i%).f4falta = "0" Then
             TbCabOrden.Edit
             TbCabOrden.Fields("f4estval") = "1"
             TbCabOrden.Update
          End If
       End If
       i% = i% + 1
    Loop

    If wadd = 1 Or wedit = 1 Then
        'For i% = 0 To GridDetalle.Rows - 1
        If TbTempfac.RecordCount > 0 Then
            TbTempfac.MoveFirst
        End If
        Do While Not TbTempfac.EOF
            'GridDetalle.Row = i%
            TBPRODUCTO.Seek "=", Trim(TbTempfac!F5CODPRO)
            If Cmbmoneda.ListIndex = 0 Then
                wdolar# = Val(Format(TbTempfac!f6valpro, "#0.000")) / IIf(Val(Format(Txttipcam.Text, "#0.000")) = 0#, 1, Val(Format(Txttipcam.Text, "#0.000")))
                wsoles# = Val(Format(TbTempfac!f6valpro, "#0.000"))
            Else
                wdolar# = Val(Format(TbTempfac!f6valpro, "#0.000"))
                wsoles# = Val(Format(TbTempfac!f6valpro, "#0.000")) * Val(Format(Txttipcam.Text, "#0.000"))
            End If
            Vale_Detalle Trim(Txtnummov.Text), Trim(TbTempfac!F5CODPRO), Val(Format(TbTempfac!F6canmov, "#0.000")), wsoles#, Trim(Txtcodalm.Text), CVDate(txtfecmov.Text), wdolar#
            '******************************************
            '******************************************
            TbProdProv.Seek "=", Trim(TbTempfac!F5CODPRO)
            If Not TbProdProv.NoMatch Then
                TbProdProv.Edit
                TbProdProv.Fields("F5valvta") = Val(Format(TbTempfac!f6valpro, "#0.000"))
                TbProdProv.Update
            End If
        'Next i%
            TbTempfac.MoveNext
        Loop
        wedit = 0
    End If
    GridDetalle.Enabled = False: Txtcodalm.Enabled = False
    txtfecmov.Enabled = False: Cmbmoneda.Enabled = False
    'Txttipcam.Enabled = False
    Txtnummov.Enabled = False
    wgraba = 1

End Sub

Private Sub GridDetalle_Change()
    wgraba = 0
End Sub

Private Sub GridDetalle_DblClick()
    
    If GridDetalle.col = 0 Or GridDetalle.col = 1 Then
        GridDetalle_KeyDown 113, 0
    End If

End Sub

Private Sub GridDetalle_KeyDown(KeyCode As Integer, Shift As Integer)

'--------------- Tecla DOWN Para Aumentar Un Item -------------------------
    Select Case KeyCode
    Case 27: GridDetalle.col = 0: 'GridDetalle.Row = GridDetalle.Rows - 1
             If Len(Trim(GridDetalle.Columns(0))) = 0 Then
                Datadetalle.Recordset.Delete
                Datadetalle.Refresh
             End If
             Txtcodalm.SetFocus
    Case Is = 115: menu310_Click 2
    Case Is = 113: menu310_Click 0
    Case Is = 13
       Select Case GridDetalle.col
         Case 0: Valida_Producto
                 GridDetalle.col = 1: GridDetalle.SetFocus
         Case 1: GridDetalle.col = 2: GridDetalle.SetFocus
         Case 2: GridDetalle.col = 3: GridDetalle.SetFocus
         Case 3: GridDetalle.Columns(3) = Format(GridDetalle.Columns(3), "###,##0.000")
                 GridDetalle.Columns(6) = Format(Val(Format(GridDetalle.Columns(3), "#0.000")) * Val(Format(GridDetalle.Columns(5), "#0.000")), "###,##0.00")
                 GridDetalle.col = 4: GridDetalle.SetFocus
         Case 4: GridDetalle.col = 5: GridDetalle.SetFocus
         Case 5: GridDetalle.Columns(5) = Format(GridDetalle.Columns(5), "###,##0.000")
                 GridDetalle.Columns(6) = Format(Val(Format(GridDetalle.Columns(3), "#0.000")) * Val(Format(GridDetalle.Columns(5), "#0.000")), "###,##0.00")
                 'If (GridDetalle.Row < GridDetalle.Rows - 1) Then
                    GridDetalle.Row = GridDetalle.Row + 1
                 'Else
                 '   GridDetalle.Row = 0
                 'End If
                 GridDetalle.col = 3: GridDetalle.SetFocus
         Case 6: GridDetalle.Columns(5) = Format(GridDetalle.Columns(5), "###,##0.00")
                 If Val(Format(GridDetalle.Columns(2), "#0.00")) > 0# Then
                    GridDetalle.Columns(4) = Format(Val(Format(GridDetalle.Columns(5), "#0.000")) / Val(Format(GridDetalle.Columns(2), "#0.000")), "###,##0.000")
                 End If
       End Select
    End Select

End Sub

Private Sub GridDetalle_KeyPress(keyascii As Integer)

    Select Case GridDetalle.col
        Case 2, 4: keyascii = 7
        Case 3:
        If keyascii < Asc(0) Or keyascii > Asc(9) Then
            If keyascii = 8 Or keyascii = 46 Then
                Exit Sub
            Else
                keyascii = 0
            End If
        End If
        Case 5:
        If keyascii < Asc(0) Or keyascii > Asc(9) Then
            If keyascii = 8 Or keyascii = 46 Then
                Exit Sub
            Else
                keyascii = 0
            End If
        End If
      '  Case 6: keyascii = IngCar(GridDetalle, 5, 2, 10, keyascii)
    End Select

End Sub

Private Sub GridDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        PopupMenu menu300
    End If

End Sub

Private Sub LIMPIA_DATOS()
    
    'Labelmsg.Caption = ""
    GridDetalle.Enabled = True: Txtcodalm.Enabled = True
    txtfecmov.Enabled = True: Cmbmoneda.Enabled = True
    Txttipcam.Enabled = True: Txtnummov.Enabled = True
    Txtcodprov.Text = ""
    Txtnomprov.Caption = ""
    Txtnumgui.Text = ""
    Txtsergui.Text = ""
    Txtnumfac.Text = ""
    Txtserfac.Text = ""
    Txtnumord.Text = ""
    CmbDocum.ListIndex = 0
    
    DbTempFac.Execute ("Delete From TempStock")
    Datadetalle.Refresh
    Nuevo_Item
    wgraba = 1
    wvalor = 0
    wedit = 0

End Sub

Private Sub menu310_Click(Index As Integer)
    
    Select Case Index
    Case Is = 0
        Me.MousePointer = 11
        gayuda = 0
        gcodprov = Trim(Txtcodprov.Text)
        FrmHelpPrvPro.Caption = "" & Txtnomprov.Caption
        FrmHelpPrvPro.Show 1
        Agrega_Items
        Me.MousePointer = 1
    Case Is = 1
        GridDetalle.Row = GridDetalle.Row + 1
        GridDetalle.col = 0: GridDetalle.SetFocus
    Case Is = 2
        If GridDetalle.Row > 1 Then
            Eliminar_Item
        End If
    Case Is = 3
        TBPRODUCTO.Index = "idcodpro"
        TBPRODUCTO.Seek "=", GridDetalle.Columns(0)
        If Not TBPRODUCTO.NoMatch Then
            If TBPRODUCTO.Fields("f5series") = "1" Then
                gcodpro = GridDetalle.Columns(0)
                frmingseries5.Show 1
            End If
        End If
    End Select

End Sub

Private Sub menuadd_Click()
    BtnAdd_Click
End Sub

Private Sub menubox_Click()

    If menubox.Checked Then
        menubox.Checked = False
        PanelCab.Visible = False
    Else
        menubox.Checked = True
        PanelCab.Visible = True
    End If

End Sub

Private Sub menucontrol_Click()
    BtnConsul_Click
End Sub

Private Sub menuexit_Click()
    BtnExit_Click
End Sub

Private Sub menugraba_Click()
    BtnGraba_Click
End Sub

Private Sub menuprint_Click()
    BtnPrint_Click
End Sub

Private Sub Nuevo_Datos()
    
    Calcula_Numero
    LIMPIA_DATOS
    Txtcodalm.SetFocus
    wvalor = 0
    
End Sub

Private Sub Nuevo_Item()

    'Datadetalle.Recordset.AddNew
    'Datadetalle.Recordset.Fields("f5codpro") = ""
    'Datadetalle.Recordset.Fields("f5nompro") = ""
    'Datadetalle.Recordset.Fields("f6canmov") = 0#
    'Datadetalle.Recordset.Fields("f7sigmed") = ""
    'Datadetalle.Recordset.Fields("f6valpro") = 0#
    'Datadetalle.Recordset.Fields("f6total") = 0#
    'Datadetalle.Recordset.Update

End Sub

Private Sub Txtcodalm_Change()

    TBALMACEN.Seek "=", Txtcodalm.Text
    If Not TBALMACEN.NoMatch Then
        Txtnomalm.Caption = "" & TBALMACEN.Fields("F2nomalm")
        wultinv = CVDate(TBALMACEN.Fields("F1ultinv"))
    End If

End Sub

Private Sub Txtcodalm_DblClick()
    Txtcodalm_KeyUp 113, 0
End Sub

Private Sub TXTCODALM_KEYPRESS(keyascii As Integer)
    
    If keyascii = 13 Then
        Txtcodalm.Text = Format(Txtcodalm.Text, "00")
        TBALMACEN.Seek "=", Txtcodalm.Text
        If Not TBALMACEN.NoMatch Then
            Calcula_Numero
            SendKeys "{TAB}"
        Else
            Beep
            MsgBox "Error en Código de Almacén", 16, "Inventario"
            Txtcodalm.SetFocus
        End If
    End If
    
End Sub

Private Sub Txtcodalm_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodalm = "" & Txtcodalm.Text
        FrmAyudaalm.Show 1
        Txtcodalm.Text = gcodalm
        Me.MousePointer = 1
        TXTCODALM_KEYPRESS 13
    End If

End Sub

Private Sub Txtcodprov_Change()
        
    Tbproveedor.Seek "=", Txtcodprov.Text
    If Not Tbproveedor.NoMatch Then
        Txtnomprov.Caption = "" & Tbproveedor.Fields("F2nomprov")
    End If
    wgraba = 0
End Sub

Private Sub Txtcodprov_DblClick()
    Txtcodprov_KeyUp 113, 0
End Sub

Private Sub Txtcodprov_KeyPress(keyascii As Integer)
    
    If keyascii = 13 Then
        Txtcodprov.Text = Trim(Txtcodprov.Text)
        Tbproveedor.Seek "=", Txtcodprov.Text
        If Not Tbproveedor.NoMatch Then
            SendKeys "{TAB}"
        Else
            Beep
        End If
    End If
    
End Sub

Private Sub Txtcodprov_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodprov = "" & Txtcodprov.Text
        gtipmov = "I"
        FrmAyudaProv.Show 1
        Txtcodprov.Text = gcodprov
        Me.MousePointer = 1
        Txtcodprov_KeyPress 13
    End If

End Sub

Private Sub Txtfecmov_KeyPress(keyascii As Integer)

    If keyascii = 13 Then
        DataTipCam.Recordset.FindFirst "CVDATE(F6FECHA)  = CVDATE('" + txtfecmov.Text + "')"
        Calcula_Numero
        If Not DataTipCam.Recordset.NoMatch Then
            Txttipcam.Text = Val(Format(DataTipCam.Recordset.Fields("F6TIPCAM"), "#0.000"))
        Else
            Txttipcam.Text = 0#
        End If
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TxtNumFac_Change()
    wgraba = 0
End Sub

Private Sub Txtnumfac_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        Txtnumfac.Text = Format(Txtnumfac.Text, "0000000")
        SendKeys "{tab}"
    End If
End Sub

Private Sub Txtnumgui_Change()
    wgraba = 0
End Sub

Private Sub Txtnumgui_KeyPress(keyascii As Integer)
    
    If keyascii = 13 Then
        Txtnumgui.Text = Format(Txtnumgui.Text, "0000000")
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Txtnummov_Keypress(keyascii As Integer)

    If keyascii = 13 Then
        Txtnummov.Text = Format(Txtnummov.Text, "I-000000")
        TbStockCab.Index = "IDALMVAL"
        TbStockCab.Seek "=", Txtcodalm.Text, Txtnummov.Text
        If Not TbStockCab.NoMatch Then
            If Left(TbStockCab.Fields("F4NUMVAL"), 1) = "I" Then
                Actualiza_Datos
                'Txtcodprov.SetFocus
            Else
                LIMPIA_DATOS
                SendKeys "{TAB}"
            End If
        Else
            If COMPRAS <> 1 Then
                LIMPIA_DATOS
                SendKeys "{TAB}"
            Else
                SendKeys "{TAB}"
            End If
        End If
    End If

End Sub

Private Sub Txtnumord_Change()
    wgraba = 0
End Sub

Private Sub Txtnumord_DblClick()
Txtnumord_KeyUp 113, 0
End Sub

Private Sub Txtnumord_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txtnumord_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        Me.MousePointer = 11
        gcodord = "" & Format(Txtnumord.Text, "0000000")
        FrmAyudaOrdCom.Show 1
        Txtnumord.Text = Format(gcodord, "0000000")
        Me.MousePointer = 1
        'txtnumord_keypress 13
    End If

End Sub

Private Sub Txtserfac_Change()
    wgraba = 0
End Sub

Private Sub Txtserfac_KeyPress(keyascii As Integer)
    
    If keyascii = 13 Then
        Txtserfac.Text = Format(Txtserfac.Text, "000")
        SendKeys "{tab}"
    End If

End Sub

Private Sub Txtsergui_Change()
    wgraba = 0
End Sub

Private Sub Txtsergui_KeyPress(keyascii As Integer)
    
    If keyascii = 13 Then
        Txtsergui.Text = Format(Txtsergui.Text, "000")
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Txttipcam_Change()
    wgraba = 0
End Sub

Private Sub Txttipcam_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        GridDetalle.col = 0: SendKeys "{TAB}"
    End If
End Sub

Private Sub Valida_Producto()
    
    TBPRODUCTO.Seek "=", Trim(GridDetalle.Columns(0))
    If Not TBPRODUCTO.NoMatch Then
        GridDetalle.Columns(1) = "" & TBPRODUCTO.Fields("F5modelo")
        GridDetalle.Columns(2) = "" & TBPRODUCTO.Fields("F5nompro")
        GridDetalle.Columns(4) = "" & TBPRODUCTO.Fields("F7codmed")
        SendKeys "{Enter}"
    End If

End Sub
