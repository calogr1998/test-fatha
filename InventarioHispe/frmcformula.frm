VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmcformula 
   Caption         =   "Consulta de Formulas"
   ClientHeight    =   4200
   ClientLeft      =   1995
   ClientTop       =   3375
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8790
   Begin Threed.SSPanel SSPanel1 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8790
      _Version        =   65536
      _ExtentX        =   15505
      _ExtentY        =   7408
      _StockProps     =   15
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.PictureBox fg 
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   3930
         Left            =   180
         ScaleHeight     =   3870
         ScaleWidth      =   8100
         TabIndex        =   1
         Top             =   90
         Width           =   8160
         Begin MSAdodcLib.Adodc ado 
            Height          =   330
            Left            =   6030
            Top             =   3015
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "frmcformula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub formatea_fg()
    
    fg.ColWidth(0) = 300
    fg.ColWidth(1) = 500
    fg.TextMatrix(0, 1) = "grupo"
    fg.ColWidth(2) = 900
    fg.TextMatrix(0, 2) = "Còdigo"
    fg.ColWidth(3) = 4050
    fg.TextMatrix(0, 3) = "Descripcion"
    fg.ColWidth(4) = 930
    fg.TextMatrix(0, 4) = "F.Base"
    fg.ColWidth(5) = 960
    fg.TextMatrix(0, 5) = "U.F.Base"

End Sub

Private Sub fg_DblClick()
    
    fg_KeyPress 13

End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
    Case 13:
        Grupo = "" & fg.TextMatrix(fg.Row, 1)
        codmprima = "" & fg.TextMatrix(fg.Row, 2)
        sw = 1
        Unload Me
    Case 27:
        'grupo = ""
        'codmprima = ""
        sw = 0
        Unload Me
    End Select

End Sub

Private Sub Form_Load()
    ado.ConnectionString = cnn_dbbancos
    'SQL = "SELECT DISTINCTROW IF4FORMULA.F4GRUPO, IF4FORMULA.F4CODPRO, IF5PLA.F5NOMPRO, IF4FORMULA.F4FBASE, IF4FORMULA.F4UFBASE " _
          & " FROM IF4FORMULA INNER JOIN IF5PLA ON (IF4FORMULA.F4CODPRO = IF5PLA.F5CODPRO) AND (IF4FORMULA.F4GRUPO = IF5PLA.F5GRUPO) " _
          & " Where ((IF4FORMULA.F4GRUPO = '" & ggrupo & "')) ORDER BY IF4FORMULA.F4CODPRO "
    
    Rem EMB SQL = "SELECT DISTINCTROW IF4FORMULA.F4GRUPO, IF4FORMULA.F4CODPRO, IF5PLA.F5NOMPRO, IF4FORMULA.F4FBASE, IF4FORMULA.F4UFBASE " _
          & " FROM IF4FORMULA INNER JOIN IF5PLA ON (IF4FORMULA.F4CODPRO = IF5PLA.F5CODPRO) AND (IF4FORMULA.F4GRUPO = IF5PLA.F5GRUPO) " _
          & " ORDER BY IF4FORMULA.F4CODPRO "
    SQL = "SELECT DISTINCTROW IF4FORMULA.F4GRUPO, IF4FORMULA.F4CODPRO, IF5PLA.F5NOMPRO, IF4FORMULA.F4FBASE, IF4FORMULA.F4UFBASE " _
          & " FROM IF4FORMULA INNER JOIN IF5PLA ON (IF4FORMULA.F4CODPRO = IF5PLA.F5CODPRO) " _
          & " ORDER BY IF4FORMULA.F4CODPRO "
    
    ado.RecordSource = SQL
    ado.Refresh
    Set fg.DataSource = ado
    formatea_fg
    sw = 0
End Sub
