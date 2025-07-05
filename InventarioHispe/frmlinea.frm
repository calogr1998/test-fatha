VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form frmlinea 
   Caption         =   "Línea de Productos"
   ClientHeight    =   6990
   ClientLeft      =   2280
   ClientTop       =   1650
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9480
   Begin ActiveToolBars.SSActiveToolBars atbmenu 
      Left            =   495
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "frmlinea.frx":0000
      ToolBars        =   "frmlinea.frx":3F48
   End
   Begin VSFlex7Ctl.VSFlexGrid grdnivel1 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      _cx             =   15901
      _cy             =   10610
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmlinea.frx":405C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin Threed.SSPanel pnllinea 
         Height          =   2895
         Left            =   2040
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   5106
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
         Begin VB.CommandButton cmdcerrar 
            Cancel          =   -1  'True
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   235
            Left            =   5100
            TabIndex        =   8
            Top             =   145
            Width           =   255
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   5295
            _Version        =   65536
            _ExtentX        =   9340
            _ExtentY        =   873
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
            BevelOuter      =   0
            Begin ActiveToolBars.SSActiveToolBars atbmenu2 
               Left            =   120
               Top             =   0
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   131082
               ToolBarsCount   =   1
               ToolsCount      =   2
               Tools           =   "frmlinea.frx":40C2
               ToolBars        =   "frmlinea.frx":5A36
            End
         End
         Begin VB.TextBox txtdescripcion 
            Height          =   285
            Left            =   1440
            MaxLength       =   100
            TabIndex        =   5
            Top             =   2040
            Width           =   3855
         End
         Begin VB.TextBox txtcodigo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txttitulo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   120
            Width           =   5295
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            Top             =   2040
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   360
            TabIndex        =   3
            Top             =   1320
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmlinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim nuevo As Boolean

Private Sub atbmenu_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "idnuevo"
        nuevo = True
        txtcodigo.Text = Format(GeneraCodigo, "00")
        txtdescripcion.Text = ""
        txttitulo.Text = "Nueva Línea"
        pnllinea.Visible = True
        txtdescripcion.SetFocus
    Case "idmodificar"
        grdnivel1_DblClick
    Case "ideliminar"
        eliminar
    Case "idnivel"
        frmlinea2.wcod = grdnivel1.TextMatrix(grdnivel1.Row, 1)
        frmlinea2.wdes = grdnivel1.TextMatrix(grdnivel1.Row, 2)
        frmlinea2.Show vbModal
    Case "idsalir"
        pnllinea.Visible = False
        Unload Me
End Select
End Sub

Private Sub atbmenu2_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "idgrabar"
        grabar
        
    Case "idsalir"
        pnllinea.Visible = False
End Select
End Sub

Private Sub cmdcerrar_Click()
pnllinea.Visible = False
End Sub

Private Sub Form_Load()
Set rst = New ADODB.Recordset
CargarNivel
End Sub

Public Sub CargarNivel()
If rst.State = adStateOpen Then rst.Close
SQL = "select * from sf7nivel01 order by f7codcon"
rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
grdnivel1.Rows = 1
If Not rst.EOF Then
    Do While Not rst.EOF
        With grdnivel1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = "" & rst("f7codcon")
            .TextMatrix(.Rows - 1, 2) = "" & rst("f7descon")
        End With
        rst.MoveNext
    Loop
End If
grdnivel1.ColWidth(0) = 300
grdnivel1.ColWidth(1) = 1000
grdnivel1.ColWidth(2) = 5000
End Sub

Public Function GeneraCodigo()
Dim rst As New ADODB.Recordset

SQL = "select max(val(f7codcon)) from sf7nivel01"
rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    If Val("" & rst(0).Value) = 0 Then
        GeneraCodigo = "1"
    Else
        GeneraCodigo = Val(rst(0).Value) + 1
    End If
End If
End Function

Private Sub grdnivel1_DblClick()
If grdnivel1.Rows > 1 Then
    With grdnivel1
        nuevo = False
        xcod = grdnivel1.TextMatrix(grdnivel1.Row, 1)
        xdes = grdnivel1.TextMatrix(grdnivel1.Row, 2)
        txttitulo.Text = "Modificación de " & xcod & " " & Left(xdes, 35)
        txtcodigo.Text = .TextMatrix(.Row, 1)
        txtdescripcion.Text = .TextMatrix(.Row, 2)
    End With
    pnllinea.Visible = True
End If
End Sub

Public Sub grabar()
If Trim(txtdescripcion.Text) = "" Then
    MsgBox "Debe Ingresar la Descripción", vbInformation, "Sistema de Logística"
    txtdescripcion.Text = ""
    txtdescripcion.SetFocus
    Exit Sub
End If

If nuevo Then
    SQL = "insert into sf7nivel01 (f7codcon, f7descon) " _
    & "values ('" & txtcodigo.Text & "','" & txtdescripcion.Text & "')"
Else
    SQL = "update sf7nivel01 set f7descon='" & txtdescripcion.Text & "' where f7codcon='" & txtcodigo.Text & "'"
End If
cnn_dbbancos.Execute SQL
CargarNivel

pnllinea.Visible = False
End Sub

Public Sub eliminar()
If grdnivel1.Rows > 1 Then
    xcod = grdnivel1.TextMatrix(grdnivel1.Row, 1)
    xdes = grdnivel1.TextMatrix(grdnivel1.Row, 2)
    resp = MsgBox("¿Está Seguro de Eliminar la Línea " & xcod & " " & xdes & "?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Logística")
    If resp = vbYes Then
        SQL = "select * from sf7nivel02 where f7nivel01='" & xcod & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open SQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
        If Not rst.EOF Then
            MsgBox "La Línea " & xcod & " " & xdes & " Tiene Registros Asociados. Elimine Primero Estos.", vbInformation, "Sistema de Logística"
            rst.Close
            Exit Sub
        End If
        rst.Close
        
        SQL = "delete from sf7nivel01 where f7codcon='" & xcod & "'"
        cnn_dbbancos.Execute SQL
        CargarNivel
    End If
End If
End Sub

Private Sub grdnivel1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    grdnivel1_DblClick
End If
End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    grabar
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

