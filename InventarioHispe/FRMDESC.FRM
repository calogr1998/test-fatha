VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmdesc 
   Appearance      =   0  'Flat
   Caption         =   "Consistencia de Comprobantes sin cuadrar"
   ClientHeight    =   2460
   ClientLeft      =   3390
   ClientTop       =   2415
   ClientWidth     =   4275
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   4275
   Begin Crystal.CrystalReport cryreporte 
      Left            =   225
      Top             =   2025
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   465
      Left            =   2340
      TabIndex        =   6
      Top             =   1890
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   1770
      Left            =   90
      ScaleHeight     =   1710
      ScaleWidth      =   4050
      TabIndex        =   1
      Top             =   45
      Width           =   4110
      Begin Threed.SSFrame SSFrame1 
         Height          =   870
         Left            =   180
         TabIndex        =   3
         Top             =   675
         Width           =   3750
         _Version        =   65536
         _ExtentX        =   6615
         _ExtentY        =   1535
         _StockProps     =   14
         Caption         =   "Moneda"
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
         Begin Threed.SSOption chkdolar 
            Height          =   240
            Left            =   2520
            TabIndex        =   5
            Top             =   405
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Dólares"
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
         Begin Threed.SSOption chksoles 
            Height          =   240
            Left            =   360
            TabIndex        =   4
            Top             =   405
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Soles"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin VB.TextBox txtmesdesde 
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
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   0
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   1530
         TabIndex        =   2
         Top             =   270
         Width           =   300
      End
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   465
      Left            =   855
      TabIndex        =   7
      Top             =   1890
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      Font3D          =   3
   End
End
Attribute VB_Name = "frmdesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mmes As String * 2
Dim ctipo As String * 1

Private Sub chkdolar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If

End Sub

Private Sub chksoles_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        chkdolar.SetFocus
    End If

End Sub

Private Sub cmdaceptar_Click()
Dim ctipo As String * 1

    If chksoles.Value = True And chkdolar.Value = True Then
        ctipo = "A"
    Else
        If chksoles.Value = True Then ctipo = "S"
        If chkdolar.Value = True Then ctipo = "D"
    End If
    consis_cuadra ctipo

End Sub

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub consis_cuadra(ctipo As String)
Dim dbtempo     As DAO.Database
Dim tbtempo     As DAO.Recordset
Dim tbconsis    As DAO.Recordset
Dim ccompro     As String
Dim xorigen     As String
Dim xfecha      As Variant
Dim ndebes      As Double, nhabers As Double
Dim ndebed      As Double, nhaberd As Double
   
   mmes = Format(txtmesdesde.Text, "00")
   Set dbtempo = OpenDatabase(wrutatemp & "\db_conta.mdb")
   Set tbtempo = dbtempo.OpenRecordset("contable")
   tbtempo.Index = "conta"
   dbtempo.Execute ("delete * from cuadre")
   Set tbconsis = dbtempo.OpenRecordset("cuadre")
   
   If tbtempo.RecordCount > 0 Then
      tbtempo.MoveFirst
      If Mid(tbtempo.Fields("f3proame"), 5, 2) = mmes Then
         Me.MousePointer = 11
         Do While Not tbtempo.EOF
            ndebes = 0#: nhabers = 0#: ndebed = 0#: nhaberd = 0#
            xfecha = tbtempo.Fields("f3fchopr")
            xorigen = tbtempo.Fields("f3origen") & ""
            ccompro = tbtempo.Fields("f3compro")
            Do While ccompro = tbtempo.Fields("f3compro") And Not tbtempo.EOF
               If tbtempo.Fields("f3debhab") = "D" Then
                  ndebes = tbtempo.Fields("f3importe") + ndebes
                  ndebed = tbtempo.Fields("f3imported") + ndebed
               Else
                  nhabers = tbtempo.Fields("f3importe") + nhabers
                  nhaberd = tbtempo.Fields("f3imported") + nhaberd
               End If
               tbtempo.MoveNext
               If tbtempo.EOF Then Exit Do
               If tbtempo.Fields("f3compro") <> ccompro Then Exit Do
            Loop
            If ctipo = "S" Then
               If Format(ndebes, "0.00") <> Format(nhabers, "0.00") Then
                  tbconsis.AddNew
                  tbconsis.Fields("f3compro") = ccompro
                  tbconsis.Fields("f3fecha") = xfecha
                  tbconsis.Fields("f3origen") = xorigen
                  tbconsis.Fields("f3debe") = ndebes
                  tbconsis.Fields("f3haber") = nhabers
                  tbconsis.Fields("f3empresa") = wnomcia
                  tbconsis.Fields("f3titulo") = "COMPROBANTES DE " & UCase(dev_mes(mmes)) & " DESCUADRADOS EN SOLES"
                  tbconsis.Update
               End If
            End If
            If ctipo = "D" Then
               If Format(ndebed, "0.00") <> Format(nhaberd, "0.00") Then
                  tbconsis.AddNew
                  tbconsis.Fields("f3compro") = ccompro
                  tbconsis.Fields("f3fecha") = xfecha
                  tbconsis.Fields("f3origen") = xorigen
                  tbconsis.Fields("f3debe") = ndebed
                  tbconsis.Fields("f3haber") = nhaberd
                  tbconsis.Fields("f3empresa") = wnomcia
                  tbconsis.Fields("f3titulo") = "COMPROBANTES DE " & UCase(dev_mes(mmes)) & " DESCUADRADOS EN DOLARES"
                  tbconsis.Update
               End If
            End If
            If ctipo = "A" Then
               If Format(ndebes, "0.00") <> Format(nhabers, "0.00") Or Format(ndebed, "0.00") <> Format(nhaberd, "0.00") Then
                  tbconsis.AddNew
                  tbconsis.Fields("f3compro") = ccompro
                  tbconsis.Fields("f3fecha") = xfecha
                  tbconsis.Fields("f3origen") = xorigen
                  tbconsis.Fields("f3debe") = ndebes
                  tbconsis.Fields("f3haber") = nhabers
                  tbconsis.Fields("f3debed") = ndebed
                  tbconsis.Fields("f3haberd") = nhaberd
                  tbconsis.Fields("f3empresa") = wnomcia
                  tbconsis.Fields("f3titulo") = "COMPROBANTES DE " & UCase(dev_mes(mmes)) & " DESCUADRADOS EN SOLES/DOLARES"
                  tbconsis.Update
               End If
            End If
         Loop
         Me.MousePointer = 1
         cryreporte.DataFiles(0) = wrutatemp & "\db_conta.mdb"
         cryreporte.ReportFileName = wrutatemp & "\rdescua.rpt"
         cryreporte.Action = 1
      Else
         MsgBox "El mes no corresponde con los movimientos generados.", 48, "Atención"
      End If
   Else
      MsgBox "No existen movimientos generados para la consistencia.", 48, "Atención"
   End If
   
End Sub

Private Sub Form_Load()

    txtmesdesde.Text = mes

End Sub

Private Sub txtmesdesde_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      txtmesdesde.Text = Format(txtmesdesde.Text, "00")
      If Val(txtmesdesde.Text) & "" >= 1 And Val(txtmesdesde.Text) & "" <= 12 Then
         chksoles.SetFocus
      Else
         MsgBox "Mes incorrecto. Verifique.", 48, "Atención"
         txtmesdesde.SetFocus
      End If
   End If

End Sub

Private Sub txtmesdesde_LostFocus()

   If Len(Trim(txtmesdesde.Text)) > 0 Then
      txtmesdesde.Text = Format(txtmesdesde.Text, "00")
      If Val(txtmesdesde.Text) & "" >= 1 And Val(txtmesdesde.Text) & "" <= 12 Then
         chksoles.SetFocus
      Else
         MsgBox "Mes incorrecto. Verifique.", 48, "Atención"
      End If
   End If

End Sub

