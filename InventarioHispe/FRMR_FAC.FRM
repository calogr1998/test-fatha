VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmr_fac 
   Appearance      =   0  'Flat
   Caption         =   "Registro de Facturas"
   ClientHeight    =   3810
   ClientLeft      =   1590
   ClientTop       =   2880
   ClientWidth     =   8925
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
   ScaleHeight     =   3810
   ScaleWidth      =   8925
   Begin Threed.SSCommand cmdgrabar 
      Height          =   375
      Left            =   2340
      TabIndex        =   8
      Top             =   1080
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Grabar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin VB.PictureBox Panel3D2 
      BackColor       =   &H80000000&
      Height          =   2130
      Left            =   45
      ScaleHeight     =   2070
      ScaleWidth      =   8775
      TabIndex        =   3
      Top             =   1575
      Width           =   8835
      Begin VB.Data datfacturas 
         Appearance      =   0  'Flat
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   900
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid grdfacturas 
         Bindings        =   "frmr_fac.frx":0000
         Height          =   1905
         Left            =   135
         OleObjectBlob   =   "frmr_fac.frx":001A
         TabIndex        =   11
         Top             =   90
         Width           =   8520
      End
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   870
      Left            =   135
      ScaleHeight     =   810
      ScaleWidth      =   8640
      TabIndex        =   2
      Top             =   90
      Width           =   8700
      Begin VB.TextBox txtfecha 
         BackColor       =   &H80000000&
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
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1140
      End
      Begin Threed.SSPanel PnlTotal 
         Height          =   285
         Left            =   6165
         TabIndex        =   7
         Top             =   270
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.00"
         BackColor       =   -2147483648
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
         Alignment       =   4
      End
      Begin VB.TextBox TxtSerDoc 
         Appearance      =   0  'Flat
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
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   0
         Top             =   225
         Width           =   420
      End
      Begin VB.TextBox TxtNumDoc 
         Appearance      =   0  'Flat
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
         Left            =   1710
         MaxLength       =   7
         TabIndex        =   1
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Total"
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
         Height          =   240
         Index           =   1
         Left            =   5670
         TabIndex        =   6
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Fecha"
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
         Height          =   240
         Index           =   0
         Left            =   3555
         TabIndex        =   5
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serie/Doc."
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
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
   End
   Begin Threed.SSCommand cmdeliminar 
      Height          =   375
      Left            =   3735
      TabIndex        =   9
      Top             =   1080
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Eliminar"
      ForeColor       =   8388608
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   375
      Left            =   5130
      TabIndex        =   10
      Top             =   1080
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      Font3D          =   3
   End
End
Attribute VB_Name = "frmr_fac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbtemp      As DAO.Database
Dim tbtemp      As DAO.Recordset
Dim xtipdoc     As String
Dim xcodprov    As String
Dim xserie      As String
Dim xdocum      As String
Dim tbdoc_fac   As DAO.Recordset

Private Sub actualiza()
Dim xdoc        As String
Dim xconsulta   As DAO.Recordset
Dim xabrev      As String
Dim xnumdoc     As String
Dim xserie2     As String
Dim xdocum2     As String
   
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", xtipdoc
    If Not TbDocumento.NoMatch Then
        xabrev = Trim(TbDocumento.Fields("f2abrev") & "")
    End If

    xdoc = "" & Format(Left(Trim(Right(Trim(frmregiscom.CmbTipDoc.Text), 5)), 2), "00")
    xserie = frmregiscom.TxtSerDoc.Text
    xnumdoc = frmregiscom.TxtNumDoc.Text

    Set xconsulta = dbcompras.OpenRecordset("select * from doc_fac where f2codprov = '" & xcodprov & "' and f2tipdoc = '" & xdoc & "' and f2serdoc = '" & xserie & "' and f2numdoc = '" & xnumdoc & "'")
    
    If xconsulta.RecordCount > 0 Then
        xconsulta.MoveFirst
        Do While Not xconsulta.EOF
            tbtemp.AddNew
            tbtemp.Fields("F2CODPROV") = xconsulta.Fields("F2CODPROV") & ""
            tbtemp.Fields("F2TIPDOC") = xconsulta.Fields("F2TIPDOC") & ""
            tbtemp.Fields("F2SERDOC") = xconsulta.Fields("F2SERDOC") & ""
            tbtemp.Fields("F2NUMDOC") = xconsulta.Fields("F2NUMDOC") & ""
            tbtemp.Fields("F2TIPFAC") = xconsulta.Fields("F2TIPFAC") & ""
            tbtemp.Fields("F2SERIEFAC") = xconsulta.Fields("F2SERIEFAC") & ""
            tbtemp.Fields("F2NUMFAC") = xconsulta.Fields("F2NUMFAC") & ""
            tbtemp.Fields("F2ABREVFAC") = xabrev

            xserie2 = Format(xconsulta.Fields("F2SERIEFAC") & "", "000")
            xdocum2 = Format(xconsulta.Fields("F2NUMFAC") & "", "0000000")
            TbCabRegis.Index = "IDPROVDOC"
            TbCabRegis.Seek "=", xcodprov, xtipdoc, xserie2, xdocum2
            If Not TbCabRegis.NoMatch Then
                tbtemp.Fields("FECHA") = TbCabRegis.Fields("f4fecha")
                tbtemp.Fields("TOTAL") = Format(TbCabRegis.Fields("f4total"), "0.00")
                tbtemp.Fields("MONEDA") = TbCabRegis.Fields("f4moneda") & ""
                tbtemp.Fields("f4refere") = Trim(TbCabRegis.Fields("f4refere") & "")
            End If
            
            tbtemp.Update
            xconsulta.MoveNext
        Loop
    End If
    
    TbCabRegis.Index = "IDMESNUM"
    
    datfacturas.Refresh

End Sub

Private Function buscar_fac()
Dim sw As Integer
Dim xserie As String
Dim xdocum As String

    sw = False
    xserie = Format(TxtSerDoc.Text, "000")
    xdocum = Format(TxtNumDoc.Text, "0000000")
    TbCabRegis.Index = "IDPROVDOC"
    TbCabRegis.Seek "=", xcodprov, xtipdoc, xserie, xdocum
    If Not TbCabRegis.NoMatch Then
        sw = True
    Else
        sw = False
    End If
    buscar_fac = sw

End Function

Private Sub CMDELIMINAR_Click()

    If Len(Trim(TxtSerDoc.Text)) > 0 And Len(Trim(TxtNumDoc.Text)) > 0 Then
        xserie = Format(TxtSerDoc.Text, "000")
        xdocum = Format(TxtNumDoc.Text, "0000000")
        tbtemp.Index = "DOC_FAC"
        tbtemp.Seek "=", xcodprov, "" & Format(Left(Trim(Right(Trim(frmregiscom.CmbTipDoc.Text), 5)), 2), "00"), frmregiscom.TxtSerDoc.Text, frmregiscom.TxtNumDoc.Text, xtipdoc, xserie, xdocum
        If Not tbtemp.NoMatch Then
            tbtemp.Delete
        Else
            MsgBox "El documento no ha sido ubicado.", 48, "Compras"
        End If
        datfacturas.Refresh
    End If

End Sub

Private Sub cmdgrabar_Click()

    grabar

End Sub

Private Sub cmdsalir_Click()

    graba_real
    Unload Me

End Sub

Private Sub Form_Load()

    Set tbdoc_fac = dbcompras.OpenRecordset("doc_fac")

    Set dbtemp = OpenDatabase(wrutatemp & "\temp_com.mdb")
    dbtemp.Execute ("delete * from doc_fac")
    Set tbtemp = dbtemp.OpenRecordset("doc_fac")

    datfacturas.DatabaseName = wrutatemp & "\temp_com.Mdb"
    datfacturas.RecordSource = "doc_fac"
    datfacturas.Refresh

    TbDocumento.MoveFirst
    Do While Not TbDocumento.EOF
        If UCase(Trim(TbDocumento.Fields("F2ABREV") & "")) = "FAC" Then
            xtipdoc = TbDocumento.Fields("F2CODDOC") & ""
            Exit Do
        End If
        TbDocumento.MoveNext
    Loop

    xcodprov = frmregiscom.TxtCodPrv.Text

    actualiza
    Nuevo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    tbdoc_fac.Close

    tbtemp.Close
    dbtemp.Close
    
    TbCabRegis.Index = "IDMESNUM"

End Sub

Private Sub graba_real()
Dim xdoc    As String
Dim xnumdoc As String

    xdoc = "" & Format(Left(Trim(Right(Trim(frmregiscom.CmbTipDoc.Text), 5)), 2), "00")
    xserie = frmregiscom.TxtSerDoc.Text
    xnumdoc = frmregiscom.TxtNumDoc.Text

    dbcompras.Execute ("delete * from doc_fac where f2codprov = '" & xcodprov & "' and f2tipdoc = '" & xdoc & "' and f2serdoc = '" & xserie & "' and f2numdoc = '" & xnumdoc & "'")
    
    If tbtemp.RecordCount > 0 Then
        tbtemp.MoveFirst
        Do While Not tbtemp.EOF
            tbdoc_fac.AddNew
            tbdoc_fac.Fields("F2CODPROV") = tbtemp.Fields("F2CODPROV") & ""
            tbdoc_fac.Fields("F2TIPDOC") = tbtemp.Fields("F2TIPDOC") & ""
            tbdoc_fac.Fields("F2SERDOC") = tbtemp.Fields("F2SERDOC") & ""
            tbdoc_fac.Fields("F2NUMDOC") = tbtemp.Fields("F2NUMDOC") & ""
            tbdoc_fac.Fields("F2TIPFAC") = tbtemp.Fields("F2TIPFAC") & ""
            tbdoc_fac.Fields("F2SERIEFAC") = tbtemp.Fields("F2SERIEFAC") & ""
            tbdoc_fac.Fields("F2NUMFAC") = tbtemp.Fields("F2NUMFAC") & ""
            tbdoc_fac.Update
            tbtemp.MoveNext
        Loop
    End If
    
End Sub

Private Sub grabar()
Dim xabrev As String
   
    TbDocumento.Index = "IDCODDOC"
    TbDocumento.Seek "=", xtipdoc
    If Not TbDocumento.NoMatch Then
        xabrev = Trim(TbDocumento.Fields("f2abrev") & "")
    End If
    
    xserie = Format(TxtSerDoc.Text, "000")
    xdocum = Format(TxtNumDoc.Text, "0000000")
    tbtemp.Index = "DOC_FAC"
    tbtemp.Seek "=", xcodprov, "" & Format(Left(Trim(Right(Trim(frmregiscom.CmbTipDoc.Text), 5)), 2), "00"), frmregiscom.TxtSerDoc.Text, frmregiscom.TxtNumDoc.Text, xtipdoc, xserie, xdocum
    If tbtemp.NoMatch Then
        tbtemp.AddNew
    Else
        tbtemp.Edit
    End If
    TbCabRegis.Index = "IDPROVDOC"
    TbCabRegis.Seek "=", xcodprov, xtipdoc, xserie, xdocum
    If Not TbCabRegis.NoMatch Then
        tbtemp.Fields("f2codprov") = xcodprov
        tbtemp.Fields("F2TIPDOC") = "" & Format(Left(Trim(Right(Trim(frmregiscom.CmbTipDoc.Text), 5)), 2), "00")
        tbtemp.Fields("F2SERDOC") = frmregiscom.TxtSerDoc.Text
        tbtemp.Fields("F2NUMDOC") = frmregiscom.TxtNumDoc.Text
        tbtemp.Fields("F2TIPFAC") = xtipdoc
        tbtemp.Fields("F2ABREVFAC") = xabrev
        tbtemp.Fields("F2SERIEFAC") = TxtSerDoc.Text
        tbtemp.Fields("F2NUMFAC") = TxtNumDoc.Text
        tbtemp.Fields("FECHA") = TxtFecha.Text
        tbtemp.Fields("TOTAL") = Val(Format(PnlTotal.Caption, "0.00"))
        tbtemp.Fields("MONEDA") = TbCabRegis.Fields("f4moneda") & ""
        tbtemp.Fields("f4refere") = Trim(TbCabRegis.Fields("f4refere") & "")
    End If
    tbtemp.Update
    datfacturas.Refresh

End Sub

Private Sub grdfacturas_DblClick()

    modifica

End Sub

Private Sub modifica()

    TxtSerDoc.Text = grdfacturas.Columns(1)
    TxtNumDoc.Text = grdfacturas.Columns(2)
    TxtFecha.Text = grdfacturas.Columns(3)
    PnlTotal.Caption = grdfacturas.Columns(4)

End Sub

Private Sub Nuevo()

    TxtSerDoc.Text = ""
    TxtNumDoc.Text = ""
    TxtFecha.Text = ""
    PnlTotal.Caption = "0.00"

End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000")
        xserie = Format(TxtSerDoc.Text, "000")
        xdocum = Format(TxtNumDoc.Text, "0000000")
        If buscar_fac() = True Then
            TbCabRegis.Index = "IDPROVDOC"
            TbCabRegis.Seek "=", xcodprov, xtipdoc, xserie, xdocum
            If Not TbCabRegis.NoMatch Then
                TxtFecha.Text = TbCabRegis.Fields("f4fecha")
                PnlTotal.Caption = Format(TbCabRegis.Fields("f4total"), "###,###,###,##0.00")
                CmdGrabar.SetFocus
            Else
                MsgBox "La factura no ha sido registrada. Verifique.", 48, "Compras"
                TxtNumDoc.SetFocus
            End If
        Else
            MsgBox "La factura no ha sido registrada. Verifique.", 48, "Compras"
            TxtNumDoc.SetFocus
        End If
    End If
    

End Sub

Private Sub TxtSerDoc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtSerDoc.Text = Format(TxtSerDoc.Text, "000")
        TxtNumDoc.SetFocus
    End If

End Sub

