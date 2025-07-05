VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Imprime_RegCompras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprime Registro de Compras"
   ClientHeight    =   5085
   ClientLeft      =   5775
   ClientTop       =   4245
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Ver Reporte"
      Height          =   1335
      Left            =   60
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
      Begin VB.CheckBox ChkDetallado 
         Caption         =   "Mostrar Detallado"
         Height          =   195
         Left            =   2940
         TabIndex        =   30
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox ChkOficial 
         Caption         =   "Oficial"
         Height          =   195
         Left            =   2940
         TabIndex        =   27
         Top             =   780
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox ChkRH 
         Caption         =   "Recibo de Honorarios"
         Height          =   255
         Left            =   2940
         TabIndex        =   26
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox ChkAgrupa 
         Caption         =   "Agrupado por Centro de Costo"
         Height          =   255
         Left            =   2940
         TabIndex        =   25
         Top             =   180
         Width           =   2535
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dólares"
         Height          =   255
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar"
      Height          =   2955
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox CboCategoria 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1800
         Width           =   2355
      End
      Begin VB.TextBox txtcentro 
         BackColor       =   &H00FFFFFF&
         Height          =   280
         Left            =   360
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TxtProv 
         BackColor       =   &H00FFFFFF&
         Height          =   280
         Left            =   360
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1440
         Width           =   885
      End
      Begin VB.Frame Frame4 
         Caption         =   "Centro de Costo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   180
         TabIndex        =   21
         Top             =   2220
         Width           =   5295
         Begin VB.TextBox LblNomCentro 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1140
            TabIndex        =   22
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Proveedor"
         Enabled         =   0   'False
         Height          =   975
         Left            =   180
         TabIndex        =   19
         Top             =   1200
         Width           =   5295
         Begin VB.TextBox LblNomProv 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1140
            TabIndex        =   20
            Top             =   240
            Width           =   3975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
            Height          =   195
            Index           =   15
            Left            =   180
            OleObjectBlob   =   "Imprime_RegCompras.frx":0000
            TabIndex        =   28
            Top             =   660
            Width           =   975
         End
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "Sin Filtro"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "N° de Registro"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame FrmFecha 
         Height          =   735
         Left            =   1560
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   3915
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "Imprime_RegCompras.frx":0070
            TabIndex        =   8
            Top             =   360
            Width           =   495
         End
         Begin MSComCtl2.DTPicker DtpDesde 
            Height          =   315
            Left            =   720
            TabIndex        =   9
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   96534529
            CurrentDate     =   39954
         End
         Begin MSComCtl2.DTPicker DtpHasta 
            Height          =   315
            Left            =   2520
            TabIndex        =   10
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   96534529
            CurrentDate     =   39954
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   2
            Left            =   2040
            OleObjectBlob   =   "Imprime_RegCompras.frx":00DA
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "Fec.Emi."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1155
      End
      Begin VB.Frame FrmMov 
         Height          =   735
         Left            =   1560
         TabIndex        =   12
         Top             =   300
         Visible         =   0   'False
         Width           =   3915
         Begin VB.TextBox TxtDesde 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   14
            Text            =   "0000000"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox TxtHasta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   13
            Text            =   "0000000"
            Top             =   300
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "Imprime_RegCompras.frx":0144
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Index           =   3
            Left            =   2040
            OleObjectBlob   =   "Imprime_RegCompras.frx":01AE
            TabIndex        =   16
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3000
      TabIndex        =   1
      Top             =   4560
      Width           =   1995
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   900
      TabIndex        =   0
      Top             =   4560
      Width           =   1995
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   300
      OleObjectBlob   =   "Imprime_RegCompras.frx":0218
      Top             =   2160
   End
End
Attribute VB_Name = "Imprime_RegCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub cmdaceptar_Click()
Dim StrNomReport As String
Dim StrNomPeriodo As String

'ANIODUA
'F4BASIMP
'F4BASIMP2
'F4BASIMP3
'IGVDG
'IGVDM
'IGVDE
'NUMDETRACCION
'FECHADETRACCION
'FECDOCREF
'TIPDOCREF
'SERDOCREF
'NUMDOCREF


wcodcosto = Trim(txtcentro.Text)
wcodcliprov = Trim(TxtProv.Text)

Dim X As Object

Set X = New acr_regcompras

With X
    .DataControl1.ConnectionString = StrConexDbBancos
    '**************
    csql = "SELECT '" & worigen & "' & right(COMC.F4NUMMOV,6) as F4COMPRO, UCase(CENTROS.F3DESCRIP) AS CCOSTO, COMC.F4FECHA, COMC.F4TIPDOC, COMC.F4SERDOC, COMC.F4NUMDOC, COMC.F4NUMMOV, "
    csql = csql & "COMC.F4RUCPRV, EF2PROVEEDORES.F2NOMPROV as F4NOMPRV, COMC.F4MONEDA, COMC.F4FECVEN,'06' as TIPODOCAUX,"
    If OptDolares.value = True Then
        csql = csql & "IIf(COMC.F4MONEDA='S',COMC.F4MONTORET/COMC.F4TIPCAM,COMC.F4MONTORET) AS F4MONTORET, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='S',COMC.F4BASIMP/COMC.F4TIPCAM,COMC.F4BASIMP)*-1,IIf(COMC.F4MONEDA='S',COMC.F4BASIMP/COMC.F4TIPCAM,COMC.F4BASIMP)) AS F4BASIMP, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='S',COMC.F4MONINA/COMC.F4TIPCAM,COMC.F4MONINA)*-1,IIf(COMC.F4MONEDA='S',COMC.F4MONINA/COMC.F4TIPCAM,COMC.F4MONINA)) AS F4MONINA, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='S',COMC.F4IGV/COMC.F4TIPCAM,COMC.F4IGV)*-1,IIf(COMC.F4MONEDA='S',COMC.F4IGV/COMC.F4TIPCAM,COMC.F4IGV)) AS F4IGV, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='S',COMC.F4TOTAL/COMC.F4TIPCAM,COMC.F4TOTAL)*-1,IIf(COMC.F4MONEDA='S',COMC.F4TOTAL/COMC.F4TIPCAM,COMC.F4TOTAL)) AS TOTAL, "
        csql = csql & "IIf(COMC.F4MONEDA='S',iif(PAG_DCTO.deb_hab='D',COMC.F4TOTAL*-1,COMC.F4TOTAL),'') AS TOTALTC, "
        csql = csql & "iif(COMC.F4CODIGV ='001', COMC.F4BASIMP) as F4BASIMP1,"
        csql = csql & "iif(COMC.F4CODIGV ='002', COMC.F4BASIMP) as F4BASIMP2,"
        csql = csql & "iif(COMC.F4CODIGV ='003', COMC.F4BASIMP) as F4BASIMP3,"
        csql = csql & "iif(COMC.F4CODIGV ='001', COMC.F4IGV) AS IGVDG,"
        csql = csql & "iif(COMC.F4CODIGV ='002', COMC.F4IGV) AS IGVDM,"
        csql = csql & "iif(COMC.F4CODIGV ='003', COMC.F4IGV) AS IGVDE,"
    ElseIf OptSoles.value = True Then
        csql = csql & "IIf(COMC.F4MONEDA='S',COMC.F4MONTORET,COMC.F4MONTORET*COMC.F4TIPCAM) AS F4MONTORET, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='D',COMC.F4BASIMP*COMC.F4TIPCAM,COMC.F4BASIMP)*-1,IIf(COMC.F4MONEDA='D',COMC.F4BASIMP*COMC.F4TIPCAM,COMC.F4BASIMP)) AS F4BASIMPO, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='D',COMC.F4MONINA*COMC.F4TIPCAM,COMC.F4MONINA)*-1,IIf(COMC.F4MONEDA='D',COMC.F4MONINA*COMC.F4TIPCAM,COMC.F4MONINA)) AS F4MONINA, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='D',COMC.F4IGV*COMC.F4TIPCAM,COMC.F4IGV)*-1,IIf(COMC.F4MONEDA='D',COMC.F4IGV*COMC.F4TIPCAM,COMC.F4IGV)) AS F4IGV, "
        csql = csql & "iif(PAG_DCTO.deb_hab='D',IIf(COMC.F4MONEDA='D',COMC.F4TOTAL*COMC.F4TIPCAM,COMC.F4TOTAL)*-1,IIf(COMC.F4MONEDA='D',COMC.F4TOTAL*COMC.F4TIPCAM,COMC.F4TOTAL)) AS TOTAL, "
        csql = csql & "IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4TOTAL*-1,COMC.F4TOTAL),'') AS TOTALTC, "
        csql = csql & "iif(COMC.F4CODIGV ='001', IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4BASIMP*-1*COMC.F4TIPCAM,COMC.F4BASIMP*COMC.F4TIPCAM),iif(PAG_DCTO.deb_hab='D',COMC.F4BASIMP*-1,COMC.F4BASIMP))) as F4BASIMP1,"
        csql = csql & "iif(COMC.F4CODIGV ='002', IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4BASIMP*-1*COMC.F4TIPCAM,COMC.F4BASIMP*COMC.F4TIPCAM),iif(PAG_DCTO.deb_hab='D',COMC.F4BASIMP*-1,COMC.F4BASIMP))) as F4BASIMP2,"
        csql = csql & "iif(COMC.F4CODIGV ='003', IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4BASIMP*-1*COMC.F4TIPCAM,COMC.F4BASIMP*COMC.F4TIPCAM),iif(PAG_DCTO.deb_hab='D',COMC.F4BASIMP*-1,COMC.F4BASIMP))) as F4BASIMP3,"
        csql = csql & "iif(COMC.F4CODIGV ='001', IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4IGV*-1*COMC.F4TIPCAM,COMC.F4IGV*COMC.F4TIPCAM),iif(PAG_DCTO.deb_hab='D',COMC.F4IGV*-1,COMC.F4IGV))) AS IGVDG,"
        csql = csql & "iif(COMC.F4CODIGV ='002', IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4IGV*-1*COMC.F4TIPCAM,COMC.F4IGV*COMC.F4TIPCAM),iif(PAG_DCTO.deb_hab='D',COMC.F4IGV*-1,COMC.F4IGV))) AS IGVDM,"
        csql = csql & "iif(COMC.F4CODIGV ='003', IIf(COMC.F4MONEDA='D',iif(PAG_DCTO.deb_hab='D',COMC.F4IGV*-1*COMC.F4TIPCAM,COMC.F4IGV*COMC.F4TIPCAM),iif(PAG_DCTO.deb_hab='D',COMC.F4IGV*-1,COMC.F4IGV))) AS IGVDE,"
    End If
    csql = csql & "COMC.F4TIPCAM, "
    csql = csql & "COMC.ANIODUA, "
    csql = csql & "COMC.NUMDETRACCION,COMC.FECHADETRACCION,COMC.FECDOCREF,COMC.TIPODOCREF,COMC.SERDOCREF,COMC.NUMDOCREF"
    csql = csql & " FROM (((REGISDOC AS COMC LEFT JOIN DOCUMENTOS AS DOC ON COMC.F4TIPDOC = DOC.F2CODDOC) LEFT JOIN CENTROS ON COMC.F4OBRA = CENTROS.F3COSTO) LEFT JOIN PAG_DCTO ON COMC.F4CORRELA = PAG_DCTO.correla) LEFT JOIN EF2PROVEEDORES ON COMC.F4CODPRV = EF2PROVEEDORES.F2CODPROV "

    If OptDolares.value = True Then
        '.FldTitleMon.Text = "(Expresado en Dólares Americanos)"
        .LblTotalTC.Caption = "TOTAL (MN)"
    ElseIf OptSoles.value = True Then
        '.FldTitleMon.Text = "(Expresado en Nuevos Soles)"
        .LblTotalTC.Caption = "TOTAL (ME)"
    End If
    If Val(Lista_RegCompras.CboMes.ListIndex) = 0 Then
        csql = csql & "where COMC.F4MESMOV like '" & Lista_RegCompras.CboAnno.Text & "%' "
    Else
        csql = csql & "where COMC.F4MESMOV='" & Lista_RegCompras.CboAnno.Text & Format(Lista_RegCompras.CboMes.ListIndex, "00") & "' "
    End If
    
    If ChkRH.value = 1 Then
        .LblTipImp.Caption = "Retención"
        '.FldIgv.DataField = "F4MONTORET"
        .IGVDG.DataField = "F4MONTORET"
        '.FldZZIGV.DataField = "F4MONTORET"
        StrNomReport = "REGISTRO DE 4TA CATEGORIA"
        .Label98.Caption = "Retención"
        csql = csql & "AND COMC.F4TIPDOC = '02' "
    Else
        .LblTipImp.Caption = "I.G.V."
        '.FldIgv.DataField = "F4IGV"
        '.FldZIGv.DataField = "F4IGV"
        '.FldZZIGV.DataField = "F4IGV"
        StrNomReport = "REGISTRO DE COMPRAS"
        csql = csql & "AND COMC.F4TIPDOC <> '02' "
    End If

    If Opcion(0).value = True Then
        csql = csql & "AND COMC.F4FECHA>=#" & Format(DtpDesde.value, "mm/dd/yyyy") & "# And COMC.F4FECHA<=#" & Format(DtpHasta.value, "mm/dd/yyyy") & "# "
    ElseIf Opcion(1).value = True Then
        csql = csql & "AND COMC.F4NUMMOV between '" & TxtDesde.Text & "' and '" & TxtHasta.Text & "' "
    End If
    If Len(Trim(wcodcosto)) > 0 Then
        csql = csql & "and COMC.f4obra='" & wcodcosto & "' "
    End If
    
    If Len(Trim(wcodcliprov)) > 0 Then
        csql = csql & "and COMC.f4codprv='" & wcodcliprov & "' "
    End If
    
    If ChkOficial.value = 1 Then
        csql = csql & "and doc.F2OFICIAL=-1 "
    End If
    
    If Val(right(CboCategoria.Text, 8)) <> 0 Then
        csql = csql & "and COMC.IntCodCategoria= " & Val(right(CboCategoria.Text, 8)) & " "
        .FldCategoria.Text = "" & Trim(left(CboCategoria.Text, Len(CboCategoria) - 8))
    Else
        '.FldCategoria.Text = ""
    End If
    
    
    If ChkAgrupa.value = 1 Then
        .GroupHeader1.Visible = True
        .GroupFooter1.Visible = True
        If Opcion(0).value = True Then
            csql = csql & "ORDER BY CENTROS.F3DESCRIP, COMC.F4FECHA"
        ElseIf Opcion(1).value = True Then
            csql = csql & "ORDER BY CENTROS.F3DESCRIP, COMC.F4NUMMOV"
        ElseIf Opcion(2).value = True Then
            csql = csql & "ORDER BY CENTROS.F3DESCRIP, COMC.F4NUMMOV"
        End If
    Else
        .GroupHeader1.Visible = False
        .GroupFooter1.Visible = False
        If Opcion(0).value = True Then
            csql = csql & "ORDER BY COMC.F4FECHA"
        ElseIf Opcion(2).value = True Then
            csql = csql & "ORDER BY COMC.F4NUMMOV"
        End If
    End If
    .DataControl1.Source = csql
    .Lbl_1.Caption = StrNomReport
    StrNomPeriodo = dev_mes(Lista_RegCompras.CboMes.ListIndex) & " de " & Lista_RegCompras.CboAnno.Text
    '.fldtitulo.Text = StrNomPeriodo
    .fldPeriodo.Text = StrNomPeriodo
    .FldEmpresa.Text = wnomcia
    .fldRucEmpresa.Text = wrucempresa
    '.fldfecha.Text = Format(Date, "dd/mm/yyyy")

End With
If Not X Is Nothing Then
    Load ReporteChildFalse
    ReporteChildFalse.Caption = StrNomReport & " - " & StrNomPeriodo
    Set ReporteChildFalse.arvPreview.object = X
    
    ReporteChildFalse.Show 1
End If
Unload Me
End Sub




Private Sub CmdCancelar_Click()
wRespuesta = 0
Unload Me
End Sub

Private Sub Form_Load()

If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
cnn_dbbancos.Open StrConexDbBancos

DtpDesde.value = DateSerial(Year(Date), Month(Date) + 0, 1)
DtpHasta.value = DateSerial(Year(Date), Month(Date) + 1, 0)

CargaCategoria

End Sub

Private Sub CargaCategoria()
Dim Af As New ADOFunctions
Dim RsCat As New ADODB.Recordset
csql = "select IntCodCategoria,StrDesCategoria from Categoria order by strdescategoria"
Set RsCat = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
CboCategoria.Clear
CboCategoria.AddItem "Todos" & Space(299) & "00000000"
If RsCat.RecordCount > 0 Then
    RsCat.MoveFirst
    Do While Not RsCat.EOF
        CboCategoria.AddItem RsCat!strdescategoria & Space(299) & Format(RsCat!intCodCategoria, "00000000")
        RsCat.MoveNext
    Loop
    CboCategoria.ListIndex = 0
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
'Set cnn_DbBancos = Nothing
End Sub

Private Sub OptFecha_Click()
FrmMov.Visible = False
FrmFecha.Visible = True
End Sub

Private Sub OptNumero_Click()
FrmMov.Visible = True
FrmFecha.Visible = False
End Sub

Private Sub Opcion_Click(Index As Integer)
Select Case Index
Case 0
FrmMov.Visible = False
FrmFecha.Visible = True
Case 1
FrmMov.Visible = True
FrmFecha.Visible = False
Case 2
FrmMov.Visible = False
FrmFecha.Visible = False
End Select
End Sub

Private Sub txtdesde_GotFocus()
TxtDesde.SelStart = 0: TxtDesde.SelLength = Len(TxtDesde.Text)

End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57, 8: KeyAscii = KeyAscii
Case 13
    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub TxtDesde_LostFocus()
TxtDesde.Text = Format(TxtDesde.Text, "0000000")
End Sub

Private Sub txthasta_GotFocus()
TxtHasta.SelStart = 0: TxtHasta.SelLength = Len(TxtHasta.Text)
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57, 8: KeyAscii = KeyAscii
Case 13
    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub TxtHasta_LostFocus()
TxtHasta.Text = Format(TxtHasta.Text, "0000000")
End Sub

Private Sub txtcentro_Change()

If Len(Trim(txtcentro.Text)) = 3 Then
    wcodcosto = txtcentro.Text
    LblNomCentro.Text = ObtenerCampo("centros", "f3abrev", "f3costo", txtcentro.Text, "T", cnn_dbbancos)
ElseIf Len(Trim(txtcentro.Text)) = 0 Then
    wcodcosto = ""
    pnlcosto.Caption = ""
End If
End Sub

Private Sub txtcentro_DblClick()
txtcentro_KeyDown 113, 0
End Sub

Private Sub txtcentro_GotFocus()
txtcentro.SelStart = 0: txtcentro.SelLength = Len(txtcentro.Text)
End Sub

Private Sub txtcentro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw_ayuda = True
        wcodcosto = ""
        'Ayuda_CENTROS.SelectInto = "'999'"
        Ayuda_Centros.Show 1
        Unload Ayuda_Centros
        Set Ayuda_Centros = Nothing
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            
            txtcentro.Text = wcodcosto
            LblNomCentro.Text = wunicosto
            txtcentro_KeyPress 13
        End If
    End If
End Sub

Private Sub txtcentro_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

Private Sub TxtProv_Change()

If Len(Trim(TxtProv.Text)) = 4 Then
    wcodcosto = TxtProv.Text
    LblNomProv.Text = ObtenerCampo("ef2proveedores", "f2nomprov", "f2codprov", TxtProv.Text, "T", cnn_dbbancos)
ElseIf Len(Trim(TxtProv.Text)) = 0 Then
    wcodcosto = ""
    LblNomProv.Text = ""
End If
End Sub

Private Sub TxtProv_DblClick()
TxtProv_KeyDown 113, 0
End Sub

Private Sub TxtProv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw_ayuda = True
        wcodcliprov = ""
        Ayuda_Proveedores.Show 1
        Unload Ayuda_Proveedores
        Set Ayuda_Centros = Nothing
        sw_ayuda = False
        If Len(Trim(wcodcliprov)) > 0 Then
            
            TxtProv.Text = wcodcliprov
            LblNomProv.Text = wnomcliprov
            TxtProv_KeyPress 13
        End If
    End If

End Sub

Private Sub TxtProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
End Sub

