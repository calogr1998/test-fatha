VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form aplica 
   Caption         =   " Aplicación de Documentos"
   ClientHeight    =   5340
   ClientLeft      =   1440
   ClientTop       =   1830
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   11550
   Begin Threed.SSPanel SSPanel6 
      Height          =   1140
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   11310
      _Version        =   65536
      _ExtentX        =   19950
      _ExtentY        =   2011
      _StockProps     =   15
      BackColor       =   -2147483644
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
      Begin VB.TextBox txttc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10350
         TabIndex        =   1
         Text            =   "0.000"
         Top             =   135
         Width           =   780
      End
      Begin VB.TextBox txtfecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   675
         Width           =   1275
      End
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9675
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   675
         Width           =   1500
      End
      Begin VB.TextBox txtdocumento 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   2310
      End
      Begin Threed.SSPanel pnlnombre 
         Height          =   330
         Left            =   1890
         TabIndex        =   6
         Top             =   135
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   582
         _StockProps     =   15
         BackColor       =   -2147483644
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
      Begin VB.TextBox txtcodigo 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   135
         Width           =   645
      End
      Begin aBoxCtl.aBox abofecha 
         Height          =   315
         Left            =   8460
         TabIndex        =   0
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         ABoxType        =   ""
         MinValue        =   "D01000101"
         MaxValue        =   "D99991231"
         ABoxStyle       =   2
         Alignment       =   1
         AlignmentVertical=   2
         HideSelection   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ApplyTextFormat =   -1  'True
         TextFormat      =   "dd/mm/yyyy"
         Text            =   "02/09/2004"
         DateFormat      =   "dd/mm/yyyy"
         FocusDateFormat =   1
         NegativeForeColor=   255
         NumberFormat    =   17
         DecimalPlaces   =   0
         HotAppearance   =   2
         CalendarTrailingForeColor=   -2147483629
         BeginProperty CalendarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowButton      =   1
         ButtonPicture   =   "aplica.frx":0000
         ButtonWidth     =   21
         UpDownWidth     =   14
         NullText        =   ""
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcDisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalcBtnHotStyle =   4
         CalcBackColor   =   -2147483643
         CalcBtnBackColor=   -2147483643
         CalcBtnDigitColor=   -2147483646
         CalcBtnFuntionColor=   8388736
         CalcDisplayFrameColor=   65535
         CalcHeaderBackColor=   -2147483646
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "T/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9990
         TabIndex        =   15
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Aplicación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7110
         TabIndex        =   14
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4815
         TabIndex        =   13
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblmoneda 
         AutoSize        =   -1  'True
         Caption         =   "US$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9270
         TabIndex        =   11
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8550
         TabIndex        =   10
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   4
         Top             =   180
         Width           =   480
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   180
      Top             =   4860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "aplica.frx":0352
      ToolBars        =   "aplica.frx":4F2E
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3495
      Left            =   90
      OleObjectBlob   =   "aplica.frx":4FEE
      TabIndex        =   2
      Top             =   1305
      Width           =   11355
   End
End
Attribute VB_Name = "aplica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnombase            As String
Dim cnomtabla           As String
Dim cconex_form         As String
Dim cnn_form            As New ADODB.Connection

Dim Values()            As Variant
Dim amovs(0 To 10)      As a_grabacion
Dim amovs_cab(0 To 0)   As a_grabacion
Dim amovs_apli(0 To 0)   As a_grabacion

Dim dfecha_aplica       As Date
Dim ntc                 As Double

Dim cmoneda             As String
Dim ncorrela            As Double
Dim nmonto              As Double
Dim dfecha              As Date
Dim ccodigo             As String
Dim cnombre             As String
Dim cdocumento          As String

Private Sub CONFIGURA_GRID()

    With dxDBGrid1.Options
        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        .Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCanNavigation)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        '.Set (egoRowSelect)
    End With
    
    dxDBGrid1.Columns(0).Visible = False
    dxDBGrid1.Columns(10).Visible = False
    dxDBGrid1.Columns(11).Visible = False
    dxDBGrid1.Columns(12).Visible = False
    
    dxDBGrid1.Columns(1).DisableEditor = True
    dxDBGrid1.Columns(2).DisableEditor = True
    dxDBGrid1.Columns(3).DisableEditor = True
    dxDBGrid1.Columns(4).DisableEditor = True
    dxDBGrid1.Columns(5).DisableEditor = True
    dxDBGrid1.Columns(6).DisableEditor = True
    dxDBGrid1.Columns(7).DisableEditor = True
    dxDBGrid1.Columns(9).DisableEditor = True
    
    dxDBGrid1.Columns(8).SummaryFooterType = cstSum
       
End Sub

Private Sub abofecha_GotFocus()

    abofecha.FocusSelect = True
    
End Sub

Private Sub abofecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txttc.SetFocus
    End If

End Sub

Private Sub abofecha_LostFocus()

    If IsDate(abofecha.Value) = True Then
        rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.Value & "')", cnn_dbbancos
        If Not rscambios.EOF Then
            txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
        Else
            txttc.Text = Format(0, "0.000")
        End If
        rscambios.Close
    End If

End Sub

Private Sub dxDBGrid1_GotFocus()

    If Val(txttc.Text) = 0# Then
        MsgBox "El tipo de cambio no puede ser cero.", vbInformation, "Atención"
        txttc.SetFocus
    End If
    
End Sub

Private Sub dxDBGrid1_OnDblClick()

    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Columns(8).Value = Format(dxDBGrid1.Columns(7).Value, "###,###,##0.00")
    dxDBGrid1.Dataset.Refresh
    If Val(Format(dxDBGrid1.Columns(8).SummaryFooterValue, "0.00")) <= Val(Format(txtimporte.Text, "0.00")) Then
    Else
        MsgBox "El monto a aplicar no puede ser mayor al documento.", vbInformation, "Atención"
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns(8).Value = Format(0, "###,###,#0.00")
    End If
    dxDBGrid1.Columns(9).Value = Format(Val(Format(dxDBGrid1.Columns(7).Value, "0.00")) - Val(Format(dxDBGrid1.Columns(8).Value, "0.00")), "###,###,##0.00")
    dxDBGrid1.Dataset.Refresh
    
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    If Val(Format(dxDBGrid1.Columns(8).Value, "###,###,##0.00")) > 0# Then
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns(8).Value = Format(dxDBGrid1.Columns(8).Value, "###,###,##0.00")
        dxDBGrid1.Dataset.Refresh
        
        If ((Val(Format(dxDBGrid1.Columns(7).Value, "0.00")) - Val(Format(dxDBGrid1.Columns(8).Value, "0.00"))) > Val(Format(dxDBGrid1.Columns(7).Value, "0.00"))) Or (Val(Format(dxDBGrid1.Columns(7).Value, "0.00")) - Val(Format(dxDBGrid1.Columns(8).Value, "0.00"))) < 0 Then
            MsgBox "El monto no puede ser mayor al saldo. Verifique.", vbCritical, "Atención"
            dxDBGrid1.Columns(8).Value = "0.00"
        Else
            If Val(Format(dxDBGrid1.Columns(8).SummaryFooterValue, "0.00")) <= Val(Format(txtimporte.Text, "0.00")) Then
            Else
                MsgBox "El monto a aplicar no puede ser mayor al documento.", vbInformation, "Atención"
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns(8).Value = Format(0, "###,###,#0.00")
            End If
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns(9).Value = Format(Val(Format(dxDBGrid1.Columns(7).Value, "0.00")) - Val(Format(dxDBGrid1.Columns(8).Value, "0.00")), "###,###,##0.00")
            dxDBGrid1.Dataset.Refresh
        End If
    Else
        If Val(Format(dxDBGrid1.Columns(8).Value, "###,###,##0.00")) = 0# Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns(9).Value = Format(Val(Format(dxDBGrid1.Columns(7).Value, "0.00")) - Val(Format(dxDBGrid1.Columns(8).Value, "0.00")), "###,###,##0.00")
        End If
    End If
    dxDBGrid1.Dataset.Refresh
    
End Sub

Private Sub Form_Load()
Dim CadSql          As String

    cnombase = wusuario & "APLICA" & Format(Time, "hh_mm_ss") & ".MDB"
    CREATEDATABASE_N wrutatemp & "\", cnombase
    cnomtabla = "DETALLE"
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    CadSql = "(ITEM TEXT(2),DOCUM TEXT(15),REFER TEXT(100),FECHA DATE," & _
             "MONEDA TEXT(1),TOTAL DOUBLE,MONTO DOUBLE,SALDO DOUBLE,CORRELA DOUBLE," & _
             "VALORANT DOUBLE,SALDOORIG DOUBLE,SALDOREST DOUBLE,TC DOUBLE)"
    CREATETABLE_N cnomtabla, CadSql, cnn_form
    DELETEREC_N cnomtabla, cnn_form
    CONFIGURA_GRID
    
    '----------------------- VARIABLES : PARAMETROS DE FACTURACION
    cmoneda = "S"
    nmonto = 1000
    dfecha = "25/01/2002"
    ccodigo = "0001"
    cnombre = "CLIENTE 1"
    cdocumento = "Fac001/0000001"
    ncorrela = 1
    '-----------------------------------------
    
    txtfecha.Text = Format(dfecha, "DD/MM/YYYY")
    txtcodigo.Text = ccodigo
    pnlnombre.Caption = cnombre
    If cmoneda = "S" Then
        lblmoneda.Caption = "S/."
    Else
        lblmoneda.Caption = "US$"
    End If
    txtdocumento.Text = cdocumento
    abofecha.Value = Format(Date, "DD/MM/YYYY")
    txtimporte.Text = Format(nmonto, "###,###,##0.00")
    
    rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & abofecha.Value & "')", cnn_dbbancos
    If Not rscambios.EOF Then
        txttc.Text = Format(Val(rscambios.Fields("CAMBIO") & ""), "0.000")
    Else
        txttc.Text = Format(0, "0.000")
    End If
    rscambios.Close
    
    LLENA_ANTICIPOS txtcodigo.Text, "C", cnn_dbbancos, cmoneda, "H"
    
End Sub

Private Sub LLENA_ANTICIPOS(pcodigo As String, ptipo As String, pconexion As ADODB.Connection, pmoneda As String, pdebhab As String)
Dim csql        As String

    dxDBGrid1.Dataset.Close
    DELETEREC_N cnomtabla, cnn_form
    
    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls
    With dxDBGrid1.Dataset
        csql = "SELECT * FROM CTA_DCTO WHERE TIPO='" & ptipo & "' AND CODIGO='" & pcodigo & "' AND SALDO >0 AND DEB_HAB='" & pdebhab & "'"
        RSCTA_DCTO.Open csql, pconexion, adOpenDynamic, adLockOptimistic
        If Not RSCTA_DCTO.EOF Then
            i = 1
            sw_nuevo_item = True
            RSCTA_DCTO.MoveFirst
            Do While Not RSCTA_DCTO.EOF
                .Append
                .FieldValues("ITEM") = i
                If Len(Trim(RSCTA_DCTO.Fields("SERDOC") & "")) > 0 Then
                    .FieldValues("DOCUM") = RSCTA_DCTO.Fields("TIPDOCU") & RSCTA_DCTO.Fields("SERDOC") & "/" & RSCTA_DCTO.Fields("DOCUM")
                Else
                    .FieldValues("DOCUM") = RSCTA_DCTO.Fields("TIPDOCU") & RSCTA_DCTO.Fields("DOCUM")
                End If
                .FieldValues("REFER") = RSCTA_DCTO.Fields("REFERENCIA") & ""
                .FieldValues("FECHA") = Format(RSCTA_DCTO.Fields("FECHA"), "DD/MM/YYYY")
                .FieldValues("MONEDA") = RSCTA_DCTO.Fields("MONEDA") & ""
                .FieldValues("TOTAL") = Format(Val(RSCTA_DCTO.Fields("TOTAL") & ""), "###,###,##0.00")
                If pmoneda = "S" Then
                    If RSCTA_DCTO.Fields("MONEDA") & "" = "S" Then
                        .FieldValues("SALDO") = Format(Val(RSCTA_DCTO.Fields("SALDO") & ""), "###,###,##0.00")
                        .FieldValues("SALDOREST") = Format(Val(RSCTA_DCTO.Fields("SALDO") & ""), "###,###,##0.00")
                    Else
                        .FieldValues("SALDO") = Format(Val(RSCTA_DCTO.Fields("SALDO") & "") * RSCTA_DCTO.Fields("TIPCAM"), "###,###,##0.00")
                        .FieldValues("SALDOREST") = Format(Val(RSCTA_DCTO.Fields("SALDO") & "") * RSCTA_DCTO.Fields("TIPCAM"), "###,###,##0.00")
                    End If
                Else
                    If RSCTA_DCTO.Fields("MONEDA") & "" = "D" Then
                        .FieldValues("SALDO") = Format(Val(RSCTA_DCTO.Fields("SALDO") & ""), "###,###,##0.00")
                        .FieldValues("SALDOREST") = Format(Val(RSCTA_DCTO.Fields("SALDO") & ""), "###,###,##0.00")
                    Else
                        .FieldValues("SALDO") = Format(Val(RSCTA_DCTO.Fields("SALDO") & "") / RSCTA_DCTO.Fields("TIPCAM"), "###,###,##0.00")
                        .FieldValues("SALDOREST") = Format(Val(RSCTA_DCTO.Fields("SALDO") & "") / RSCTA_DCTO.Fields("TIPCAM"), "###,###,##0.00")
                    End If
                End If
                .FieldValues("MONTO") = Format(0, "###,###,##0.00")
                .FieldValues("TC") = Format(Val(RSCTA_DCTO.Fields("TIPCAM") & ""), "0.000")
                .FieldValues("CORRELA") = RSCTA_DCTO.Fields("CORRELA")
                .FieldValues("VALORANT") = Format(0, "###,###,##0.00")
                .FieldValues("SALDOORIG") = Format(Val(RSCTA_DCTO.Fields("SALDO") & ""), "###,###,##0.00")
                RSCTA_DCTO.MoveNext
                i = i + 1
            Loop
            .Post
            sw_nuevo_item = False
        End If
        RSCTA_DCTO.Close
    End With
    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    cnn_form.Close

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Grabar":
            dfecha_aplica = Format(abofecha.Value, "DD/MM/YYYY")
            ntc = txttc.Text
            GRABAR_APLICACION
            SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub

Private Sub GRABAR_APLICACION()
Dim csql        As String
Dim RSDETALLE   As New ADODB.Recordset
Dim nsaldo      As Double

    csql = "SELECT * FROM " & cnomtabla & " WHERE MONTO > 0"
    RSDETALLE.Open csql, cnn_form, adOpenDynamic, adLockOptimistic
    If Not RSDETALLE.EOF Then
        RSDETALLE.MoveFirst
        Do While Not RSDETALLE.EOF
            
            amovs(0).campo = "TIPO": amovs(0).valor = "C": amovs(0).TIPO = "T"
            amovs(1).campo = "CODIGO": amovs(1).valor = ccodigo: amovs(1).TIPO = "T"
            amovs(2).campo = "CORR_COMP": amovs(2).valor = RSDETALLE.Fields("CORRELA"): amovs(2).TIPO = "N"
            amovs(3).campo = "CORR_DCTO": amovs(3).valor = ncorrela: amovs(3).TIPO = "N"
            If RSDETALLE.Fields("MONEDA") & "" = "S" Then
                amovs(4).campo = "IMPUTASO": amovs(4).valor = RSDETALLE.Fields("MONTO"): amovs(4).TIPO = "N"
                amovs(5).campo = "IMPUTADO": amovs(5).valor = 0: amovs(5).TIPO = "N"
            Else
                amovs(4).campo = "IMPUTASO": amovs(4).valor = 0: amovs(4).TIPO = "N"
                amovs(5).campo = "IMPUTADO": amovs(5).valor = RSDETALLE.Fields("MONTO"): amovs(5).TIPO = "N"
            End If
            amovs(6).campo = "TCAMBIO": amovs(6).valor = ntc: amovs(6).TIPO = "N"
            amovs(7).campo = "FCH_REPO": amovs(7).valor = dfecha_aplica: amovs(7).TIPO = "F"
            amovs(8).campo = "ANO_REPO": amovs(8).valor = Year(dfecha_aplica): amovs(8).TIPO = "T"
            amovs(9).campo = "NRO_REPO": amovs(9).valor = Day(dfecha_aplica) & Format(Month(dfecha_aplica), "00"): amovs(9).TIPO = "N"
            amovs(10).campo = "FCH_MVTO": amovs(10).valor = dfecha_aplica: amovs(10).TIPO = "F"
            
            nsaldo = 0#
            If RSDETALLE.Fields("MONEDA") & "" = "S" Then '---- moneda del anticipo
                If lblmoneda.Caption = "S/." Then         '---- moneda del documento
                    nsaldo = RSDETALLE.Fields("MONTO")
                Else
                    nsaldo = Format(RSDETALLE.Fields("MONTO") / ntc, "0.00")
                End If
            Else
                If lblmoneda.Caption = "US$" Then
                    nsaldo = RSDETALLE.Fields("MONTO")
                Else
                    nsaldo = Format(RSDETALLE.Fields("MONTO") * ntc, "0.00")
                End If
            End If
            
            nsaldodoc = 0#
            RSCTA_DCTO.Open "SELECT SALDO FROM CTA_DCTO WHERE CORRELA=" & ncorrela & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCTA_DCTO.EOF Then
                nsaldodoc = RSCTA_DCTO.Fields("SALDO")
            End If
            RSCTA_DCTO.Close
            
            nsaldoapli = 0#
            RSCTA_DCTO.Open "SELECT SALDO FROM CTA_DCTO WHERE CORRELA=" & RSDETALLE.Fields("CORRELA") & "", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RSCTA_DCTO.EOF Then
                nsaldoapli = RSCTA_DCTO.Fields("SALDO")
            End If
            RSCTA_DCTO.Close
            
            amovs_cab(0).campo = "SALDO": amovs_cab(0).valor = nsaldodoc - nsaldo: amovs_cab(0).TIPO = "N"
            amovs_apli(0).campo = "SALDO": amovs_apli(0).valor = nsaldoapli - RSDETALLE.Fields("MONTO"): amovs_apli(0).TIPO = "N"
            
            GRABA_REGISTRO amovs(), "CTA_MVTO", "A", 10, cnn_dbbancos, ""
            GRABA_REGISTRO amovs_cab(), "CTA_DCTO", "M", 0, cnn_dbbancos, "TIPO='C' AND CORRELA =" & ncorrela & ""
            GRABA_REGISTRO amovs_apli(), "CTA_DCTO", "M", 0, cnn_dbbancos, "TIPO='C' AND CORRELA =" & RSDETALLE.Fields("CORRELA") & ""
                        
            RSDETALLE.MoveNext
            
        Loop
    End If
    RSDETALLE.Close

End Sub

Private Sub txttc_GotFocus()

    txttc.SelStart = 0
    txttc.SelLength = Len(txttc.Text)

End Sub

Private Sub txttc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
    End If

End Sub

Private Sub txttc_LostFocus()

    If Val(txttc.Text) = 0# Then
        MsgBox "El tipo de cambio no puede ser cero.", vbInformation, "Atención"
        txttc.SetFocus
    End If

End Sub
