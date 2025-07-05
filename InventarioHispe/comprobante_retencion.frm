VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form comprobante_retencion 
   ClientHeight    =   7005
   ClientLeft      =   1695
   ClientTop       =   1320
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10440
   Begin Threed.SSPanel pnlretenido 
      Height          =   330
      Left            =   9000
      TabIndex        =   22
      Top             =   5805
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "0.00"
      BackColor       =   13160660
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
      Alignment       =   4
   End
   Begin Crystal.CrystalReport crycompret 
      Left            =   5175
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2385
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1845
      Width           =   7980
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   1485
      Width           =   7980
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1275
      Left            =   7155
      TabIndex        =   14
      Top             =   90
      Width           =   3210
      _Version        =   65536
      _ExtentX        =   5662
      _ExtentY        =   2249
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
      Begin VB.TextBox documento 
         Height          =   315
         Left            =   1665
         MaxLength       =   7
         TabIndex        =   20
         Top             =   810
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   585
         MaxLength       =   3
         TabIndex        =   0
         Top             =   810
         Width           =   645
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         Height          =   210
         Left            =   1395
         TabIndex        =   19
         Top             =   855
         Width           =   180
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "COMPROBANTE  DE  RETENCION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   16
         Top             =   450
         Width           =   3090
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "R.U.C.                 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   15
         Top             =   135
         Width           =   2790
      End
   End
   Begin aBoxCtl.aBox TxtFecVen 
      Height          =   315
      Left            =   2385
      TabIndex        =   3
      Top             =   2205
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      ABoxType        =   ""
      MinValue        =   "D10000101"
      MaxValue        =   "D99991231"
      ABoxStyle       =   2
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
      FocusSelect     =   -1  'True
      ApplyTextFormat =   -1  'True
      TextFormat      =   "dd/mm/yyyy"
      Text            =   "10/12/2002"
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
      ButtonPicture   =   "comprobante_retencion.frx":0000
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
   Begin Threed.SSPanel SSPanel4 
      Height          =   285
      Left            =   630
      TabIndex        =   5
      Top             =   6255
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   503
      _StockProps     =   15
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
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   3030
      Left            =   90
      OleObjectBlob   =   "comprobante_retencion.frx":0352
      TabIndex        =   4
      Top             =   2655
      Width           =   10275
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   45
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "comprobante_retencion.frx":3868
      ToolBars        =   "comprobante_retencion.frx":8448
   End
   Begin Threed.SSPanel pnltotal 
      Height          =   330
      Left            =   7515
      TabIndex        =   23
      Top             =   5805
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "0.00"
      BackColor       =   13160660
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
      Alignment       =   4
   End
   Begin VB.Label lblanulado 
      Alignment       =   2  'Center
      Caption         =   "A N U L A D O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4410
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Nuevos Soles"
      Height          =   210
      Left            =   9360
      TabIndex        =   18
      Top             =   6300
      Width           =   1005
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "SON  :"
      Height          =   210
      Left            =   90
      TabIndex        =   17
      Top             =   6300
      Width           =   465
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   210
      Left            =   2070
      TabIndex        =   13
      Top             =   2250
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   210
      Left            =   2070
      TabIndex        =   12
      Top             =   1935
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   210
      Left            =   2070
      TabIndex        =   11
      Top             =   1530
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Emisión"
      Height          =   210
      Left            =   135
      TabIndex        =   10
      Top             =   2250
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "R.U.C."
      Height          =   210
      Left            =   135
      TabIndex        =   9
      Top             =   1845
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Señor(es)"
      Height          =   210
      Left            =   135
      TabIndex        =   8
      Top             =   1530
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   7
      Top             =   765
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Top             =   405
      Width           =   60
   End
End
Attribute VB_Name = "comprobante_retencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bancos              As ADODB.Connection
Dim temp                As ADODB.Connection
Dim sw_nuevo_doc        As Boolean
Dim sw_nuevo_item       As Boolean
Dim rsnuevo             As ADODB.Recordset
Dim RSDETALLE           As ADODB.Recordset
Dim cnumdoc             As String
Dim global1             As Double
Dim sw_cabecera         As Boolean
Dim sw_detalle          As Boolean
Dim amovs_cab(0 To 7)   As a_grabacion
Dim amovs_det(0 To 8)   As a_grabacion
Dim DBTable             As String
Dim SQL                 As String
Dim nombreruc           As String
Dim base_temp           As String
Dim CON                 As String
Dim rsregisofi      As New ADODB.Recordset

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    dxDBGrid1.Columns.FocusedIndex = 3
    
End Sub

Private Sub Form_Load()

    Set bancos = New ADODB.Connection
    Set rsnuevo = New ADODB.Recordset
    
'''    Me.Height = 8100
'''    Me.Width = 10530
'''    Me.Left = 1550
'''    Me.Top = 700
    
    global1 = 0.06
    
    TxtFecVen.Value = Format(Now, "dd/mm/yyyy")
    Me.Caption = "Comprobante de Retencion"
    With bancos
        .Provider = "Microsoft.JET.OLEDB.4.0; Data Source=" & wrutabancos & "\db_bancos.mdb; Persist Security Info=False"
        .Open
    End With
    sw_nuevo_doc = True
    sw_nuevo_item = False
    sw_detalle = False
    
    SQL = "Select * from param_com where f1codemp='" & wempresa & "'"
    If rsnuevo.State = adStateOpen Then rsnuevo.Close
    rsnuevo.Open SQL, bancos, adOpenStatic, adLockOptimistic
    If Not rsnuevo.EOF Then
        Label1.Caption = "" & rsnuevo.Fields("f1nomemp")
        Label2.Caption = "" & rsnuevo.Fields("f1diremp")
        'Label3.Caption = "-Lima"
        nombreruc = "" & rsnuevo.Fields("f1rucemp")
    End If
    Label10.Caption = "R.U.C.  " & nombreruc
    
    BASE_TEMPORAL
    TABLA_TEMPORAL
    
    DELETEREC_N DBTable, temp
    dxDBGrid1.Dataset.Refresh
    
    Conf_Grid

End Sub

Private Sub Conf_Grid()
    
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
        .Set (egoShowBands)
        'dxDBGrid1.Columns(5).SummaryFooterType = cstSum
        'dxDBGrid1.Columns(6).SummaryFooterType = cstSum
        
    End With
    Call AdicionaItem
    
End Sub

Public Sub BASE_TEMPORAL()

    Set temp = New ADODB.Connection
    base_temp = "TEMP_RET.MDB"
    CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & base_temp & "; Persist Security Info=False"
    temp.Open CON
    
End Sub

Public Sub TABLA_TEMPORAL()
    
    DBTable = "DETALLE"
    
End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean
Dim i               As Integer
 
    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_doc = False Then
        DELETEREC_N DBTable, temp
        dxDBGrid1.Dataset.Refresh
    End If
    dxDBGrid1.Dataset.ADODataset.ConnectionString = temp
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    
    With dxDBGrid1.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
        For i = 1 To 1
            If sw_nuevo_temp = True Then
                If sw_nuevo_doc = True Then
                    .Edit
                Else
                    .Append
                End If
                sw_nuevo_temp = True
            Else
                .Append
            End If
            .FieldValues("ITEM") = i
            .FieldValues("TIPO") = ""
            .FieldValues("SERIE") = ""
            .FieldValues("NUMERO_CORRELA") = ""
            .FieldValues("FECHA_EMISION") = Null
            .FieldValues("MONTO_PAGO") = Format(0, "###,##0.00")
            .FieldValues("IMPORTE_RETENIDO") = Format(0, "###,##0.00")
        Next
        .Post
        sw_nuevo_item = False
    End With
     
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
          
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    'KeyAscii = TxtNum1(KeyAscii)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtFecVen.SetFocus
    End If
    'KeyAscii = TxtNum(KeyAscii)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text3.Text = Format(Text3.Text, "000")
       Text1.SetFocus
    End If
    'KeyAscii = TxtNum(KeyAscii)

End Sub

Private Sub TxtFecVen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
        dxDBGrid1.Columns.FocusedIndex = 0
    End If
    
End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.FocusedIndex = 0
        End If
    End If

End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns(1).Value & "")) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid1.Dataset.Refresh
        End If
    End If

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            Call PROCEDIMIENTO_NUEVO
        Case "ID_Grabar"
            Me.MousePointer = 11
            dxDBGrid1.Dataset.Edit
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                dxDBGrid1.Dataset.Post
                sw_detalle = True
            End If
            If sw_cabecera = True Or sw_detalle = True Then
                grabar
                sw_detalle = False
                sw_cabecera = False
            End If
            Me.MousePointer = 1
        Case "ID_Borrar"
            Me.MousePointer = 11
            elimina Text3.Text, documento.Text
            
            Me.MousePointer = 1
        Case "ID_Imprimir"
        
            If MsgBox("Está seguro de imprimir el comprobante de retención ?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                Rem NSE IMPRIMIR_COMPRET Text3.Text, documento.Text
                IMPRIMIR_TEXTO_COMPRET Text3.Text, documento.Text
                cnn_dbbancos.Execute ("UPDATE RETENDOC SET IMPRESO='S' WHERE SERIE='" & Text3.Text & "' AND NUM_DOCUMENTO='" & documento.Text & "'")
            End If
            
        Case "ID_Modificar"
            Call PROCEDIMIENTO_NUEVO
            Lista_Comprobante.Show 1
            Unload Lista_Comprobante
            CONSULTA
            sw_nuevo_doc = False
        Case "ID_Salir"
            If dxDBGrid1.Dataset.State = dsEdit Then
                dxDBGrid1.Dataset.Post
                sw_nuevo_item = True
            End If
            If sw_cabecera = True Or sw_detalle = True Then
                If MsgBox("Desea grabar el movimiento?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                    grabar
                    sw_detalle = False
                End If
            End If
            Unload Me
    End Select
    
End Sub

Public Sub CONSULTA()
Dim tbretendoc  As ADODB.Recordset
Dim tbretenmov  As ADODB.Recordset
Dim a           As Double
Dim b           As Double
Dim c           As Double
Dim ctipo       As String

    Set tbretendoc = New ADODB.Recordset
    Set tbretenmov = New ADODB.Recordset
    
    If Len(Trim(gnummov)) And Len(gserie) > 0 Then
        SQL = "Select * from RETENDOC where SERIE='" & gserie & "' and  NUM_DOCUMENTO='" & gnummov & "'"
        If tbretendoc.State = adStateOpen Then tbretendoc.Close
        tbretendoc.Open SQL, bancos, adOpenDynamic, adLockOptimistic
        If Not tbretendoc.EOF Then
            Text3.Text = "" & tbretendoc.Fields("SERIE")
            documento.Text = "" & tbretendoc.Fields("NUM_DOCUMENTO")
            Text1.Text = "" & tbretendoc.Fields("NOMBRE")
            Text2.Text = "" & tbretendoc.Fields("RUC")
            TxtFecVen.Value = "" & tbretendoc.Fields("FECHA")
            
            If "" & tbretendoc.Fields("ANULADO") = "S" Then
                lblanulado.Visible = True
            Else
                lblanulado.Visible = False
            End If
                  
            SQL = "Select * from RETENMOV where SERIE_D='" & gserie & "' and NUM_DOCUMENTOS='" & gnummov & "'"
            If tbretenmov.State = adStateOpen Then tbretenmov.Close
            tbretenmov.Open SQL, bancos, adOpenDynamic, adLockOptimistic
            tbretenmov.MoveFirst
            dxDBGrid1.Dataset.Delete
            If Not tbretenmov.EOF Then
                Do While Not tbretenmov.EOF
                    dxDBGrid1.Dataset.Append
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Dataset.FieldValues("TIPO") = "" & tbretenmov.Fields("TIPO")
                    dxDBGrid1.Dataset.FieldValues("SERIE") = "" & tbretenmov.Fields("SERIE")
                    dxDBGrid1.Dataset.FieldValues("NUMERO_CORRELA") = "" & tbretenmov.Fields("NUMERO_CORRELA")
                    dxDBGrid1.Dataset.FieldValues("FECHA_EMISION") = "" & tbretenmov.Fields("FECHA_EMISION")
                    dxDBGrid1.Dataset.FieldValues("MONTO_PAGO") = Val("" & tbretenmov.Fields("MONTO_PAGO"))
                    dxDBGrid1.Dataset.FieldValues("IMPORTE_RETENIDO") = Val("" & tbretenmov.Fields("IMPORTE_RETENIDO"))
                    dxDBGrid1.Dataset.FieldValues("REGCOMP") = "" & tbretenmov.Fields("REGCOMP")
                    tbretenmov.MoveNext
               Loop
            End If
            dxDBGrid1.Dataset.Post
            'a = dxDBGrid1.Columns(5).SummaryFooterValue
            'b = dxDBGrid1.Columns(6).SummaryFooterValue
            'c = a + b
            'SSPanel4.Caption = CADENANUM(Val(Format(b, "#0.00")), "S", "")
            calcula_totales
        End If
    End If
    ctipo = "M"
    
End Sub

Private Sub elimina(pserie As String, pnumero As String)
Dim rscabret    As New ADODB.Recordset
Dim cimpreso    As String

    If Len(Trim("" & documento.Text)) = 0 Then
        MsgBox "El comprobante no ha sido grabado. Verifique", vbCritical, "Atencion"
        Exit Sub
    End If
    
    If lblanulado.Visible = True Then
        MsgBox "El comprobante ya ha sido anulado. Verifique.", vbInformation, "Atención"
    Else

        If rscabret.State = adStateOpen Then rscabret.Close
        rscabret.Open "SELECT IMPRESO FROM RETENDOC WHERE SERIE='" & pserie & "' AND NUM_DOCUMENTO='" & pnumero & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rscabret.EOF Then
            cimpreso = Trim("" & rscabret.Fields("IMPRESO"))
        End If
        rscabret.Close
    
        If cimpreso = "S" Then
            If MsgBox("Está seguro(a) de anular el comprobante ?", vbYesNo, "Atencion") = vbYes Then
                cnn_dbbancos.Execute ("UPDATE RETENDOC SET ANULADO='S' WHERE SERIE='" & pserie & "' AND NUM_DOCUMENTO='" & pnumero & "' ")
                lblanulado.Visible = True
            Else
                If MsgBox("Está seguro(a) de eliminar el comprobante ?", vbYesNo, "Atencion") = vbYes Then
                    cnn_dbbancos.Execute ("DELETE * FROM RETENDOC WHERE SERIE='" & pserie & "' AND NUM_DOCUMENTO='" & pnumero & "' ")
                    cnn_dbbancos.Execute ("DELETE * FROM RETENMOV WHERE SERIE_D='" & pserie & "' AND NUM_DOCUMENTOS='" & pnumero & "' ")
                
                    sw_nuevo_doc = True
                    Nuevo
                    dxDBGrid1.Dataset.Close
                    DELETEREC_N DBTable, temp
                    AdicionaItem
                End If
            End If
        Else
            If MsgBox("Está seguro(a) de eliminar el comprobante ?", vbYesNo, "Atencion") = vbYes Then
                cnn_dbbancos.Execute ("DELETE * FROM RETENDOC WHERE SERIE='" & pserie & "' AND NUM_DOCUMENTO='" & pnumero & "' ")
                cnn_dbbancos.Execute ("DELETE * FROM RETENMOV WHERE SERIE_D='" & pserie & "' AND NUM_DOCUMENTOS='" & pnumero & "' ")
            
                sw_nuevo_doc = True
                Nuevo
                dxDBGrid1.Dataset.Close
                DELETEREC_N DBTable, temp
                AdicionaItem
            
            End If
        End If
    
    End If

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    If KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
                Rem NSE Calculo
                calcula_totales
            End If
            sw_nuevo_item = False
        End If
    End If
    
    If KeyCode = 46 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
            sw_nuevo_item = False
        End If
    End If
    
End Sub

Public Sub PROCEDIMIENTO_NUEVO()
    
    Me.MousePointer = 11
    sw_nuevo_doc = False
    sw_detalle = False
    Nuevo
    AdicionaItem
    AdicionaItem
    sw_nuevo_doc = True
    Me.MousePointer = 1

End Sub

Private Sub Nuevo()
    
    Text1.Text = ""
    Text2.Text = ""
    TxtFecVen.Value = Format(Now, "dd/mm/yyyy")
    Text3.Text = ""
    documento.Text = ""
    SSPanel4.Caption = ""
    lblanulado.Visible = False
    Text3.SetFocus
    dxDBGrid1.Columns.FocusedIndex = 0

End Sub

Private Sub grabar()
    
    GRABA_ING_PROVEEDOR
    sw_nuevo_doc = False

End Sub

Private Function GENERA_NUMCOMPRO(pserie As String)
Dim tbmes1      As ADODB.Recordset
Dim WNUMERO     As String
    
    Set tbmes1 = New ADODB.Recordset
    SQL = "Select * from RETENDOC where SERIE='" & pserie & "'"
    If tbmes1.State = adStateOpen Then tbmes1.Close
    tbmes1.Open SQL, bancos, adOpenDynamic, adLockOptimistic
    If Not tbmes1.EOF Then
        tbmes1.MoveLast
        WNUMERO = "" & Val(tbmes1.Fields("NUM_DOCUMENTO")) + 1
    Else
        WNUMERO = Format(1)
    End If
    cnumdoc = Format(WNUMERO, "0000000")
    GENERA_NUMCOMPRO = cnumdoc
    
End Function

Private Sub GRABA_ING_PROVEEDOR()
Dim ccampo          As String
Dim ctipo           As String
Dim cvalores        As String
Dim nitems          As Integer
Dim ntipo           As String
Dim nfil            As Integer
Dim cregcomp        As String
Dim ccomp           As String
    
    Set RSDETALLE = New ADODB.Recordset
    If sw_nuevo_doc = True Then
        cnumdoc = GENERA_NUMCOMPRO(Text3)
        documento.Text = cnumdoc
        ctipo = "A"
    Else
        cnumdoc = documento.Text
        ctipo = "M"
    End If
    '------------------------- ASIGNA DATOS DE LA CABECERA
    amovs_cab(0).campo = "NOMBRE": amovs_cab(0).valor = Text1.Text: amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "RUC": amovs_cab(1).valor = Text2.Text: amovs_cab(1).TIPO = "T"
    amovs_cab(2).campo = "FECHA": amovs_cab(2).valor = TxtFecVen.Value: amovs_cab(2).TIPO = "F"
    amovs_cab(3).campo = "SERIE": amovs_cab(3).valor = Text3.Text: amovs_cab(3).TIPO = "T"
    amovs_cab(4).campo = "NUM_DOCUMENTO": amovs_cab(4).valor = documento.Text: amovs_cab(4).TIPO = "T"
    amovs_cab(5).campo = "BASE": amovs_cab(5).valor = Val(Format(PnlTotal.Caption, "0.00")): amovs_cab(5).TIPO = "N"
    amovs_cab(6).campo = "RETENIDO": amovs_cab(6).valor = Val(Format(pnlretenido.Caption, "0.00")): amovs_cab(6).TIPO = "N"
    amovs_cab(7).campo = "ANULADO": amovs_cab(7).valor = "N": amovs_cab(7).TIPO = "T"
    '------------------------- ASIGNA DATOS AL DETALLE
    amovs_det(0).campo = "TIPO": amovs_det(0).valor = "": amovs_det(0).TIPO = "T"
    amovs_det(1).campo = "SERIE": amovs_det(1).valor = "": amovs_det(1).TIPO = "T"
    amovs_det(2).campo = "NUMERO_CORRELA": amovs_det(2).valor = "": amovs_det(2).TIPO = "T"
    amovs_det(3).campo = "FECHA_EMISION": amovs_det(3).valor = "": amovs_det(3).TIPO = "F"
    amovs_det(4).campo = "MONTO_PAGO": amovs_det(4).valor = "": amovs_det(4).TIPO = "N"
    amovs_det(5).campo = "IMPORTE_RETENIDO": amovs_det(5).valor = "": amovs_det(5).TIPO = "N"
    amovs_det(6).campo = "SERIE_D": amovs_det(6).valor = "": amovs_det(6).TIPO = "T"
    amovs_det(7).campo = "NUM_DOCUMENTOS": amovs_det(7).valor = "": amovs_det(7).TIPO = "T"
    amovs_det(8).campo = "REGCOMP": amovs_det(8).valor = "": amovs_det(8).TIPO = "T"
    '------------------- CALCULA NUMERO DE FILAS
    nitems = 0
    If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
    SQL = "SELECT COUNT(TIPO) AS NTIPO FROM DETALLE WHERE LEN(TRIM(TIPO)) > 0 "
    RSDETALLE.Open SQL, temp, adOpenDynamic, adLockOptimistic
    
    If Not RSDETALLE.EOF Then
        ntipo = Val("" & RSDETALLE.Fields("NTIPO"))
    End If
    RSDETALLE.Close
    
    ReDim Values(8, ntipo)
        
    If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
    RSDETALLE.Open "SELECT * FROM DETALLE", temp
    
    If Not RSDETALLE.EOF Then
         nfil = 0
         RSDETALLE.MoveFirst
         Do While Not RSDETALLE.EOF
             If Len(Trim(RSDETALLE.Fields("TIPO") & "")) > 0 Then
                Values(0, nfil) = RSDETALLE.Fields("TIPO") & ""
                Values(1, nfil) = RSDETALLE.Fields("SERIE") & ""
                Values(2, nfil) = RSDETALLE.Fields("NUMERO_CORRELA") & ""
                Values(3, nfil) = RSDETALLE.Fields("FECHA_EMISION") & ""
                Values(4, nfil) = RSDETALLE.Fields("MONTO_PAGO") & ""
                Values(5, nfil) = RSDETALLE.Fields("IMPORTE_RETENIDO") & ""
                Values(6, nfil) = Text3.Text & ""
                Values(7, nfil) = documento.Text & ""
                cregcomp = ""
                If RSDETALLE.Fields("TIPO") & "" = "07" Then
                    ccomp = "SELECT F4MESMOV,F4NUMMOV FROM REGISOFI WHERE F4RUCPRV='" & Text2.Text & _
                            "' AND F4TIPDOC='" & RSDETALLE.Fields("TIPO") & "" & _
                            "' AND F4SERDOC='" & Format(RSDETALLE.Fields("SERIE") & "", "000") & _
                            "' AND F4NUMDOC='" & Format(RSDETALLE.Fields("NUMERO_CORRELA") & "", "0000000") & "'"
                        If rsregisofi.State = adStateOpen Then rsregisofi.Close
                        rsregisofi.Open ccomp, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsregisofi.EOF Then
                            cregcomp = rsregisofi.Fields("F4MESMOV") & rsregisofi.Fields("F4NUMMOV")
                        End If
                        rsregisofi.Close
                Else
                    cregcomp = RSDETALLE.Fields("REGCOMP") & ""
                End If
                Values(8, nfil) = cregcomp
                nfil = nfil + 1
             End If
             RSDETALLE.MoveNext
         Loop
     End If
     RSDETALLE.Close
     cvalores = "111111111"
    
     If ctipo = "A" Then     '--- Nuevo
         '------- GRABA CABECERA
         GRABA_REGISTRO amovs_cab(), "RETENDOC", ctipo, 7, bancos, ""
         If sw_graba_registro = True Then
             '------- GRABA DETALLE
            GRABA_REGISTRO_DET amovs_det(), "RETENMOV", ctipo, 8, bancos, "", Values(), nfil - 1, cvalores, "", ""
         End If
     Else    '--- Modificación
        '------- GRABA CABECERA
        GRABA_REGISTRO amovs_cab(), "RETENDOC", ctipo, 7, bancos, "SERIE = '" & Text3.Text & "' AND NUM_DOCUMENTO='" & documento.Text & "'"
        '------- GRABA DETALLE
        bancos.Execute ("DELETE * FROM RETENMOV WHERE SERIE_D = '" & Text3.Text & "' AND NUM_DOCUMENTOS='" & documento & "'")
        GRABA_REGISTRO_DET amovs_det(), "RETENMOV", "A", 8, bancos, "SERIE_D  = '" & Text3.Text & "' AND NUM_DOCUMENTOS='" & documento & "'", Values(), nfil - 1, cvalores, "", ""
    End If

End Sub

Private Sub Text1_Change()
    
    If Trim(Text1.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub Text2_Change()
    
    If Trim(Text2.Text) <> "" And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim a   As Double
Dim b   As Double
Dim c   As Double

    If sw_nuevo_item = False Then
        If dxDBGrid1.Dataset.State = dsEdit Then
            sw_nuevo_item = True
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = False
        End If
        calcula_totales
        If dxDBGrid1.Columns.FocusedIndex = 4 Then
            'sw_nuevo_item = True
            'dxDBGrid1.Dataset.Edit
            'dxDBGrid1.Dataset.Post
            'dxDBGrid1.Columns(5).SummaryFooterType = cstSum
            dxDBGrid1.Columns.FocusedIndex = 5
        End If
        'a = dxDBGrid1.Columns(5).SummaryFooterValue
        'b = dxDBGrid1.Columns(6).SummaryFooterValue
        'c = a + b
        'SSPanel4.Caption = CADENANUM(Val(Format(b, "#0.00")), "S", "")
        
        sw_nuevo_item = False
    End If
    
End Sub

Private Sub Calculo()
Dim a   As Double
Dim b   As Double
Dim c   As Double
    
    sw_nuevo_item = True
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Dataset.Post
    dxDBGrid1.Columns(5).SummaryFooterType = cstSum
    dxDBGrid1.Columns(6).SummaryFooterType = cstSum
    a = dxDBGrid1.Columns(5).SummaryFooterValue
    b = dxDBGrid1.Columns(6).SummaryFooterValue
    c = a + b
    SSPanel4.Caption = CADENANUM(Val(Format(b, "#0.00")), "S", "")
    sw_nuevo_item = False

End Sub

Private Sub IMPRIMIR_COMPRET(pserie As String, pnumero As String)
Dim rscabret        As New ADODB.Recordset

    'cconex_dbbancos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
    'If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    'cnn_dbbancos.Open cconex_dbbancos
    
    If rscabret.State = adStateOpen Then rscabret.Close
    rscabret.Open "SELECT * FROM RETENDOC WHERE SERIE='" & pserie & "' AND NUM_DOCUMENTO='" & pnumero & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rscabret.EOF Then
        crycompret.DataFiles(0) = wrutabancos & "\db_bancos.mdb"
        crycompret.ReportFileName = wrutatemp & "\comprobante_retencion.rpt"
        crycompret.Formulas(0) = "RAZON_SOCIAL='" & rscabret.Fields("NOMBRE") & "'"
        crycompret.Formulas(1) = "RUC='" & rscabret.Fields("RUC") & "'"
        crycompret.Formulas(2) = "FECHA_EMISION='" & rscabret.Fields("FECHA") & "'"
        crycompret.Formulas(3) = "TOTAL_LETRAS='" & CADENANUM(Format(Val(rscabret.Fields("RETENIDO") & ""), "0.00"), "S", "*") & "'"
        crycompret.Formulas(4) = "TOTAL_PAGO='" & Format(Val("" & rscabret.Fields("BASE")), "###,###,##0.00") & "'"
        crycompret.Formulas(5) = "TOTAL_RETENIDO='" & Format(Val("" & rscabret.Fields("RETENIDO")), "###,###,##0.00") & "'"
        crycompret.SelectionFormula = "{RETENMOV.SERIE_D} = '" & pserie & "' and {RETENMOV.NUM_DOCUMENTOS} = '" & pnumero & "'"
        crycompret.Action = 1
    End If
    rscabret.Close
    'cnn_dbbancos.Close
    
End Sub

Private Sub IMPRIMIR_TEXTO_COMPRET(pserie As String, pnumero As String)
Dim nfila           As Integer
Dim ncol            As Integer
Dim rscabret        As New ADODB.Recordset
Dim rsdetret        As New ADODB.Recordset

    Printer.ScaleMode = Format(6, "#0")
    Printer.FontName = "Courier New"
    Printer.FontBold = False
    Printer.FontSize = 10
    
    'cconex_dbbancos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
    'If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    'cnn_dbbancos.Open cconex_dbbancos
    
    If rscabret.State = adStateOpen Then rscabret.Close
    rscabret.Open "SELECT * FROM RETENDOC WHERE SERIE='" & pserie & "' AND NUM_DOCUMENTO='" & pnumero & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rscabret.EOF Then
        nfila = 33: ncol = 25
        ImprimeXY rscabret.Fields("NOMBRE") & "", 0, 100, nfila, ncol
        nfila = 38: ncol = 25
        ImprimeXY rscabret.Fields("RUC") & "", 0, 11, nfila, ncol
        nfila = 43: ncol = 25
        ImprimeXY Format(rscabret.Fields("FECHA") & "", "DD/MM/YYYY"), 0, 10, nfila, ncol
        
        nfila = 64
        If rsdetret.State = adStateOpen Then rsdetret.Close
        rsdetret.Open "SELECT * FROM RETENMOV WHERE SERIE_D='" & pserie & "' AND NUM_DOCUMENTOS='" & pnumero & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsdetret.EOF Then
            rsdetret.MoveFirst
            Do While Not rsdetret.EOF
                ncol = 4
                ImprimeXY rsdetret.Fields("TIPO") & "", 0, 2, nfila, ncol
                ncol = 23
                ImprimeXY rsdetret.Fields("SERIE") & "", 0, 3, nfila, ncol
                ncol = 50
                ImprimeXY rsdetret.Fields("NUMERO_CORRELA") & "", 0, 10, nfila, ncol
                ncol = 88
                ImprimeXY Format(rsdetret.Fields("FECHA_EMISION") & "", "DD/MM/YYYY"), 0, 10, nfila, ncol
                ncol = 128
                ImprimeXY Val(rsdetret.Fields("MONTO_PAGO") & ""), 2, 15, nfila, ncol
                ncol = 168
                ImprimeXY Val(rsdetret.Fields("IMPORTE_RETENIDO") & ""), 2, 15, nfila, ncol
                rsdetret.MoveNext
                nfila = nfila + 4
            Loop
        End If
        nfila = 102: ncol = 128
        ImprimeXY Val(rscabret.Fields("BASE") & ""), 2, 15, nfila, ncol
        nfila = 102: ncol = 168
        ImprimeXY Val(rscabret.Fields("RETENIDO") & ""), 2, 15, nfila, ncol
        nfila = 109: ncol = 10
        ImprimeXY CADENANUM(Format(Val(rscabret.Fields("RETENIDO") & ""), "0.00"), "S", "*"), 0, 100, nfila, ncol
        
        rsdetret.Close
        
    End If
    rscabret.Close
    'cnn_dbbancos.Close
    
    Printer.EndDoc
    
End Sub

Private Sub calcula_totales()
Dim rstempo         As New ADODB.Recordset
Dim ntotal          As Double
Dim ntotalcre       As Double

    If rstempo.State = adStateOpen Then rstempo.Close
    rstempo.Open "SELECT SUM (MONTO_PAGO) AS NTOTAL FROM DETALLE WHERE TIPO <> '07'", temp, adOpenDynamic, adLockOptimistic
    If Not rstempo.EOF Then
        ntotal = Val(rstempo.Fields("NTOTAL") & "")
    End If
    rstempo.Close
    
    If rstempo.State = adStateOpen Then rstempo.Close
    rstempo.Open "SELECT SUM (MONTO_PAGO) AS NTOTALCRE FROM DETALLE WHERE TIPO = '07'", temp, adOpenDynamic, adLockOptimistic
    If Not rstempo.EOF Then
        ntotalcre = Val(rstempo.Fields("NTOTALCRE") & "")
    End If
    rstempo.Close
    
    PnlTotal.Caption = Format(ntotal - ntotalcre, "###,###,##0.00")
    SSPanel4.Caption = CADENANUM(ntotal - ntotalcre, "S", "")
    
    
    
    If rstempo.State = adStateOpen Then rstempo.Close
    rstempo.Open "SELECT SUM (IMPORTE_RETENIDO) AS NTOTAL FROM DETALLE WHERE TIPO <> '07'", temp, adOpenDynamic, adLockOptimistic
    If Not rstempo.EOF Then
        ntotal = Val(rstempo.Fields("NTOTAL") & "")
    End If
    rstempo.Close
    
    If rstempo.State = adStateOpen Then rstempo.Close
    rstempo.Open "SELECT SUM (IMPORTE_RETENIDO) AS NTOTALCRE FROM DETALLE WHERE TIPO = '07'", temp, adOpenDynamic, adLockOptimistic
    If Not rstempo.EOF Then
        ntotalcre = Val(rstempo.Fields("NTOTALCRE") & "")
    End If
    rstempo.Close
    
    pnlretenido.Caption = Format(ntotal - ntotalcre, "###,###,##0.00")
    
    
End Sub
