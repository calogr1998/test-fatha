VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form Transferencias 
   Caption         =   "Transferencia entre Almacenes"
   ClientHeight    =   6600
   ClientLeft      =   -1755
   ClientTop       =   1515
   ClientWidth     =   10425
   LinkTopic       =   "Form2"
   ScaleHeight     =   6600
   ScaleWidth      =   10425
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   135
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "FrmTransferencias.frx":0000
      ToolBars        =   "FrmTransferencias.frx":58B6
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   6120
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10350
      _Version        =   65536
      _ExtentX        =   18256
      _ExtentY        =   10795
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
      BevelOuter      =   1
      Begin VB.TextBox txtobserva 
         Height          =   735
         Left            =   6300
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   3930
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4065
         Left            =   45
         OleObjectBlob   =   "FrmTransferencias.frx":5A14
         TabIndex        =   17
         Top             =   1980
         Width           =   10275
      End
      Begin VB.TextBox Txtcodori 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   3
         Top             =   990
         Width           =   465
      End
      Begin VB.TextBox Txtcodpar 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   2
         Top             =   630
         Width           =   465
      End
      Begin VB.TextBox Txtcodalm 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   1
         Top             =   270
         Width           =   465
      End
      Begin Threed.SSPanel PnlNomAlm 
         Height          =   285
         Left            =   2340
         TabIndex        =   6
         Top             =   270
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
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
      End
      Begin Threed.SSPanel PnlNomOri 
         Height          =   285
         Left            =   2340
         TabIndex        =   7
         Top             =   990
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
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
      End
      Begin Threed.SSPanel PnlNomPar 
         Height          =   285
         Left            =   2340
         TabIndex        =   8
         Top             =   630
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
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
      End
      Begin Threed.SSPanel Txtnumvalo 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1395
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483635
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
      End
      Begin Threed.SSPanel Txtnumvald 
         Height          =   285
         Left            =   4770
         TabIndex        =   10
         Top             =   1395
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483635
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
      End
      Begin aBoxCtl.aBox TxtFecMov 
         Height          =   315
         Left            =   7110
         TabIndex        =   4
         Top             =   225
         Width           =   1380
         _ExtentX        =   2434
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
         Text            =   "16/10/2006"
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
         ButtonPicture   =   "FrmTransferencias.frx":9BA6
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   6435
         TabIndex        =   18
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   16
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vale Ingreso:"
         DataField       =   "<"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   15
         Top             =   1410
         Width           =   1140
      End
      Begin VB.Label LabPar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén Destino:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   14
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén Origen:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   13
         Top             =   285
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   6420
         TabIndex        =   12
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vale Salida:"
         DataField       =   "<"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   11
         Top             =   1365
         Width           =   1035
      End
   End
End
Attribute VB_Name = "Transferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sw_nuevo_doc    As Boolean, sw_ayuda As Boolean, sw_ayudaO  As Boolean
Dim sw_nuevo_item   As Boolean, sw_detalle  As Boolean, sw_cabecera As Boolean
Dim Temp            As ADODB.Connection
Dim amovs_vale(0 To 0)  As a_grabacion
Dim amovs_cab(0 To 14)  As a_grabacion
Dim amovs_det(0 To 8)   As a_grabacion
Dim sw_ayuda_prod       As Boolean

Private Sub Actualiza_Datos()
    
    Set RsStockCab = New ADODB.Recordset
    Set RsStockDet = New ADODB.Recordset
    
    SQL = "SELECT * FROM IF4VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "' AND LEFT(F4NUMVAL,1)='S' AND LEFT(F4NUMDOC,1)='I'"
    If RsStockCab.State = adStateOpen Then RsStockCab.Close
    RsStockCab.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockCab.EOF Then
        RsStockCab.MoveFirst
        Txtnumvalo.Caption = "" & RsStockCab.Fields("f4numval")
        TxtFecMov.Value = Format(RsStockCab.Fields("F4Fecval"), "dd/mm/yyyy")
        Txtcodpar.Text = "" & RsStockCab.Fields("F2CODPAR")
        If VALIDA_ALMACENO(Txtcodpar.Text) = True Then
            WcodPar = Trim(Txtcodpar.Text)
            PnlNomPar.Caption = WNomPar
        End If
        Txtcodalm.Text = "" & RsStockCab.Fields("F2CODALM")
        If VALIDA_ALMACEN(Txtcodalm.Text) = True Then
            wcod_alm = Txtcodalm.Text
            PnlNomAlm.Caption = wnomalmacen
        End If
        Txtnumvald.Caption = "" & RsStockCab.Fields("F4NUMDOC")
        Txtcodori.Text = "" & RsStockCab.Fields("f1CODORI")
        If VALIDA_ORIGEN(Txtcodori.Text) = True Then
            wcodori = Txtcodori.Text
            PnlNomOri.Caption = wnomori
        End If
        Agrega_Items
    End If

End Sub

Private Function Calcula_Numero(pcodalm As String, ptipo As String)
Dim WCONT   As String
Dim WNUMERO As String
    
    Set RsAlmacenes = New ADODB.Recordset
    SQL = "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & pcodalm & "'"
    If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
    RsAlmacenes.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsAlmacenes.EOF Then
        wmes = Format(Month(CVDate(TxtFecMov.Text)), "00")
        If ptipo = "S" Then
            WCONT = Format(Val(Mid(RsAlmacenes.Fields("F1valsal" & Format(wmes, "00")), 5, 4)) + 1, "0000")
            WNUMERO = Mid(RsAlmacenes.Fields("F1VALSAL" & Format(wmes, "00")), 1, 4) & WCONT
        Else
            WCONT = Format(Val(Mid(RsAlmacenes.Fields("F1valing" & Format(wmes, "00")), 5, 4)) + 1, "0000")
            WNUMERO = Mid(RsAlmacenes.Fields("F1VALING" & Format(wmes, "00")), 1, 4) & WCONT
        End If
    Else
        Beep
        WNUMERO = ""
    End If
    Calcula_Numero = WNUMERO
    
End Function

Private Sub Elimina_Movimientos()
Dim csql As String, CSQL1 As String

    Set RsMovAlmacen = New ADODB.Recordset
    Set RsStockCab = New ADODB.Recordset
    Set RsStockDet = New ADODB.Recordset
    
    If MsgBox("Está seguro de eliminar los movimientos registrados", vbYesNo + vbInformation, "Atención") = vbYes Then
        ''''''''''''''''''''''''PARA INGRESOS '''''''''''''''''''''''''''''''''''
        SQL = "SELECT * FROM IF4VALES WHERE F4NUMVAL ='" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'"
        If RsStockCab.State = adStateOpen Then RsStockCab.Close
        RsStockCab.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsStockCab.EOF Then
            ''''''''''''''''ELIMINAR CABECERA '''''''''''''''''
            csql = "DELETE FROM IF4VALES WHERE F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'"
            cnn_dbbancos.Execute csql
            If RsStockDet.State = adStateOpen Then RsStockDet.Close
            SQL = "SELECT * FROM IF3VALES WHERE F4NUMVAL ='" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'"
            RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsStockDet.EOF Then
                RsStockDet.MoveFirst
                Do While Not RsStockDet.EOF
                    Actualizar_Almacenes Trim(RsStockDet.Fields("F2codalm")), Trim(RsStockDet.Fields("F5codpro")), Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), CVDate(RsStockDet.Fields("F4fecval")), Val(Format(RsStockDet.Fields("F3totite"), "#0.000")), Val(Format(RsStockDet.Fields("F3totdol"), "#0.000")), "I", Val(Format(RsStockDet.Fields("F3valdol"), "#0.000"))
                    RsStockDet.MoveNext
                    If RsStockDet.EOF Then Exit Do
                Loop
                csql = "DELETE FROM IF3SERIES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
                cnn_dbbancos.Execute (csql)
                ''''''''''''''''''''ELIMINAR DETALLE ''''''''''''''''''''
                csql = "DELETE FROM IF3VALES WHERE F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'"
                cnn_dbbancos.Execute (csql)
            End If
        End If
        '''''''''''''''''''''''''ELIMINAR SALIDAS ''''''''''''''''''''''
        SQL = "SELECT * FROM IF4VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        If RsStockCab.State = adStateOpen Then RsStockCab.Close
        RsStockCab.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsStockCab.EOF Then
            CSQL1 = "DELETE FROM IF4VALES WHERE F4NUMVAL = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'"
            cnn_dbbancos.Execute CSQL1
            If RsStockDet.State = adStateOpen Then RsStockDet.Close
            SQL = "SELECT * FROM IF3VALES WHERE F4NUMVAL ='" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'"
            RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsStockDet.EOF Then
                RsStockDet.MoveFirst
                Do While Not RsStockDet.EOF
                    Actualizar_Almacenes Trim(RsStockDet.Fields("F2codalm")), Trim(RsStockDet.Fields("F5codpro")), Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), CVDate(RsStockDet.Fields("F4fecval")), Val(Format(RsStockDet.Fields("F3totite"), "#0.000")), Val(Format(RsStockDet.Fields("F3totdol"), "#0.000")), "S", Val(Format(RsStockDet.Fields("F3valdol"), "#0.000"))
                    RsStockDet.MoveNext
                    If RsStockDet.EOF Then Exit Do
                Loop
                CSQL1 = "DELETE FROM IF3SERIES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
                cnn_dbbancos.Execute (CSQL1)
                ''''''''''''''''''''ELIMINAR DETALLE ''''''''''''''''''''
                CSQL1 = "DELETE FROM IF3VALES WHERE F4NUMVAL = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'"
                cnn_dbbancos.Execute (CSQL1)
            End If
        End If
        Nuevo
    End If

End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
                dxDBGrid1.Dataset.Edit
                sw_detalle = True
                dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
        End If
    End If
'    If sw_nuevo_item = False Then
'        If Action = daInsert Then
''            dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
''            dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = Format(0)
''            dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value = Format(0)
''            dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = Format(0, "0.00")
''            dxDBGrid1.Columns.FocusedIndex = 1
''            dxDBGrid1.Dataset.Edit
'                sw_detalle = True
'                dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
'        End If
'    End If
End Sub
Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
'                If Len(Trim(dxDBGrid1.Columns(1).Value & "")) = 0 Then
'                    Allow = False
'                End If
                If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Value & "")) = 0 Then
                    Allow = False
                Else
                    dxDBGrid1.Columns.FocusedIndex = 1
                End If
            End If
        End If
        If Action = daDelete Then
            sw_detalle = True
            dxDBGrid1.Dataset.Refresh
        End If
    End If
    
End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
If dxDBGrid1.Columns.FocusedColumn.ObjectName = "Codigo" Then
    'Me.MousePointer = 11
    sw_ayuda_prod = True
    wcod_alm = Txtcodalm.Text
     If Trim(wcod_alm) = "" Then
        MsgBox "Ingrese Almacen de origen", vbInformation, "Sistema de Logistica"
        Exit Sub
    End If
    'hlp_productos.Show vbModal
    Con_Ayu = 4
    ayuda_productos.Show 1
    
    If Len(Trim(wcodproducto)) <> 0 Then
        sw_detalle = True
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
        dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
        dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = wmedida
        dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
        dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
        dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = wstockact
        dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = wprecos
        dxDBGrid1.Dataset.Post
   ' Me.MousePointer = 1
    dxDBGrid1.Columns.FocusedIndex = 6
    End If
End If
If dxDBGrid1.Columns.FocusedColumn.ObjectName = "Eliminar" Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_detalle = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
        End If
End If
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim gcod        As String
Dim ccad        As String

    Set rsconsulta = New ADODB.Recordset
    If dxDBGrid1.Dataset.State <> 0 And dxDBGrid1.Dataset.State <> 1 Then
        If dxDBGrid1.Columns.FocusedColumn.FieldName = "F5CODPRO" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "F5CODFAB" Then
            If Len(Trim(dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value)) > 0 Then
                gcod = dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value
                ccad = "(IF5PLA.F5CODPRO='" & gcod & "')"
            Else
                gcod = dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value
                ccad = "(IF5PLA.F5CODFAB='" & gcod & "')"
            End If
            If Len(Trim(gcod)) > 0 Then
                sw_detalle = True
                dxDBGrid1.Dataset.Edit
                If rsconsulta.State = adStateOpen Then rsconsulta.Close                                                                                                                                                                                                                                            '--- and por or ---------'
                Rem NSE SQL = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5MARCA, IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO, IF6ALMA.F6STOCKACT, IF5PLA.F5VALVTA, IF5PLA.F7CODMED FROM IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO WHERE (((IF5PLA.F5CODPRO)='" & gcod & "')) OR (((IF5PLA.F5CODFAB)='" & gcod & "') AND ((IF6ALMA.F2CODALM)='" & wcod_alm & "')) ORDER BY IF5PLA.F5CODPRO;"
                SQL = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5MARCA, IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO, IF6ALMA.F6STOCKACT, IF5PLA.F5VALVTA, IF5PLA.F7CODMED " & _
                      "FROM IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO " & _
                      "WHERE " & ccad & " AND ((IF6ALMA.F2CODALM)='" & wcod_alm & "') ORDER BY IF5PLA.F5CODPRO;"
                rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsconsulta.EOF Then
                    gcod = "" & Trim(rsconsulta.Fields("F5CODPRO"))
                    gfab = "" & rsconsulta.Fields("F5CODFAB")
                    wmarca = "" & rsconsulta.Fields("F5MARCA")
                    If rsmarcas.State = adStateOpen Then rsmarcas.Close
                    rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsmarcas.EOF Then
                        wmarca = rsmarcas.Fields("F2DESMAR")
                    Else
                        wmarca = "" & rsconsulta.Fields("F5marca")
                    End If
                    gnom = "" & rsconsulta.Fields("F5NOMPRO")
                    guni = "" & rsconsulta.Fields("F7CODMED")
                    gmoneda = "" & rsconsulta.Fields("F5MONEDA")
                    gstock = Val("" & rsconsulta.Fields("F6STOCKACT"))
                    gvalvta = Val("" & rsconsulta.Fields("F5VALVTA"))
                Else
                    MsgBox "Código no existe", vbInformation + vbDefaultButton1, "Atención"
                    gcodpro = "": gnompro = "": guni = "": gfab = ""
                    gvalvta = 0#: gstock = 0#: gmoneda = ""
                    dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = gcod
                    dxDBGrid1.Dataset.Post
                    Exit Sub
                End If
                dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = gcod
                dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = gnom
                dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = guni
                dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = gfab
                dxDBGrid1.Columns.ColumnByFieldName("marca").Value = wmarca
                dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = gstock
                dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = gvalvta
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = 5 'dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").ColIndex - 1
            End If
        Else
            If dxDBGrid1.Columns.FocusedColumn.FieldName = "F6CANMOV" Then
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = 7 'dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").ColIndex - 1
            End If
            If dxDBGrid1.Columns.FocusedColumn.FieldName = "F6TOTAL" Then
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = 6 'dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").ColIndex - 1
            End If
        End If
    End If
    'SE AGREGAR
    If dxDBGrid1.Columns.FocusedColumn.FieldName = "F6CANMOV" Then
            istock = Val("" & dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value)
            If istock > 0 Then
                'Verifica si hay Existencias del producto en Almacén
                If Len(Trim(wf1evalua_stock)) = 0 Then
                    xcodproducto = "" & dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value
                    nstock = CalculaExistencia(Txtcodalm.Text, xcodproducto, TxtFecMov.Value)
                    If nstock < istock Then
                        MsgBox "La Cantidad por Salir " & istock & " Es mayor a la Cantidad en Stock " & nstock, vbInformation, "Sistema de Logistica"
                        dxDBGrid1.Dataset.Edit
                        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value = ""
                        dxDBGrid1.Dataset.Post
                        dxDBGrid1.Columns.FocusedIndex = 5
                    End If
                End If
            End If
        
    End If


End Sub

Private Sub dxDBGrid1_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)

    Select Case KeyCode
        Case 123:
            wsw_codbarra = wsw_codbarra + 1
            If wsw_codbarra = 1 Then
                gcod = "": gnom = "": guni = "": gfab = ""
                gstock = 0#: gmoneda = "": gvalvta = 0#
                wcodigo_barra = ""
                lee_codigosbarra.Show 1
                If Len(Trim(wcodigo_barra)) > 0 Then
                    gcod = wcodigo_barra
                    sw_nuevo_item = True
                    dxDBGrid1.Dataset.Edit
                    If rsconsulta.State = adStateOpen Then rsconsulta.Close
                    SQL = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5marca,IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO, IF6ALMA.F6STOCKACT, IF5PLA.F5VALVTA, IF5PLA.F7CODMED FROM IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO WHERE  (IF6ALMA.F2CODALM='" & wcod_alm & "') AND ((IF5PLA.F5CODPRO='" & gcod & "') OR (IF5PLA.F5CODFAB='" & gcod & "')) ORDER BY IF5PLA.F5CODPRO;"  'Giannina
                    rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsconsulta.EOF Then
                        gcod = "" & Trim(rsconsulta.Fields("F5CODPRO"))
                        gfab = "" & rsconsulta.Fields("F5CODFAB")
                        If rsmarcas.State = adStateOpen Then rsmarcas.Close
                        rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsmarcas.EOF Then
                            wmarca = rsmarcas.Fields("F2DESMAR")
                        Else
                            wmarca = "" & rsconsulta.Fields("F5marca")
                        End If
                        Rem EMB wmarca = "" & rsconsulta.Fields("F5marca")
                        gnom = "" & rsconsulta.Fields("F5NOMPRO")
                        guni = "" & rsconsulta.Fields("F7CODMED")
                        gmoneda = "" & rsconsulta.Fields("F5MONEDA")
                        gstock = Val("" & rsconsulta.Fields("F6STOCKACT"))
                        gvalvta = Val("" & rsconsulta.Fields("F5VALVTA"))
                    Else
                        MsgBox "Código no existe", vbInformation + vbDefaultButton1, "Atención"
                        gcodpro = "": gnompro = "": guni = "": gfab = ""
                        gvalvta = 0#: gstock = 0#: gmoneda = ""
                        dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = gcodpro
                        dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = gnompro
                        dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = guni
                        dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = gfab
                        If rsmarcas.State = adStateOpen Then rsmarcas.Close
                        rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & wmarca & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsmarcas.EOF Then
                            dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = rsmarcas.Fields("F2DESMAR")
                        Else
                            dxDBGrid1.Columns.ColumnByFieldName("f5marca").Value = wmarca
                        End If
                        Rem EMB dxDBGrid1.Columns.ColumnByFieldName("F5marca").Value = wmarca
                        dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = gstock
                        dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = gvalvta
                        dxDBGrid1.Dataset.Post
                        Exit Sub
                    End If
                    dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = gcod
                    dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = gnom
                    dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = guni
                    dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = gfab
                    dxDBGrid1.Columns.ColumnByFieldName("F5marca").Value = wmarca
                    dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = gstock
                    dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = gvalvta
                    dxDBGrid1.Dataset.Post
                    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").ColIndex - 1
                    '----------------------------------------------------------------
                    rsconsulta.Close
                    
                    sw_nuevo_item = False
                    wcodigo_barra = ""
                End If
                Unload lee_codigosbarra
            Else
                wsw_codbarra = 0
            End If
    End Select

End Sub

Private Sub Form_Load()

    Me.Top = 1100
    Me.Left = 1500
    sw_nuevo_doc = True
    sw_detalle = False
    sw_nuevo_item = False
    sw_ayuda_prod = False
    BASE_TEMPORAL "TEMPFAC.MDB"
    DELETEREC_N "TmpStock", Temp
    dxDBGrid1.Dataset.Refresh
    Conf_Grid
    TxtFecMov.Value = Format(Date, "DD/MM/YYYY")

    xvale = 3
    
End Sub

Private Sub Graba_Datos()

    If rscambios.State = adStateOpen Then rscambios.Close
    rscambios.Open "SELECT * FROM CAMBIOS WHERE CVDATE(FECHA)=CVDATE('" & TxtFecMov.Value & "')", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rscambios.EOF Then
        wtipcam = Val(rscambios.Fields("CAMBIO") & "")
    Else
        wtipcam = 3.5
    End If
    rscambios.Close
    
    ''''''''''''''''' GRABANDO LA SALIDA DE ALMACEN ''''''''''''''''''
    wmes = Format(Month(CVDate(TxtFecMov.Value)), "00")
    Txtnumvalo.Caption = Calcula_Numero(Trim(Txtcodalm.Text), "S")
    ACTUALIZAR
    GRABAR_SALIDA

    ''''''''''''''''GRABANDO INGRESO AL ALMACEN '''''''''''''''''''''''
    Txtnumvald.Caption = Calcula_Numero(Trim(Txtcodpar.Text), "I")
    ACTUALIZAR1
    GRABAR_INGRESO
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    End If
    
End Sub
Private Function ValidarDatos() As Boolean
If Trim(Txtcodalm.Text) <> "" Then
    If Trim(Txtcodpar.Text) <> "" Then
        If Trim(Txtcodori.Text) <> "" Then
'           If dxDBGrid1.Columns.ColumnByFieldName(f5codpro) <> "" Then
            ValidarDatos = True
'            Else
'                MsgBox "Ingrese el Codigo de centro de Costo", vbCritical, "Sistema de Inventario"
'                ValidarDatos = False
'                dxDBGrid1.SetFocus
'                dxDBGrid1.Dataset.ADODataset.Requery
'            End If
        Else
            MsgBox "Ingrese el Codigo de centro de Costo", vbCritical, "Sistema de Inventario"
            ValidarDatos = False
            Txtcodori.SetFocus
        End If

    Else
        MsgBox "Ingrese el Codigo de almacen de Destino", vbCritical, "Sistema de Inventario"
        ValidarDatos = False
        Txtcodpar.SetFocus
    End If
Else
    MsgBox "Ingrese el Codigo de almacen de Origen", vbCritical, "Sistema de Inventario"
    ValidarDatos = False
    Txtcodalm.SetFocus
End If
End Function
Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.ID
        Case "ID_Nuevo"
            Nuevo
            SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = False
            SSActiveToolBars1.Tools("ID_Buscar").Enabled = False
            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
        Case "ID_Grabar"
            If ValidarDatos Then
                Me.MousePointer = 11
                Verifica_Datos
                dxDBGrid1.Dataset.Edit
                If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                     dxDBGrid1.Dataset.Post
                     sw_detalle = True
                End If
                If sw_cabecera = True Or sw_detalle = True Then
                     Graba_Datos
                     sw_nuevo_doc = False
                     sw_detalle = False
                     sw_cabecera = False
                End If
                Me.MousePointer = 1
                SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
                SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
                SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
                SSActiveToolBars1.Tools("ID_Buscar").Enabled = True
                SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
            End If
        Case "ID_Imprimir"
            Me.MousePointer = 11
            imprimir
            Me.MousePointer = 1
            'SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
            SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
            SSActiveToolBars1.Tools("ID_Buscar").Enabled = True
            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
        Case "ID_Buscar"
            Me.MousePointer = 11
            Buscar
            Me.MousePointer = 1
            SSActiveToolBars1.Tools("ID_Buscar").Enabled = True
            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = True
            SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = False
            SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
        Case "ID_Eliminar"
            Me.MousePointer = 11
            Elimina_Movimientos
            Me.MousePointer = 1
            SSActiveToolBars1.Tools("ID_Eliminar").Enabled = False
            SSActiveToolBars1.Tools("ID_Nuevo").Enabled = True
            SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
            SSActiveToolBars1.Tools("ID_Buscar").Enabled = True
            SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
        Case "ID_CargarData"
            frmExcel.Show 1
        Case "ID_Salir"
        
            Me.MousePointer = 11
            If dxDBGrid1.Dataset.State = dsEdit Then
                dxDBGrid1.Dataset.Post
                sw_nuevo_item = True
            End If
            If sw_cabecera = True Or sw_detalle = True Then
                If MsgBox("Desea Grabar el Movimiento?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                    If ValidarDatos Then
                        Graba_Datos
                        sw_nuevo_doc = False
                        sw_detalle = False
                        sw_cabecera = False
                       
                    End If
                Else
                    Me.MousePointer = 1
                    Unload Me
                End If
            End If
            If sw_cabecera = False Then
                Me.MousePointer = 1
                Unload Me
            End If
    End Select

End Sub

Private Sub Txtcodalm_Change()
    
    If Trim(Txtcodalm.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
    
End Sub

Private Sub Txtcodori_Change()

    If Trim(Txtcodori.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
    
End Sub

Private Sub Txtcodori_DblClick()
    
    Txtcodori_KeyDown 113, 0
    
End Sub

Private Sub Txtcodori_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtobserva.SetFocus
        
    End If
    
End Sub

Private Sub Txtcodori_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        sw_ayudaO = True
        wconcepto = ""
        wtipmov = "S"
        'hlp_conceptos_inv.Show 1
        ayuda_conceptos.Show 1
        sw_ayudaO = False
        wtipmov = ""
        If Len(Trim(wconcepto)) > 0 Then
            Txtcodori.Text = wconcepto
            PnlNomOri.Caption = wnomconcepto
            Txtcodori_KeyPress 13
        End If
    End If
   
End Sub

Private Sub Txtcodori_LostFocus()
 
    If sw_ayudaO = False Then
        If Len(Trim(Txtcodori.Text)) > 0 Then
            wnomori = ""
            If Txtcodori.Text <> "" Then
                If VALIDA_ORIGEN(Txtcodori.Text) = True Then
                    wcodori = Txtcodori.Text
                    PnlNomOri.Caption = wnomori
                Else
                    MsgBox "Código de Origen no existe. Verifique.", vbInformation, "Atención"
                    Txtcodori.SetFocus
                End If
            Else
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub TxtCodPar_Change()

    If Trim(Txtcodpar.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
    
End Sub

Private Sub TxtCodPar_DblClick()
    
    TxtCodPar_KeyDown 113, 0
    
End Sub

Private Sub Txtcodpar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Txtcodori.SetFocus

End Sub

Private Sub TxtCodPar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        'hlp_almacenes.Show 1
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            Txtcodpar.Text = wcod_alm
            PnlNomPar.Caption = wnomalmacen
            Txtcodpar_KeyPress 13
        End If
    End If

End Sub

Private Sub TxtCodPar_LostFocus()
        
    If sw_ayuda = False Then
        If Len(Trim(Txtcodpar.Text)) > 0 Then
            WNomPar = ""
            If Txtcodpar.Text <> "" Then
                If Txtcodalm.Text = Txtcodpar.Text Then
                    MsgBox "El Almacén de Origen y el De Destino no Pueden ser los Mismos", vbInformation, "Sistema de Inventario"
                    Exit Sub
                End If
                If VALIDA_ALMACENO(Txtcodpar.Text) = True Then
                    WcodPar = Trim(Txtcodpar.Text)
                    PnlNomPar.Caption = WNomPar
                Else
                    MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
                    Txtcodpar.SetFocus
                End If
            Else
                Exit Sub
            End If
        End If
    End If
   
End Sub

Private Sub Txtfecmov_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"
    
End Sub

Private Sub Verifica_Datos()
Dim Mensaje As String

    If Trim(Txtcodalm.Text) = "" Or Trim(Txtcodpar.Text) = "" Then
        MsgBox "El Almacén de Origen y el de Destino no Pueden ser los Mismos", vbInformation, "Sistema de Inventarios"
        Exit Sub
    End If
    If dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").SummaryFooterValue > 0 Then 'And (sw_cabecera = True Or sw_detalle = True) Then
        If sw_nuevo_doc = False Then
        Mensaje = "La Transferencia no ha sido grabada ... Desea Grabar ?"
        Else
        Mensaje = "Desea Grabar la Transferencia ... ?"
        End If
    Else
        Mensaje = "Debe Registrar Por lo Menos un Item para Generar un Vale"
    End If
    MsgBox Mensaje, vbYesNo + vbInformation, "Sistema de Inventarios"

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
    End With
    AdicionaItem
    
    dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Visible = False
    
    Select Case wvisualiza_cod
        Case "I"
            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = True
            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = False
        Case "F"
            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = False
            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = True
        Case "T"
            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = True
            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = True
        Case Else
            dxDBGrid1.Columns.ColumnByFieldName("CODPROD").Visible = True
            dxDBGrid1.Columns.ColumnByFieldName("CODFAB").Visible = True
    End Select
    
End Sub

Private Sub AdicionaItem()
Dim sw_nuevo_temp   As Boolean

    dxDBGrid1.Dataset.Active = False
    If sw_nuevo_doc = False Then
        DELETEREC_N "tmpstock", Temp
        dxDBGrid1.Dataset.Refresh
    End If
    dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
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
            .FieldValues("F3ITEM") = i
            .FieldValues("F5CODPRO") = ""
            .FieldValues("F5NOMPRO") = ""
            .FieldValues("F6VALPRO") = Format(0, "0.00")
            .FieldValues("F6CANMOV") = Format(0, "0.00")
            .FieldValues("F7SIGMED") = ""
            .FieldValues("F6TOTAL") = Format(0, "0.00")
            .FieldValues("F5CODFAB") = ""
            .FieldValues("marca") = ""
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Visible = False
    'dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Visible = False
    dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Visible = False
    dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Visible = False
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

End Sub

Public Sub BASE_TEMPORAL(Base As String)
    
    Set Temp = New ADODB.Connection
    CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & Base & "; Persist Security Info=False"
    Temp.Open CON

End Sub

Private Sub Txtcodalm_DblClick()

    Txtcodalm_KeyDown 113, 0
    
End Sub

Private Sub Txtcodalm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        'hlp_almacenes.Show 1
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            Txtcodalm.Text = wcod_alm
            PnlNomAlm.Caption = wnomalmacen
            Txtcodalm_KeyPress 13
        End If
    End If
    
End Sub

Private Sub Txtcodalm_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Txtcodpar.SetFocus
    
End Sub

Private Sub txtcodalm_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(Txtcodalm.Text)) > 0 Then
            wnomalmacen = ""
            If Txtcodalm.Text <> "" Then
                If VALIDA_ALMACEN(Txtcodalm.Text) = True Then
                    wcod_alm = Txtcodalm.Text
                    PnlNomAlm.Caption = wnomalmacen
                Else
                    MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
                    Txtcodalm.SetFocus
                End If
            Else
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub Nuevo()
    
    Me.MousePointer = 11
    sw_nuevo_doc = False
    sw_detalle = False
    Txtcodalm.Text = "": PnlNomAlm.Caption = ""
    Txtcodpar.Text = "": PnlNomPar.Caption = ""
    Txtcodori.Text = "": PnlNomOri.Caption = ""
    Txtnumvalo.Caption = "": Txtnumvald.Caption = ""
    
    AdicionaItem
    AdicionaItem
    Txtcodalm.SetFocus
    sw_nuevo_doc = True
    Me.MousePointer = 1
   
End Sub

Private Sub ACTUALIZAR()

    amovs_vale(0).campo = "F1VALSAL" & Format(wmes, "00"): amovs_vale(0).valor = Txtnumvalo.Caption: amovs_vale(0).TIPO = "T"
    GRABA_REGISTRO amovs_vale(), "EF2ALMACENES", "M", 0, cnn_dbbancos, "F2CODALM ='" & Trim(Txtcodalm.Text) & "'"

End Sub

Private Sub GRABAR_SALIDA()
Dim intcont  As Integer
Dim wtotord  As Double
Dim TOTA As Double

    Set rsconsulta = New ADODB.Recordset
    wtotord = 0#
    intcont = 0#
    If sw_nuevo_doc = True Then
        ctipo = "A"
    Else
        ctipo = "M"
    End If
    '---------------------ASIGNA DATOS A IF4VALES -------------------------------------
    amovs_cab(0).campo = "F2CODALM": amovs_cab(0).valor = Txtcodalm.Text: amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "F4NUMVAL": amovs_cab(1).valor = Txtnumvalo.Caption: amovs_cab(1).TIPO = "T"
    amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = TxtFecMov.Value: amovs_cab(2).TIPO = "F"
    amovs_cab(3).campo = "F2CODPAR": amovs_cab(3).valor = WcodPar: amovs_cab(3).TIPO = "T"
    amovs_cab(4).campo = "F2CODPROV": amovs_cab(4).valor = "0": amovs_cab(4).TIPO = "T"
    amovs_cab(5).campo = "F1CODORI": amovs_cab(5).valor = wcodori: amovs_cab(5).TIPO = "T"
    amovs_cab(6).campo = "F1CODDOC": amovs_cab(6).valor = "03": amovs_cab(6).TIPO = "T"
    amovs_cab(7).campo = "F4NUMDOC": amovs_cab(7).valor = Txtnumvald.Caption: amovs_cab(7).TIPO = "T"
    amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = wmoneda_productos: amovs_cab(8).TIPO = "T"
    amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = wtipcam: amovs_cab(9).TIPO = "N"
    amovs_cab(10).campo = "F4FECULT": amovs_cab(10).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(10).TIPO = "F"
    amovs_cab(11).campo = "F2CODUSE": amovs_cab(11).valor = wusuario: amovs_cab(11).TIPO = "T"
    If ctipo = "A" Then
        amovs_cab(12).campo = "F4FECGRA": amovs_cab(12).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(12).TIPO = "F"
        amovs_cab(13).campo = "F4USEGRA": amovs_cab(13).valor = wusuario: amovs_cab(13).TIPO = "T"
    Else
        amovs_cab(12).campo = "F4FECMOD": amovs_cab(12).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(12).TIPO = "F"
        amovs_cab(13).campo = "F4USEMOD": amovs_cab(13).valor = wusuario: amovs_cab(13).TIPO = "T"
    End If
    amovs_cab(14).campo = "F4OBSERVA": amovs_cab(14).valor = txtobserva.Text: amovs_cab(14).TIPO = "T"
    '---------------------ASIGNA DATOS A IF3VALES ----------------------------------------
    amovs_det(0).campo = "F4NUMVAL": amovs_det(0).valor = "": amovs_det(0).TIPO = "T"
    amovs_det(1).campo = "F5CODPRO": amovs_det(1).valor = "": amovs_det(1).TIPO = "T"
    amovs_det(2).campo = "F3CANPRO": amovs_det(2).valor = "": amovs_det(2).TIPO = "N"
    amovs_det(3).campo = "F3VALVTA": amovs_det(3).valor = "": amovs_det(3).TIPO = "N"
    amovs_det(4).campo = "F2CODALM": amovs_det(4).valor = "": amovs_det(4).TIPO = "T"
    amovs_det(5).campo = "F4FECVAL": amovs_det(5).valor = "": amovs_det(5).TIPO = "F"
    amovs_det(6).campo = "F3VALDOL": amovs_det(6).valor = "": amovs_det(6).TIPO = "N"
    amovs_det(7).campo = "F3TOTITE": amovs_det(7).valor = "": amovs_det(7).TIPO = "N"
    amovs_det(8).campo = "F3TOTDOL": amovs_det(8).valor = "": amovs_det(8).TIPO = "N"
    '---------------------Calcula el Numero de Filas
    nitems = 0
    If rsconsulta.State = adStateOpen Then rsconsulta.Close
    SQL = "Select count(F3ITEM) as NITEM from TmpStock Where LEN(TRIM(F3ITEM))> 0 "
    rsconsulta.Open SQL, Temp, adOpenDynamic, adLockOptimistic
    If Not rsconsulta.EOF Then
    nitems = Val("" & rsconsulta.Fields("NITEM"))
    End If
    rsconsulta.Close
    
    ReDim Values(8, nitems)
    If rsconsulta.State = adStateOpen Then rsconsulta.Close
    rsconsulta.Open "Select * from TmpStock", Temp
    If Not rsconsulta.EOF Then
        nfila = 0
        rsconsulta.MoveFirst
        Do While Not rsconsulta.EOF
            If Len(Trim(rsconsulta.Fields("F5CODPRO") & "")) > 0 Then
                Values(0, nfila) = "" & Txtnumvalo.Caption
                Values(1, nfila) = "" & rsconsulta.Fields("F5CODPRO")
                Values(2, nfila) = "" & rsconsulta.Fields("F6CANMOV")
                Values(3, nfila) = "" & Format(rsconsulta.Fields("F6TOTAL") * wtipcam, "0.00")
                Values(4, nfila) = "" & Txtcodalm.Text
                Values(5, nfila) = "" & CVDate(TxtFecMov.Value)
                Values(6, nfila) = "" & rsconsulta.Fields("F6TOTAL")
                Values(7, nfila) = "" & Format(Val(rsconsulta.Fields("F6CANMOV") * Values(3, nfila)), "0.00")
                Values(8, nfila) = "" & Format(Val(rsconsulta.Fields("F6CANMOV") * rsconsulta.Fields("F6TOTAL")), "0.00")
                Vales_Detalle Txtnumvalo.Caption, rsconsulta.Fields("F5CODPRO"), rsconsulta.Fields("F6CANMOV"), (Format(rsconsulta.Fields("F6TOTAL") * wtipcam, "0.00")), Txtcodalm.Text, CVDate(TxtFecMov.Value), rsconsulta.Fields("F6TOTAL")
                nfila = nfila + 1
            End If
            rsconsulta.MoveNext
        Loop
    End If
    rsconsulta.Close
    sw_graba_registro = True
    
    If ctipo = "A" Then '---Nuevo
        '-----Graba Cabecera
        GRABA_REGISTRO amovs_cab(), "IF4VALES", "A", 14, cnn_dbbancos, ""
        If sw_graba_registro = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "", Values(), nfila - 1, "111111111111", "", ""
        End If
    Else    '----------Modificacion
        '------- GRABA CABECERA
        GRABA_REGISTRO amovs_cab(), "IF4VALES", "M", 14, cnn_dbbancos, "F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'"
        '------- GRABA DETALLE
        cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'")
        GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "F4NUMVAL  = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'", Values(), nfila - 1, "111111111111", "", ""
    End If

    '''graba envio
    If wIndEnvia = "*" Then 'cnn_dbEnvia
        SQL = "delete from if4vales where f2codalm='" & Txtcodalm.Text & "' and f4numval='" & Txtnumvalo.Caption & "'"
        cnn_dbEnvia.Execute SQL
        
        GRABA_REGISTRO amovs_cab(), "IF4VALES", "A", 14, cnn_dbEnvia, ""
        If sw_graba_registro = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbEnvia, "", Values(), nfila - 1, "111111111111", "", ""
        End If
    End If
    ''''''''
End Sub

Private Sub ACTUALIZAR1()

    amovs_vale(0).campo = "F1VALING" & Format(wmes, "00"): amovs_vale(0).valor = Txtnumvald.Caption: amovs_vale(0).TIPO = "T"
    GRABA_REGISTRO amovs_vale(), "EF2ALMACENES", "M", 0, cnn_dbbancos, "F2CODALM ='" & Trim(Txtcodpar.Text) & "'"

End Sub

Private Sub GRABAR_INGRESO()
Dim intcont  As Integer
Dim wtotord  As Double
Dim TOTA As Double

    Set rsconsulta = New ADODB.Recordset
    wtotord = 0#
    intcont = 0#
    If sw_nuevo_doc = True Then
        ctipo = "A"
    Else
        ctipo = "M"
    End If
    '---------------------ASIGNA DATOS A IF4VALES -------------------------------------
    amovs_cab(0).campo = "F2CODALM": amovs_cab(0).valor = Txtcodpar.Text: amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "F4NUMVAL": amovs_cab(1).valor = Txtnumvald.Caption: amovs_cab(1).TIPO = "T"
    amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = TxtFecMov.Value: amovs_cab(2).TIPO = "F"
    amovs_cab(3).campo = "F2CODPAR": amovs_cab(3).valor = wcod_alm: amovs_cab(3).TIPO = "T"
    amovs_cab(4).campo = "F2CODPROV": amovs_cab(4).valor = "0": amovs_cab(4).TIPO = "T"
    amovs_cab(5).campo = "F1CODORI": amovs_cab(5).valor = wcodori: amovs_cab(5).TIPO = "T"
    amovs_cab(6).campo = "F1CODDOC": amovs_cab(6).valor = "03": amovs_cab(6).TIPO = "T"
    amovs_cab(7).campo = "F4NUMDOC": amovs_cab(7).valor = Txtnumvalo.Caption: amovs_cab(7).TIPO = "T"
    amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = wmoneda_productos: amovs_cab(8).TIPO = "T"
    amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = wtipcam: amovs_cab(9).TIPO = "N"
    amovs_cab(10).campo = "F4FECULT": amovs_cab(10).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(10).TIPO = "F"
    amovs_cab(11).campo = "F2CODUSE": amovs_cab(11).valor = wusuario: amovs_cab(11).TIPO = "T"
    If ctipo = "A" Then
        amovs_cab(12).campo = "F4FECGRA": amovs_cab(12).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(12).TIPO = "F"
        amovs_cab(13).campo = "F4USEGRA": amovs_cab(13).valor = wusuario: amovs_cab(13).TIPO = "T"
    Else
        amovs_cab(12).campo = "F4FECMOD": amovs_cab(12).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(12).TIPO = "F"
        amovs_cab(13).campo = "F4USEMOD": amovs_cab(13).valor = wusuario: amovs_cab(13).TIPO = "T"
    End If
    amovs_cab(14).campo = "F4OBSERVA": amovs_cab(14).valor = txtobserva.Text: amovs_cab(14).TIPO = "T"
    '---------------------ASIGNA DATOS A IF3VALES ----------------------------------------
    amovs_det(0).campo = "F4NUMVAL": amovs_det(0).valor = "": amovs_det(0).TIPO = "T"
    amovs_det(1).campo = "F5CODPRO": amovs_det(1).valor = "": amovs_det(1).TIPO = "T"
    amovs_det(2).campo = "F3CANPRO": amovs_det(2).valor = "": amovs_det(2).TIPO = "N"
    amovs_det(3).campo = "F3VALVTA": amovs_det(3).valor = "": amovs_det(3).TIPO = "N"
    amovs_det(4).campo = "F2CODALM": amovs_det(4).valor = "": amovs_det(4).TIPO = "T"
    amovs_det(5).campo = "F4FECVAL": amovs_det(5).valor = "": amovs_det(5).TIPO = "F"
    amovs_det(6).campo = "F3VALDOL": amovs_det(6).valor = "": amovs_det(6).TIPO = "N"
    amovs_det(7).campo = "F3TOTITE": amovs_det(7).valor = "": amovs_det(7).TIPO = "N"
    amovs_det(8).campo = "F3TOTDOL": amovs_det(8).valor = "": amovs_det(8).TIPO = "N"
    
    '---------------------Calcula el Numero de Filas
    nitems = 0
    If rsconsulta.State = adStateOpen Then rsconsulta.Close
    SQL = "Select count(F3ITEM) as NITEM from TmpStock Where LEN(TRIM(F3ITEM))> 0 "
    rsconsulta.Open SQL, Temp, adOpenDynamic, adLockOptimistic
    If Not rsconsulta.EOF Then
    nitems = Val("" & rsconsulta.Fields("NITEM"))
    End If
    rsconsulta.Close
    
    ReDim Values(8, nitems)
    
    If rsconsulta.State = adStateOpen Then rsconsulta.Close
    rsconsulta.Open "Select * from TmpStock", Temp
    If Not rsconsulta.EOF Then
        nfila = 0
        rsconsulta.MoveFirst
        Do While Not rsconsulta.EOF
            If Len(Trim(rsconsulta.Fields("F5CODPRO") & "")) > 0 Then
                Values(0, nfila) = "" & Txtnumvald.Caption
                Values(1, nfila) = "" & rsconsulta.Fields("F5CODPRO")
                Values(2, nfila) = "" & rsconsulta.Fields("F6CANMOV")
                Values(3, nfila) = "" & Format(rsconsulta.Fields("F6TOTAL") * wtipcam, "0.00")
                Values(4, nfila) = "" & Txtcodpar.Text
                Values(5, nfila) = "" & CVDate(TxtFecMov.Value)
                Values(6, nfila) = "" & rsconsulta.Fields("F6TOTAL")
                Values(7, nfila) = "" & Format(Val(rsconsulta.Fields("F6CANMOV") * Values(3, nfila)), "0.00")
                Values(8, nfila) = "" & Format(Val(rsconsulta.Fields("F6CANMOV") * rsconsulta.Fields("F6TOTAL")), "0.00")
                Vales_Detalle Txtnumvald.Caption, rsconsulta.Fields("F5CODPRO"), rsconsulta.Fields("F6CANMOV"), Format(rsconsulta.Fields("F6TOTAL") * wtipcam, "0.00"), Txtcodpar.Text, CVDate(TxtFecMov.Value), rsconsulta.Fields("F6TOTAL")
                nfila = nfila + 1
            End If
            rsconsulta.MoveNext
        Loop
    End If
    rsconsulta.Close
    sw_graba_registro = True
    
    If ctipo = "A" Then '---Nuevo
        '-----Graba Cabecera
        GRABA_REGISTRO amovs_cab(), "IF4VALES", "A", 14, cnn_dbbancos, ""
        If sw_graba_registro = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "", Values(), nfila - 1, "111111111111", "", ""
        End If
    Else    '----------Modificacion
        '------- GRABA CABECERA
        GRABA_REGISTRO amovs_cab(), "IF4VALES", "M", 14, cnn_dbbancos, "F4NUMVAL = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'"
        '------- GRABA DETALLE
        cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'")
        GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "F4NUMVAL  = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'", Values(), nfila - 1, "111111111111", "", ""
    End If
    
    
    '''graba envio
    If wIndEnvia = "*" Then 'cnn_dbEnvia
        SQL = "delete from if4vales where f2codalm='" & Txtcodpar.Text & "' and f4numval='" & Txtnumvald.Caption & "'"
        cnn_dbEnvia.Execute SQL
        
        GRABA_REGISTRO amovs_cab(), "IF4VALES", "A", 14, cnn_dbEnvia, ""
        If sw_graba_registro = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbEnvia, "", Values(), nfila - 1, "111111111111", "", ""
        End If
    End If
    ''''''''

End Sub

Private Sub Buscar()

    Rem Gtipval = "3"
    Rem FrmAyudaVale.Show 1
    Rem Txtcodalm.Text = wcod_alm
    Rem Txtnumvalo.Caption = wnumval
    Rem Actualiza_Datos

End Sub

Private Sub Agrega_Items()
Dim SSQL, csql As String

    Set RsProducto = New ADODB.Recordset
    Set RsMovAlmacen = New ADODB.Recordset
    
    dxDBGrid1.Dataset.Close
    DELETEREC_N "TmpSTOCK", Temp
    
    SSQL = "SELECT * FROM IF3VALES WHERE F2CODALM ='" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "' AND CVDATE(F4FECVAL) = '" & CVDate(TxtFecMov.Value) & "'"
    If RsStockDet.State = adStateOpen Then RsStockDet.Close
    RsStockDet.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockDet.EOF Then
        RsStockDet.MoveFirst
        Do While Not RsStockDet.EOF
            With dxDBGrid1.Dataset
                gcod = "" & RsStockDet.Fields("f5codpro")
                csql = "SELECT f5nompro,F7codmed,F5CODFAB,F5VALVTA FROM IF5PLA WHERE F5CODPRO = '" & gcod & "'"
                If RsProducto.State = adStateOpen Then RsProducto.Close
                RsProducto.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsProducto.EOF Then
                    gnom = "" & RsProducto.Fields("f5nompro")
                    guni = "" & RsProducto.Fields("F7codmed")
                    gfab = "" & RsProducto.Fields("F5CODFAB")
                    gvalvta = "" & RsProducto.Fields("F5VALVTA")
                End If
                RsProducto.Close
                gcanmov = Val(Format(RsStockDet.Fields("f3canpro"), "#0.000"))
                SQL = "SELECT f6stockact FROM IF6ALMA WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F5CODPRO = '" & RsStockDet.Fields("F5CODPRO") & "'"
                If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
                RsMovAlmacen.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RsMovAlmacen.EOF Then
                    RsMovAlmacen.MoveFirst
                    gstock = Val(Format(RsMovAlmacen.Fields("f6stockact"), "#0.000"))
                End If
                RsMovAlmacen.Close
                SQL = "INSERT INTO TmpStock(F5CODPRO,F5NOMPRO,F6VALPRO,F6CANMOV,F7SIGMED,F6TOTAL,F5CODFAB) VALUES('" & gcod & "','" & gnom & "' ," & gstock & "," & gcanmov & ",'" & guni & "'," & gvalvta & ",'" & gfab & "')"
                Temp.Execute (SQL)
                RsStockDet.MoveNext
                If RsStockDet.EOF Then Exit Do
            End With
        Loop
        dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.Dataset.Open
        dxDBGrid1.Dataset.First
        dxDBGrid1.Columns.FocusedIndex = 0
    End If

End Sub

Private Sub imprimir()

    With Acr_Vale_Salida
        .DataControl1.ConnectionString = cnn_dbbancos
        .DataControl1.Source = "SELECT *,IF5PLA.F5CODPRO AS CODIGO,F5NOMPRO,F7CODMED,F5CODFAB,F2DESMAR FROM IF3VALES ,IF5PLA ,ef2marcas  WHERE IF5PLA.F5CODPRO=IF3VALES.F5CODPRO and ef2marcaS.f2codmar=if5pla.f5marca AND F2CODALM='" & Txtcodalm.Text & "' AND F4NUMVAL='" & Txtnumvalo.Caption & "'"
        .fldempresa.Text = wempresa
        .fldfecha.Text = TxtFecMov.Value
        .Lbl_vale.Caption = "VALE DE SALIDA"
        .fldalma.Text = Txtcodalm.Text
        .fldalmacen.Text = PnlNomAlm.Caption
        .fldvale.Text = Txtnumvalo.Caption
        .fldcon.Text = Txtcodori.Text
        .F1NOMORI.Text = PnlNomOri.Caption
        .flddoc.Visible = False
        .NUMDOC.Visible = False
        .Show 1
    End With
    
    With Acr_vale_ingreso
        .DataControl1.ConnectionString = cnn_dbbancos
        '.DataControl1.Source = "SELECT *,F5NOMPRO,F7CODMED,f5codfab,f2desmar FROM IF3VALES,IF5PLA,ef2marcas WHERE IF5PLA.F5CODPRO=IF3VALES.F5CODPRO and ef2marcaS.f2codmar=if5pla.f5marca AND F2CODALM='" & Txtcodpar.Text & "' AND F4NUMVAL='" & Txtnumvald.Caption & "'"
        .DataControl1.Source = "SELECT *,IF5PLA.F5CODPRO AS CODIGO,F5NOMPRO,F7CODMED,F5CODFAB,F2DESMAR FROM IF3VALES ,IF5PLA ,ef2marcas WHERE IF5PLA.F5CODPRO=IF3VALES.F5CODPRO and ef2marcaS.f2codmar=if5pla.f5marca AND F2CODALM='" & Txtcodpar.Text & "' AND F4NUMVAL='" & Txtnumvald.Caption & "'"
        .fldempresa.Text = wempresa
        .fldfecha.Text = TxtFecMov.Value
        .Lbl_vale.Caption = "VALE DE INGRESO"
        .fldalma.Text = Txtcodpar.Text
        .fldalmacen.Text = PnlNomPar.Caption
        .fldvale.Text = Txtnumvald.Caption
        .fldcon.Text = Txtcodori.Text
        .F1NOMORI.Text = PnlNomOri.Caption
        .flddoc.Visible = True
        .NUMDOC.Visible = True
        .flddoc.Text = Txtnumvalo.Caption
        
        .Show 1
    End With
    
    'ImprimendoV Trim(Txtcodalm.Text), Trim(Txtnumvalo), 0
    'ImprimendoV Trim(Txtcodpar.Text), Trim(Txtnumvald), 0

End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 115 Or KeyCode = 46 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_detalle = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
            End If
        End If
    End If
    If KeyCode = 113 Then
        Me.MousePointer = 11
        sw_ayuda_prod = True
        wcod_alm = Txtcodalm.Text
         If Trim(wcod_alm) = "" Then
            MsgBox "Ingrese Almacen de origen", vbCritical, "Sistema de Logistica"
            Exit Sub
        End If
        'hlp_productos.Show vbModal
        Con_Ayu = 4
        ayuda_productos.Show 1
        
        If Len(Trim(wcodproducto)) <> 0 Then
            sw_detalle = True
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
            dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
            dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = wmedida
            dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
            dxDBGrid1.Columns.ColumnByFieldName("MARCA").Value = wmarca
            dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = wstockact
            dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = wprecos
            dxDBGrid1.Dataset.Post
        End If
        Me.MousePointer = 1
        dxDBGrid1.Columns.FocusedIndex = 6
End If
End Sub

Private Sub txtobserva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
        dxDBGrid1.Columns.FocusedIndex = 1
    Else
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End If
End Sub
