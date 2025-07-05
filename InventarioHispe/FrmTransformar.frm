VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form FrmTransforma 
   Caption         =   "Transformaciones"
   ClientHeight    =   7050
   ClientLeft      =   2025
   ClientTop       =   3030
   ClientWidth     =   10335
   LinkTopic       =   "Form3"
   ScaleHeight     =   7050
   ScaleWidth      =   10335
   Begin Threed.SSPanel SSPanel2 
      Height          =   6825
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _Version        =   65536
      _ExtentX        =   17806
      _ExtentY        =   12039
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
      Begin Threed.SSFrame SSFrame2 
         Height          =   2895
         Left            =   165
         TabIndex        =   12
         Top             =   3840
         Width           =   9810
         _Version        =   65536
         _ExtentX        =   17304
         _ExtentY        =   5106
         _StockProps     =   14
         Caption         =   "Insumos de Productos a Transferir"
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
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
            Height          =   2415
            Left            =   120
            OleObjectBlob   =   "FrmTransformar.frx":0000
            TabIndex        =   13
            Top             =   360
            Width           =   9555
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   2055
         Left            =   150
         TabIndex        =   11
         Top             =   1680
         Width           =   9825
         _Version        =   65536
         _ExtentX        =   17330
         _ExtentY        =   3625
         _StockProps     =   14
         Caption         =   "Productos a Ingresar"
         ForeColor       =   0
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
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "FrmTransformar.frx":3F12
            TabIndex        =   19
            Top             =   240
            Width           =   9585
         End
      End
      Begin VB.TextBox Txtcodori 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "XTR"
         Top             =   780
         Width           =   465
      End
      Begin VB.TextBox Txtcodpar 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   2
         Top             =   465
         Width           =   465
      End
      Begin VB.TextBox Txtcodalm 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   1
         Top             =   165
         Width           =   465
      End
      Begin aBoxCtl.aBox TxtFecMov 
         Height          =   315
         Left            =   8520
         TabIndex        =   10
         Top             =   165
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
         Text            =   "31/01/2003"
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
         ButtonPicture   =   "FrmTransformar.frx":7E3A
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
      Begin Threed.SSPanel Txtnumvalo 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   1200
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483644
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
      Begin Threed.SSPanel Txtnumvald 
         Height          =   285
         Left            =   4890
         TabIndex        =   15
         Top             =   1200
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   -2147483644
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
      Begin Threed.SSPanel PnlNomAlm 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   165
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel PnlNomOri 
         Height          =   285
         Left            =   2520
         TabIndex        =   17
         Top             =   780
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Transformaciòn"
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
      End
      Begin Threed.SSPanel PnlNomPar 
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   465
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Concepto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Vale Ingreso:"
         DataField       =   "<"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3675
         TabIndex        =   8
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label LabPar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Almacén Destino:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   315
         TabIndex        =   7
         Top             =   525
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Almacén Origen:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   315
         TabIndex        =   6
         Top             =   195
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Fecha:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   7680
         TabIndex        =   5
         Top             =   165
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Vale Salida:"
         DataField       =   "<"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   1185
         Width           =   840
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   135
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "FrmTransformar.frx":818C
      ToolBars        =   "FrmTransformar.frx":CD99
   End
End
Attribute VB_Name = "FrmTransforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------
Dim sw_nuevo_doc, sw_ayuda, sw_ayudaO As Boolean
Dim sw_nuevo_item, sw_detalle, sw_cabecera As Boolean
Dim Temp As ADODB.Connection
Dim Temp1 As ADODB.Connection
Dim amovs_vale(0 To 0)  As a_grabacion
Dim amovs_cab(0 To 11) As a_grabacion
Dim amovs_det(0 To 8) As a_grabacion

Dim flag    As Boolean
Dim CANTI   As Double
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


Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Inventario") = vbYes Then
            If dxDBGrid1.Dataset.RecNo = 1 Then
               dxDBGrid1.Dataset.Delete
               AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
                RENUMERARITEMS
            End If
        End If
    End If
End Sub

Private Sub dxDBGrid2_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 115 Then
    If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Inventario") = vbYes Then
        If dxDBGrid2.Dataset.RecNo = 1 Then
           dxDBGrid2.Dataset.Delete
           AdicionaItem2
        Else
            dxDBGrid2.Dataset.Delete
            RENUMERARITEMS2
        End If
    End If
End If
End Sub

Private Sub Form_Load()


    sw_nuevo_doc = True
    sw_detalle = False
    sw_nuevo_item = False
    
    Me.Height = 8040
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    
    '-------PARA GRID 1 ----------
    BASE_TEMPORAL "TEMPFAC.MDB"
    DELETEREC_N "TmpStock", Temp
    dxDBGrid1.Dataset.Refresh
    Conf_Grid
    
    '--------PARA GRID 2--------
    BASE_TEMPORAL "TEMFORMU.MDB"
    DELETEREC_N "TEMFORM", Temp
    dxDBGrid2.Dataset.Refresh
    Conf_Grid2
    
    wtipcam = gtipcam
    If wtipcam = 0# Then wtipcam = 3.55
    TxtFecMov.Value = Format(Now, "dd/mm/yyyy")
    SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
    
End Sub



Public Sub BASE_TEMPORAL(Base As String)
Dim CON As String
Set Temp = New ADODB.Connection

If Temp.State = adStateOpen Then Temp.Close
CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & Base & "; Persist Security Info=False"
Temp.Open CON

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
    
End Sub

Private Sub AdicionaItem()
Dim i As Integer
Dim sw_nuevo_temp   As Boolean

dxDBGrid1.Dataset.Active = False

If sw_nuevo_doc = False Then
    BASE_TEMPORAL "TEMPFAC.MDB"
    DELETEREC_N "TmpStock", Temp
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
    .FieldValues("F6VALPRO") = Format(0, "###,##0.00")
    .FieldValues("F6CANMOV") = Format(0, "###,##0.00")
    .FieldValues("F4VASOS") = Format(0, "###,##0.00")
    .FieldValues("F7SIGMED") = ""
    .FieldValues("F6TOTAL") = Format(0, "###,##0.00")
    .FieldValues("F5CODFAB") = ""
    .FieldValues("F3MONEDA") = ""
    
Next
    
    .Post
    sw_nuevo_item = False

End With

dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Visible = False
dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Visible = False
dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Visible = False
dxDBGrid1.Columns.ColumnByFieldName("F3MONEDA").Visible = False
dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Visible = False

dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open
End Sub

Private Sub Conf_Grid2()
    
    With dxDBGrid2.Options
       
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
  
    AdicionaItem2
    
End Sub

Private Sub AdicionaItem2()
Dim i As Integer
Dim sw_nuevo_temp   As Boolean

dxDBGrid2.Dataset.Active = False

If sw_nuevo_doc = False Then
    BASE_TEMPORAL "TEMFORMU.MDB"
    DELETEREC_N "TEMFORM", Temp
    dxDBGrid2.Dataset.Refresh
End If

dxDBGrid2.Dataset.ADODataset.ConnectionString = Temp
dxDBGrid2.Dataset.Active = True
dxDBGrid2.Dataset.Close
dxDBGrid2.Dataset.Open

With dxDBGrid2.Dataset

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
    .FieldValues("COD_FORMULA") = ""
    .FieldValues("COD_PROD") = ""
    .FieldValues("DES_MAT") = ""
    .FieldValues("U_MEDIDA") = ""
    .FieldValues("CANTIDAD") = Format(0, "###,##0.00")
    .FieldValues("MERMA") = Format(0, "###,##0.00")
    .FieldValues("PRECIO") = Format(0, "###,##0.00")
    .FieldValues("STOCK") = Format(0, "###,##0.00")
    .FieldValues("F3MONEDA") = ""
    
Next
    
    .Post
    sw_nuevo_item = False

End With

dxDBGrid2.Columns.ColumnByFieldName("F3ITEM").Visible = False
dxDBGrid2.Columns.ColumnByFieldName("COD_FORMULA").Visible = False
dxDBGrid2.Columns.ColumnByFieldName("STOCK").Visible = False
dxDBGrid1.Columns.ColumnByFieldName("F3MONEDA").Visible = False


dxDBGrid2.Dataset.Close
dxDBGrid2.Dataset.Open
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.Id
    Case "ID_Nuevo"
        Nuevo
        
    Case "ID_Salir"
        Me.MousePointer = 11
        If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
            dxDBGrid1.Dataset.Post
            sw_nuevo_item = True
        End If
        
        If dxDBGrid2.Dataset.State = dsEdit Or dxDBGrid2.Dataset.State = dsInsert Then
            dxDBGrid2.Dataset.Post
            sw_nuevo_item = True
        End If
        
        If sw_cabecera = True Or sw_detalle = True Then
            If MsgBox("Desea Grabar el Movimiento?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                Graba_Datos
                sw_nuevo_doc = False
                sw_detalle = False
            End If
        End If
        Me.MousePointer = 1
        Unload Me
    Case "ID_Grabar"
        Me.MousePointer = 11
        Verifica_Datos
        dxDBGrid1.Dataset.Edit
        If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
             dxDBGrid1.Dataset.Post
             sw_detalle = True
        End If
        
        If dxDBGrid2.Dataset.State = dsEdit Or dxDBGrid2.Dataset.State = dsInsert Then
             dxDBGrid2.Dataset.Post
             sw_detalle = True
        End If
        
        If sw_cabecera = True Or sw_detalle = True Then
             Graba_Datos
             sw_nuevo_doc = False
             sw_detalle = False
             sw_cabecera = False
             MsgBox "Transformación Grabada", vbInformation, "Sistema de Inventarios"
        End If
        Me.MousePointer = 1
    Case "ID_Imprimir"
        Me.MousePointer = 11
        Imprimir_Datos
        Me.MousePointer = 1
        
    Case "ID_Buscar"
        Buscando

    Case "ID_Eliminar"
        Me.MousePointer = 11
        Elimina_Movimientos
        Me.MousePointer = 1
End Select
End Sub

Private Sub Txtcodalm_Change()
If Trim(Txtcodalm.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtcodalm_DblClick()

    Txtcodalm_KeyDown 113, 0
    
End Sub

Private Sub txtcodalm_GotFocus()

    Txtcodalm.SelStart = 0
    Txtcodalm.SelLength = Len(Txtcodalm.Text)

End Sub

Private Sub Txtcodalm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        'hlp_almacenes.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            Txtcodalm.Text = wcod_alm
            PnlNomalm.Caption = wnomalmacen
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
                    PnlNomalm.Caption = wnomalmacen
                    'REM EMB Txtnumvalo.Caption = Calcula_Numero(Trim(Txtcodalm.Text), "S")
                    'SSFrame2.Caption = "Insumos de Productos a Transferir del Almacén " & Trim(PnlNomalm.Caption)
                    SSFrame2.Caption = "Insumos de Productos a Transferir del Almacén " & Trim(PnlNomalm.Caption)
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

Private Sub Txtcodori_Change()
If Trim(Txtcodori.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtcodori_DblClick()
   Txtcodori_KeyDown 113, 0
End Sub

Private Sub Txtcodori_KeyDown(KeyCode As Integer, Shift As Integer)
    
If KeyCode = 113 Then
        sw_ayudaO = True
        wcodori = ""
        'hlp_Origenes.Show 1
        sw_ayudaO = False
        If Len(Trim(wcodori)) > 0 Then
            Txtcodori.Text = wcodori
            PnlNomOri.Caption = wnomori
            Txtcodori_KeyPress 13
        End If
    End If
   
End Sub

Private Sub Txtcodori_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
        dxDBGrid1.Columns.FocusedIndex = 0
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

Private Sub TxtCodPar_KeyDown(KeyCode As Integer, Shift As Integer)
    
   If KeyCode = 113 Then
        sw_ayuda = True
        WcodPar = ""
        'hlp_almacenes.Show 1
        sw_ayuda = False
        If Len(Trim(WcodPar)) > 0 Then
            Txtcodpar.Text = WcodPar
            PnlNomPar.Caption = WNomPar
            Txtcodpar_KeyPress 13
        End If
    End If


End Sub

Private Sub Txtcodpar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Txtcodori.SetFocus

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
                    'Rem EMB Txtnumvald.Caption = Calcula_Numero(Trim(Txtcodpar.Text), "I")
                    SSFrame1.Caption = "Productos a Ingresar al Almacén " & PnlNomPar.Caption
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

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("f4vasos").Value = Format(0, "###,##0.00")
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

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Set rsconsulta = New ADODB.Recordset
    Me.MousePointer = 11
    'FrmHelpAlmPro.Show 1
    hlp_productos.Show 1
    If Len(Trim(wcodproducto)) <> 0 Then
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
        dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
        If RS.State = adStateOpen Then RS.Close
        RS.Open "SELECT F7SIGMED,F7CODMED FROM EF7MEDIDAS WHERE F7CODMED='" & wmedida & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
            wmedida = "" & RS.Fields("F7SIGMED")
        End If
        dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = wmedida
        dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
        'Sacando el Stock del producto
        If rsif6alma.State = adStateOpen Then rsif6alma.Close
        rsif6alma.Open "SELECT F6STOCKACT FROM IF6ALMA WHERE F5CODPRO = '" & wcodproducto & "" & "' AND F2CODALM = '" & wcod_alm & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif6alma.EOF Then
            dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = Format(Val(rsif6alma.Fields("F6STOCKACT") & ""), "###,###,##0.00")
            wstock = Format(Val(rsif6alma.Fields("F6STOCKACT") & ""), "###,###,##0.00")
        End If
        rsif6alma.Close
        'dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = wstock
        dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = wvalvta
        dxDBGrid1.Columns.ColumnByFieldName("F4VASOS").Value = 0#
        If rsconsulta.State = adStateOpen Then rsconsulta.Close
        rsconsulta.Open "SELECT * FROM IF5PLA WHERE F5CODPRO = '" & wcodproducto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconsulta.EOF Then
            dxDBGrid1.Columns.ColumnByFieldName("F3MONEDA").Value = "" & rsconsulta.Fields("F5MONEDA")
            gmoneda = "" & rsconsulta.Fields("F5MONEDA")
        End If
        VALIDA wcodproducto
        dxDBGrid1.Dataset.Post
    End If
    Me.MousePointer = 1
    dxDBGrid1.Columns.FocusedIndex = 3
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

Set rsconsulta = New ADODB.Recordset

    If dxDBGrid1.Dataset.State <> 0 And dxDBGrid1.Dataset.State <> 1 Then
        
        If dxDBGrid1.Columns.FocusedColumn.FieldName = "F5CODPRO" Then
            wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid1.Dataset.Edit
                If rsconsulta.State = adStateOpen Then rsconsulta.Close                                                                                                                                                                                                     'AND (if6alma.f6stockact > 0)
                SQL = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB,IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO, IF6ALMA.F6STOCKACT, IF5PLA.F5VALVTA, IF5PLA.F7CODMED FROM IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO WHERE  (IF6ALMA.F2CODALM='" & wcod_alm & "') AND ((IF5PLA.F5CODPRO='" & wcodproducto & "') OR (IF5PLA.F5CODFAB='" & wcodproducto & "')) ORDER BY IF5PLA.F5CODPRO;"
                rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsconsulta.EOF Then
                    wcodproducto = "" & Trim(rsconsulta.Fields("F5CODPRO"))
                    wcodfab = "" & rsconsulta.Fields("F5CODFAB")
                    wdesproducto = "" & rsconsulta.Fields("F5NOMPRO")
                    If RS.State = adStateOpen Then RS.Close
                    RS.Open "SELECT F7SIGMED FROM EF7MEDIDAS WHERE F7CODMED='" & wmedida & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not RS.EOF Then
                        wmedida = "" & RS.Fields("F7SIGMED")
                    Else
                        wmedida = "" & rsconsulta.Fields("F7CODMED")
                    End If
                    gmoneda = "" & rsconsulta.Fields("F5MONEDA")
                    wstock = Val("" & rsconsulta.Fields("F6STOCKACT"))
                    wvalvta = Val("" & rsconsulta.Fields("F5VALVTA"))
                    VALIDA wcodproducto
                Else
                    MsgBox "Código no existe", vbInformation + vbDefaultButton1, "Atención"
                    wcodproducto = "": wdesproducto = "": wmedida = "": wcodfab = ""
                    wvalvta = 0#: wstock = 0#: gmoneda = ""
                    dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
                    dxDBGrid1.Dataset.Post
                    Exit Sub
                End If
                dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
                dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = wmedida
                dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
                dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = wstock
                dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = wvalvta
                dxDBGrid1.Columns.ColumnByFieldName("F4VASOS").Value = 0#
                dxDBGrid1.Columns.ColumnByFieldName("F5MONEDA").Value = gmoneda
                dxDBGrid1.Dataset.Post
                
                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").ColIndex - 1
            End If
        Else
            If dxDBGrid1.Columns.FocusedColumn.FieldName = "F6CANMOV" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "F4VASOS" Then
                CANTI = dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value
                dxDBGrid1.Dataset.Post
                SQL = "select * from if3formula where f3codpro='" & dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value & "'"
                If rsDetFormula.State = adStateOpen Then rsDetFormula.Close
                rsDetFormula.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                dxDBGrid2.Dataset.First
                Do While Not rsDetFormula.EOF
                    
                    If dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").Value = rsDetFormula.Fields("F3CODPROINS") And dxDBGrid2.Columns.ColumnByFieldName("COD_FORMULA").Value = rsDetFormula.Fields("F3CODPRO") Then
                        dxDBGrid2.Dataset.Edit
                        dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").Value = ((CANTI) * Val("" & rsDetFormula.Fields("F3CANTIDAD")) + Val("" & dxDBGrid1.Columns.ColumnByFieldName("F4VASOS").Value))
                        dxDBGrid2.Dataset.Post
                        rsDetFormula.MoveNext
                    End If
                    dxDBGrid2.Dataset.Next
                Loop
                'If dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value < dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value Then
                '    MsgBox "La Cantidad No Puedes Ser Mayor al Stock Actual", vbInformation, "Sistema de Inventarios"
                '    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").ColIndex - 1
                '    dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value = 0#
                'Else
                '    dxDBGrid1.Dataset.Post
                '    dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").ColIndex - 1
                    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
                'End If
            End If
        End If
    End If

End Sub


Private Sub dxDBGrid2_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid2.Columns.ColumnByFieldName("F3ITEM").Value = dxDBGrid2.Dataset.RecordCount + 1
            dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").Value = Format(0, "###,##0.00")
            dxDBGrid2.Columns.ColumnByFieldName("MERMA").Value = Format(0, "###,##0.00")
            dxDBGrid2.Columns.ColumnByFieldName("PRECIO").Value = Format(0, "###,##0.00")
            dxDBGrid2.Columns.ColumnByFieldName("STOCK").Value = Format(0, "###,##0.00")
            dxDBGrid2.Columns.FocusedIndex = 0
        End If
    End If

End Sub

Private Sub dxDBGrid2_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    If sw_nuevo_item = False Then
        If Action = daInsert Then
       
            If dxDBGrid2.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid2.Columns(1).Value & "")) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid2.Dataset.Refresh
        End If
    End If

End Sub

Private Sub dxDBGrid2_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    Set rsconsulta = New ADODB.Recordset
    Me.MousePointer = 11
    'FrmHelpAlmPro.Show 1
    hlp_productos.Show 1
    If Len(Trim(wcodproducto)) <> 0 Then
        dxDBGrid2.Dataset.Edit
        dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").Value = wcodproducto
        dxDBGrid2.Columns.ColumnByFieldName("DES_MAT").Value = wdesproducto
        dxDBGrid2.Columns.ColumnByFieldName("U_MEDIDA").Value = wmedida
        dxDBGrid2.Columns.ColumnByFieldName("PRECIO").Value = wvalvta
        dxDBGrid2.Columns.ColumnByFieldName("STOCK").Value = wstock
        If rsconsulta.State = adStateOpen Then rsconsulta.Close
        rsconsulta.Open "SELECT * FROM IF5PLA WHERE F5CODPRO = '" & wcodproducto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsconsulta.EOF Then
            dxDBGrid2.Columns.ColumnByFieldName("F3MONEDA").Value = "" & rsconsulta.Fields("F5MONEDA")
        End If
       
        dxDBGrid2.Dataset.Post
    End If
    Me.MousePointer = 1
    dxDBGrid2.Columns.FocusedIndex = 3
End Sub


Private Sub dxDBGrid2_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Set rsconsulta = New ADODB.Recordset

    If dxDBGrid2.Dataset.State <> 0 And dxDBGrid2.Dataset.State <> 1 Then
                
        If dxDBGrid2.Columns.FocusedColumn.FieldName = "COD_PROD" Then
            wcodproducto = dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").Value
            If Len(Trim(wcodproducto)) > 0 Then
                dxDBGrid2.Dataset.Edit
                If rsconsulta.State = adStateOpen Then rsconsulta.Close
                SQL = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB,IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO, IF6ALMA.F6STOCKACT, IF5PLA.F5VALVTA, IF5PLA.F7CODMED FROM IF5PLA INNER JOIN IF6ALMA ON IF5PLA.F5CODPRO = IF6ALMA.F5CODPRO WHERE  (IF6ALMA.F2CODALM='" & wcod_alm & "') AND (if6alma.f6stockact > 0) AND ((IF5PLA.F5CODPRO='" & wcodproducto & "') OR (IF5PLA.F5CODFAB='" & wcodproducto & "')) ORDER BY IF5PLA.F5CODPRO;"
                rsconsulta.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rsconsulta.EOF Then
                    wcodproducto = "" & Trim(rsconsulta.Fields("F5CODPRO"))
                    wcodfab = "" & rsconsulta.Fields("F5CODFAB")
                    wdesproducto = "" & rsconsulta.Fields("F5NOMPRO")
                    wmedida = "" & rsconsulta.Fields("F7CODMED")
                    gmoneda = "" & rsconsulta.Fields("F5MONEDA")
                    wstock = Val("" & rsconsulta.Fields("F6STOCKACT"))
                    wvalvta = Val("" & rsconsulta.Fields("F5VALVTA"))
                    
                Else
                    MsgBox "Código no existe", vbInformation + vbDefaultButton1, "Atención"
                    wcodproducto = "": wdesproducto = "": wmedida = "": wcodfab = ""
                    wvalvta = 0#: wstock = 0#: gmoneda = ""
                    dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").Value = wcodproducto
                    dxDBGrid2.Dataset.Post
                    Exit Sub
                End If
                dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").Value = wcodproducto
                dxDBGrid2.Columns.ColumnByFieldName("DES_MAT").Value = wdesproducto
                dxDBGrid2.Columns.ColumnByFieldName("U_MEDIDA").Value = wmedida
                dxDBGrid2.Columns.ColumnByFieldName("STOCK").Value = wstock
                dxDBGrid2.Columns.ColumnByFieldName("PRECIO").Value = wvalvta
                dxDBGrid2.Columns.ColumnByFieldName("F5MONEDA").Value = gmoneda
                dxDBGrid2.Dataset.Post
                
                dxDBGrid2.Columns.FocusedIndex = dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").ColIndex - 1
            End If
        Else
            If dxDBGrid2.Columns.FocusedColumn.FieldName = "CANTIDAD" Then
                If Val(dxDBGrid2.Columns.ColumnByFieldName("STOCK").Value) < Val(dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").Value) Then
                    MsgBox "La Cantidad No Puedes Ser Mayor al Stock Actual", vbInformation, "Sistema de Inventarios"
                    dxDBGrid2.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex - 1
                    dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").Value = 0#
                Else
                dxDBGrid2.Dataset.Post
                dxDBGrid2.Columns.FocusedIndex = dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").ColIndex - 1
                End If
            End If
        End If
    End If

End Sub


Private Sub Nuevo()
    Me.MousePointer = 11
    sw_nuevo_doc = False
    sw_detalle = False
    Txtcodalm.Text = "": PnlNomalm.Caption = ""
    Txtcodpar.Text = "": PnlNomPar.Caption = ""
    Txtcodori.Text = "": PnlNomOri.Caption = ""
    Txtnumvalo.Caption = "": Txtnumvald.Caption = ""
    
    AdicionaItem
    'AdicionaItem
    AdicionaItem2
    'AdicionaItem2
    Txtcodalm.SetFocus
    sw_nuevo_doc = True
    Me.MousePointer = 1
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
   
End Sub

Private Sub Verifica_Datos()
Dim Mensaje As String

If Trim(Txtcodalm.Text) = "" Or Trim(Txtcodpar.Text) = "" Then
    MsgBox "El Almacén de Origen y el de Destino no Pueden ser los Mismos", vbInformation, "Sistema de Inventarios"
    Exit Sub
End If

If dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").SummaryFooterValue > 0 Then 'And (sw_cabecera = True Or sw_detalle = True) Then
    If sw_nuevo_doc = False Then
    Mensaje = "La Transferencia/Transformación no ha sido grabada ... Desea Grabar ?"
    Else
    Mensaje = "Desea Grabar la Transferencia/Transformación ... ?"
    End If
Else
    Mensaje = "Debe Registrar Por lo Menos un Item para Generar un Vale"
End If
MsgBox Mensaje, vbYesNo + vbInformation, "Sistema de Inventarios"

End Sub

Sub VALIDA(valor As String)
    Dim cod As String
    Dim sqlx As String
    Dim codformu$
    Dim i%
    Dim X As Integer
    
    
    cod = valor
    
    Set rsCabFormula = New ADODB.Recordset
    Set rsDetFormula = New ADODB.Recordset
    
    If Trim(UCase(Txtcodori.Text)) = "CV1" Or Trim(UCase(Txtcodori.Text)) = "XTR" Then
        sqlx = "select * from if4formula where f4codpro='" & valor & "'"
        rsCabFormula.Open sqlx, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsCabFormula.EOF Then
            Rem EMB dxDBGrid1.Columns.ColumnByFieldName("f6total") = Format(VAL("" & rsCabFormula.Fields("F4PRECIO")), "#0.00")   'MONTO
            sqlx = "select * from if3formula where f3codpro='" & valor & "'"
            rsDetFormula.Open sqlx, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            codformu$ = rsCabFormula.Fields("F4CODPRO")
            'HACIENDO EL LLENADO
            If Not rsDetFormula.EOF Then
                flag = False
                If rsDetFormula.RecordCount = 0 Then
                    'rsDetFormula.Recordset.MoveFirst
                Else
                    'rsDetFormula.Refresh
                    'rsDetFormula.Recordset.MoveFirst
                    rsDetFormula.MoveFirst
                    For i% = 0 To rsDetFormula.RecordCount - 1
                        If rsDetFormula.Fields("F3CODPRO") = rsCabFormula.Fields("F4CODPRO") Then
                            flag = True
                        End If
                        rsDetFormula.MoveNext
                    Next i%
                End If
                X = dxDBGrid1.Dataset.RecordCount
                If dxDBGrid1.Dataset.RecordCount >= 1 Then X = X + 1
                dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Value = X
                dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = "" & rsCabFormula.Fields("F4CODPRO")
                dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = "" & rsCabFormula.Fields("F4NOMPRO")
                dxDBGrid1.Columns.ColumnByFieldName("F6VALPRO").Value = "" & rsCabFormula.Fields("F4PRECIO")
                dxDBGrid1.Columns.ColumnByFieldName("F7SIGMED").Value = "" & rsCabFormula.Fields("F4MEDIDA")
                dxDBGrid1.Columns.ColumnByFieldName("F6CANMOV").Value = rsCabFormula.Fields("F4CANTIDAD")
                dxDBGrid1.Columns.ColumnByFieldName("F6TOTAL").Value = 0#
                dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = 0#
                dxDBGrid1.Columns.ColumnByFieldName("F3MONEDA").Value = ""
                'X = 1
                flag = False
                Do While Not rsDetFormula.EOF

                    If rsCabFormula.Fields("F4CODPRO") = rsDetFormula.Fields("F3CODPRO") Then
                        'If dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = rsDetFormula.Fields("F3CODPRO") Then
                        If flag = False Then
                            If dxDBGrid2.Dataset.RecordCount > 1 Then
                                dxDBGrid2.Dataset.Append
                            Else
                                dxDBGrid2.Dataset.Edit
                            End If
                            flag = True
                        Else
                            dxDBGrid2.Dataset.Append
                            'X = X + 1
                        End If
                        dxDBGrid2.Columns.ColumnByFieldName("F3ITEM").Value = rsDetFormula.Fields("F3ITEM")
                        dxDBGrid2.Columns.ColumnByFieldName("COD_FORMULA").Value = "" & rsDetFormula.Fields("F3CODPRO")
                        dxDBGrid2.Columns.ColumnByFieldName("COD_PROD").Value = "" & rsDetFormula.Fields("F3CODPROINS")
                        dxDBGrid2.Columns.ColumnByFieldName("DES_MAT").Value = "" & rsDetFormula.Fields("F3NOMPRO")
                        dxDBGrid2.Columns.ColumnByFieldName("U_MEDIDA").Value = "" & rsDetFormula.Fields("F3UNIDAD")
                        dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").Value = rsDetFormula.Fields("F3CANTIDAD")
                        dxDBGrid2.Columns.ColumnByFieldName("MERMA").Value = 0#
                        dxDBGrid2.Columns.ColumnByFieldName("PRECIO").Value = 0#
                        dxDBGrid2.Columns.ColumnByFieldName("STOCK").Value = 0#
                        dxDBGrid2.Columns.ColumnByFieldName("F3MONEDA").Value = ""
                        'dxDBGrid2.Columns(5) = "" & rsDetFormula.Fields("MERMA")
                        dxDBGrid2.Dataset.Post
                    End If
                    If rsDetFormula.EOF Then
                        Exit Do
                    Else
                        rsDetFormula.MoveNext
                    End If
                Loop
                'dxDBGrid1.Refresh
                'rsDetFormula.Update
            End If
        Else
'          MsgBox "El Producto no tiene Formula"
'          dxDBGrid1.Enabled = True
'          DataDetalle.UpdateRecord
'          CodTrans% = 1
'          frmformula.Show 1
'          If CodTrans = 0 Then
'            rsCabFormula.MoveFirst
'            rsDetFormula.MoveFirst
'            rsCabFormula.Seek "=", Codigo
'            If Not rsCabFormula.EOF Then
'                rsDetFormula.Seek "=", rsCabFormula.Fields("F3CODPRO")
'                'HACIENDO EL LLENADO
'                If Not rsDetFormula.NoMatch Then
'    '              Do While rsCabFormula.Fields("F3CODPRO") = "" & rsDetFormula.Fields("F3CODPRO")
'                   Do While Not rsDetFormula.EOF
'                    If rsCabFormula.Fields("F3CODPRO") = "" & rsDetFormula.Fields("F3CODPRO") Then
'                        dxDBGrid1.Columns(0) = "" & rsDetFormula.Fields("F3CODPRO")
'                        dxDBGrid1.Columns(1) = "" & rsDetFormula.Fields("COD_PROD")
'                        dxDBGrid1.Columns(2) = "" & rsDetFormula.Fields("DES_MAT")
'                        dxDBGrid1.Columns(3) = "" & rsDetFormula.Fields("U_MEDIDA")
'                        dxDBGrid1.Columns(4) = "" & rsDetFormula.Fields("CANTIDAD") * (dxDBGrid1.Columns(3).Text)
'                        dxDBGrid1.Columns(5) = "" & rsDetFormula.Fields("MERMA")
'                        dxDBGrid1.MoveNext
'                        dxDBGrid1.Row = dxDBGrid1.Row + 1
'                        rsDetFormula.UpdateRecord
'                    End If
'                    If rsDetFormula.EOF Then
'                        Exit Do
'                    Else
'                        rsDetFormula.MoveNext
'                    End If
'                  Loop
'                  dxDBGrid1.Refresh
'                  rsDetFormula.Refresh
'                  dxDBGrid1.Enabled = False
'                End If
'            End If
'          Else
'
'            dxDBGrid1.Row = dxDBGrid1.Row
'            dxDBGrid1.col = 0
'            dxDBGrid1.SetFocus
'          End If
        End If
    Else
'        rsDetFormula.DatabaseName = rutatem & "TEMFORMU.MDB"
'        rsDetFormula.RecordSource = "TEMFORM"
'        rsDetFormula.Refresh
'        rsDetFormula.Recordset.AddNew
'        rsDetFormula.Recordset.Fields("COD_PROD") = Valcod
'        rsDetFormula.Recordset.Fields("DES_MAT") = Valpro
'        rsDetFormula.Recordset.Fields("U_MEDIDA") = valuni
'        rsDetFormula.Recordset.Fields("CANTIDAD") = valcan
'        rsDetFormula.Recordset.Update
'        dxDBGrid1.Refresh
    End If
End Sub

Private Sub Graba_Datos()

    
 '''''''''''''' ACTUALIZANDO LOS VALES EN EL EF2ALMACENES ''''''''''
    wmes = Format(Month(CVDate(TxtFecMov.Value)), "00")
    Txtnumvalo.Caption = Calcula_Numero(Trim(wcod_alm), "S")
    Txtnumvald.Caption = Calcula_Numero(Trim(WcodPar), "I")
    
    ACTUALIZAR
    ACTUALIZAR1
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
''''''''''''''''' GRABANDO LA SALIDA DE ALMACEN ''''''''''''''''''
    GRABAR_SALIDA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''GRABANDO INGRESO AL ALMACEN '''''''''''''''''''''''
    GRABAR_INGRESO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
 
End Sub

Private Sub ACTUALIZAR()

amovs_vale(0).campo = "F1VALSAL" & Format(wmes, "00"): amovs_vale(0).valor = Txtnumvalo.Caption: amovs_vale(0).TIPO = "T"

GRABA_REGISTRO amovs_vale(), "EF2ALMACENES", "M", 0, cnn_dbbancos, "F2CODALM ='" & Trim(wcod_alm) & "'"

End Sub


Private Sub GRABAR_SALIDA()
Dim intcont, nitems, nfila As Integer
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

BASE_TEMPORAL "TEMFORMU.MDB"

'---------------------ASIGNA DATOS A IF4VALES -------------------------------------

amovs_cab(0).campo = "F2CODALM": amovs_cab(0).valor = wcod_alm: amovs_cab(0).TIPO = "T"
amovs_cab(1).campo = "F4NUMVAL": amovs_cab(1).valor = Txtnumvalo.Caption: amovs_cab(1).TIPO = "T"
amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = TxtFecMov.Value: amovs_cab(2).TIPO = "F"
amovs_cab(3).campo = "F2CODPAR": amovs_cab(3).valor = WcodPar: amovs_cab(3).TIPO = "T"
amovs_cab(4).campo = "F2CODPROV": amovs_cab(4).valor = "0": amovs_cab(4).TIPO = "T"
amovs_cab(5).campo = "F1CODORI": amovs_cab(5).valor = wcodori: amovs_cab(5).TIPO = "T"
amovs_cab(6).campo = "F1CODDOC": amovs_cab(6).valor = "03": amovs_cab(6).TIPO = "T"
amovs_cab(7).campo = "F4NUMDOC": amovs_cab(7).valor = Txtnumvald.Caption: amovs_cab(7).TIPO = "T"
amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = gmoneda: amovs_cab(8).TIPO = "T"
amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = wtipcam: amovs_cab(9).TIPO = "T"
amovs_cab(10).campo = "F4FECULT": amovs_cab(10).valor = Format(Now, "DD/MM/YYYY"): amovs_cab(10).TIPO = "F"
amovs_cab(11).campo = "F2CODUSE": amovs_cab(11).valor = wusuario: amovs_cab(11).TIPO = "T"
    

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
SQL = "Select count(F3ITEM) as NITEM from Temform Where LEN(TRIM(F3ITEM))> 0 "
rsconsulta.Open SQL, Temp, adOpenDynamic, adLockOptimistic
If Not rsconsulta.EOF Then
nitems = Val("" & rsconsulta.Fields("NITEM"))
End If
rsconsulta.Close

ReDim Values(8, nitems)

If rsconsulta.State = adStateOpen Then rsconsulta.Close
rsconsulta.Open "Select * from Temform", Temp
If Not rsconsulta.EOF Then
nfila = 0

rsconsulta.MoveFirst
Do While Not rsconsulta.EOF
    If Len(Trim(rsconsulta.Fields("F3ITEM") & "")) > 0 Then
        Values(0, nfila) = "" & Txtnumvalo.Caption
        Values(1, nfila) = "" & rsconsulta.Fields("cod_prod")
        Values(2, nfila) = "" & rsconsulta.Fields("cantidad")
        If rsconsulta.Fields("F3MONEDA") = "S" Then
            Values(3, nfila) = Val("" & Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")), "0.00"))
        Else
            Values(3, nfila) = Val("" & Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")) * wtipcam, "0.00"))
        End If
        Values(4, nfila) = "" & wcod_alm
        Values(5, nfila) = "" & CVDate(TxtFecMov.Value)
        If rsconsulta.Fields("F3MONEDA") = "D" Then
            Values(6, nfila) = Val("" & Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")), "0.00"))
        Else
            Values(6, nfila) = Val("" & Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")) / wtipcam, "0.00"))
        End If
        Values(7, nfila) = Val("" & Format(Val(rsconsulta.Fields("cantidad") * Values(3, nfila)), "0.00"))
        Values(8, nfila) = Val("" & Format(Val(rsconsulta.Fields("cantidad") * Values(6, nfila)), "0.00"))
        
        If rsconsulta.Fields("F3MONEDA") = "S" Then
            Vales_Detalle Txtnumvalo.Caption, rsconsulta.Fields("cod_prod"), rsconsulta.Fields("cantidad"), (Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")), "0.00")), Txtcodalm.Text, CVDate(TxtFecMov.Value), (Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")) / wtipcam, "0.00"))
        Else
            Vales_Detalle Txtnumvalo.Caption, rsconsulta.Fields("cod_prod"), rsconsulta.Fields("cantidad"), (Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")) * wtipcam, "0.00")), Txtcodalm.Text, CVDate(TxtFecMov.Value), (Format(Val(rsconsulta.Fields("cantidad") * rsconsulta.Fields("precio")), "0.00"))
        End If
        
        nfila = nfila + 1
    End If
    rsconsulta.MoveNext
Loop
End If
rsconsulta.Close
sw_graba_registro = True

If ctipo = "A" Then '---Nuevo
    '-----Graba Cabecera
    GRABA_REGISTRO amovs_cab(), "IF4VALES", "A", 11, cnn_dbbancos, ""
    
    If sw_graba_registro = True Then
        '------- GRABA DETALLE
        GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "", Values(), nfila - 1, "11111111", "", ""
    End If

Else    '----------Modificacion
    '------- GRABA CABECERA
    GRABA_REGISTRO amovs_cab(), "IF4VALES", "M", 11, cnn_dbbancos, "F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'"
    
    '------- GRABA DETALLE
    cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'")
    GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "F4NUMVAL  = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'", Values(), nfila - 1, "111111111111", "", ""
End If

End Sub


Private Sub ACTUALIZAR1()

amovs_vale(0).campo = "F1VALING" & Format(wmes, "00"): amovs_vale(0).valor = Txtnumvald.Caption: amovs_vale(0).TIPO = "T"

GRABA_REGISTRO amovs_vale(), "EF2ALMACENES", "M", 0, cnn_dbbancos, "F2CODALM ='" & Trim(WcodPar) & "'"

End Sub

Private Sub GRABAR_INGRESO()
Dim intcont, nitems, nfila  As Integer
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

BASE_TEMPORAL "TEMPFAC.MDB"


'---------------------ASIGNA DATOS A IF4VALES -------------------------------------

amovs_cab(0).campo = "F2CODALM": amovs_cab(0).valor = WcodPar: amovs_cab(0).TIPO = "T"
amovs_cab(1).campo = "F4NUMVAL": amovs_cab(1).valor = Txtnumvald.Caption: amovs_cab(1).TIPO = "T"
amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = TxtFecMov.Value: amovs_cab(2).TIPO = "F"
amovs_cab(3).campo = "F2CODPAR": amovs_cab(3).valor = wcod_alm: amovs_cab(3).TIPO = "T"
amovs_cab(4).campo = "F2CODPROV": amovs_cab(4).valor = "0": amovs_cab(4).TIPO = "T"
amovs_cab(5).campo = "F1CODORI": amovs_cab(5).valor = wcodori: amovs_cab(5).TIPO = "T"
amovs_cab(6).campo = "F1CODDOC": amovs_cab(6).valor = "03": amovs_cab(6).TIPO = "T"
amovs_cab(7).campo = "F4NUMDOC": amovs_cab(7).valor = Txtnumvalo.Caption: amovs_cab(7).TIPO = "T"
amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = gmoneda: amovs_cab(8).TIPO = "T"
amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = wtipcam: amovs_cab(9).TIPO = "T"
amovs_cab(10).campo = "F4FECULT": amovs_cab(10).valor = Format(Now, "DD/MM/YYYY"): amovs_cab(10).TIPO = "F"
amovs_cab(11).campo = "F2CODUSE": amovs_cab(11).valor = wusuario: amovs_cab(11).TIPO = "T"
    

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
    If Len(Trim(rsconsulta.Fields("F3ITEM") & "")) > 0 Then
        Values(0, nfila) = "" & Txtnumvald.Caption
        Values(1, nfila) = "" & rsconsulta.Fields("F5CODPRO")
        Values(2, nfila) = "" & rsconsulta.Fields("F6CANMOV")
        If rsconsulta.Fields("F3MONEDA") = "S" Then
            Values(3, nfila) = Val("" & Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")), "0.00"))
        Else
            Values(3, nfila) = Val("" & Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")) * wtipcam, "0.00"))
        End If
        Values(4, nfila) = "" & WcodPar
        Values(5, nfila) = "" & CVDate(TxtFecMov.Value)
        If rsconsulta.Fields("F3MONEDA") = "D" Then
            Values(6, nfila) = Val("" & Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")), "0.00"))
        Else
            Values(6, nfila) = Val("" & Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")) / wtipcam, "0.00"))
        End If
        Values(7, nfila) = Val("" & Format(Val(rsconsulta.Fields("F6CANMOV") * Values(3, nfila)), "0.00"))
        Values(8, nfila) = Val("" & Format(Val(rsconsulta.Fields("F6CANMOV") * rsconsulta.Fields("F6TOTAL")), "0.00"))
        
        If rsconsulta.Fields("F3MONEDA") = "S" Then
            Vales_Detalle Txtnumvald.Caption, rsconsulta.Fields("F5CODPRO"), rsconsulta.Fields("F6CANMOV"), Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")), "0.00"), Txtcodpar.Text, CVDate(TxtFecMov.Value), Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")) / wtipcam, "0.00")
        Else
            Vales_Detalle Txtnumvald.Caption, rsconsulta.Fields("F5CODPRO"), rsconsulta.Fields("F6CANMOV"), Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")) * wtipcam, "0.00"), Txtcodpar.Text, CVDate(TxtFecMov.Value), Format(Val(rsconsulta.Fields("F6TOTAL") * rsconsulta.Fields("F6CANMOV")), "0.00")
        End If
        nfila = nfila + 1
    End If
    rsconsulta.MoveNext
Loop
End If
rsconsulta.Close
sw_graba_registro = True

If ctipo = "A" Then '---Nuevo
    '-----Graba Cabecera
    GRABA_REGISTRO amovs_cab(), "IF4VALES", "A", 11, cnn_dbbancos, ""
    
    If sw_graba_registro = True Then
        '------- GRABA DETALLE
        GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "", Values(), nfila - 1, "111111111111", "", ""
    End If

Else    '----------Modificacion
    '------- GRABA CABECERA
    GRABA_REGISTRO amovs_cab(), "IF4VALES", "M", 11, cnn_dbbancos, "F4NUMVAL = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'"
    
    '------- GRABA DETALLE
    cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & Txtnumvalo.Caption & "' AND F2CODALM = '" & Txtcodalm.Text & "'")
    GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 8, cnn_dbbancos, "F4NUMVAL  = '" & Txtnumvald.Caption & "' AND F2CODALM = '" & Txtcodpar.Text & "'", Values(), nfila - 1, "111111111111", "", ""
End If

End Sub

Private Sub Imprimir_Datos()

   ImprimendoV Trim(Txtcodalm.Text), Trim(Txtnumvalo.Caption), 0
   ImprimendoV Trim(Txtcodpar.Text), Trim(Txtnumvald.Caption), 0
   SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False


End Sub

Private Sub Buscando()

Set RsStockCab = New ADODB.Recordset

    Gtipval = "3"
    wcod_alm = Txtcodalm.Text
    wnumval = Txtnumvalo.Caption
    
    FrmAyudaVale.Show 1
    Txtcodalm.Text = wcod_alm
    Txtnumvalo.Caption = wnumval
    
    If RsStockCab.State = adStateOpen Then RsStockCab.Close
    RsStockCab.Open "SELECT * FROM IF4VALES WHERE F2CODALM='" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockCab.EOF Then
        Actualiza_Datos
    End If

End Sub

Private Sub Actualiza_Datos()
Dim csql, SSQL As String
Dim CONT As Integer
Set RsStockDet = New ADODB.Recordset
Set RsProducto = New ADODB.Recordset

Txtnumvalo.Caption = "" & RsStockCab.Fields("f4numval")
TxtFecMov.Value = Format(RsStockCab.Fields("F4Fecval"), "dd/mm/yyyy")
Txtcodpar.Text = "" & RsStockCab.Fields("F2CODPAR")
If VALIDA_ALMACENO(Txtcodpar.Text) = True Then
    WcodPar = Trim(Txtcodpar.Text)
    PnlNomPar.Caption = WNomPar
    SSFrame1.Caption = "Productos a Ingresar al  " & PnlNomPar.Caption
Else
    MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
    Txtcodpar.SetFocus
End If
Txtcodalm.Text = "" & RsStockCab.Fields("F2CODALM")
If VALIDA_ALMACEN(Txtcodalm.Text) = True Then
    wcod_alm = Txtcodalm.Text
    PnlNomalm.Caption = wnomalmacen
    SSFrame2.Caption = "Insumos de Productos a Transferir del " & Trim(PnlNomalm.Caption)
Else
    MsgBox "Código de almacén no existe. Verifique.", vbInformation, "Atención"
    Txtcodalm.SetFocus
End If

Txtnumvald.Caption = "" & RsStockCab.Fields("F4NUMDOC")
Txtcodori.Text = "" & RsStockCab.Fields("f1CODORI")

'-----------PARA GRID1 ---------------

BASE_TEMPORAL "TEMPFAC.MDB"
DELETEREC_N "TmpStock", Temp
dxDBGrid1.Dataset.Refresh
Conf_Grid

'---------PARA GRID2 -----------------

BASE_TEMPORAL "TEMFORMU.MDB"
DELETEREC_N "TEMFORM", Temp
dxDBGrid2.Dataset.Refresh
Conf_Grid2

'-----------------------------------------
   
'--------PARA VER EL INGRESO AL ALMACEN -----
If RsStockDet.State = adStateOpen Then RsStockDet.Close
SQL = "SELECT * FROM IF3VALES WHERE F2CODALM= '" & RsStockCab.Fields("F2CODPAR") & "' AND F4NUMVAL = '" & RsStockCab.Fields("F4NUMDOC") & "'"
RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockDet.EOF Then
    RsStockDet.MoveFirst
    CONT = 0
    Do While Not RsStockDet.EOF
        
        dxDBGrid1.Dataset.Append
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns(0).Value = "" & RsStockDet.Fields("F5CODPRO")
        dxDBGrid1.Columns(4).Value = "" & Format(RsStockDet.Fields("F3CANPRO"), "0.00")
        dxDBGrid1.Columns(7).Value = CONT
              
        SSQL = "Select * from IF5PLA where F5CODPRO='" & RsStockDet.Fields("F5CODPRO") & "'"
        If RsProducto.State = adStateOpen Then RsProducto.Close
        RsProducto.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsProducto.EOF Then
            dxDBGrid1.Columns(1).Value = "" & RsProducto.Fields("F5NOMPRO")
            dxDBGrid1.Columns(2).Value = "" & Format(RsProducto.Fields("F5STOCKACT"), "0.00")
            dxDBGrid1.Columns(6).Value = "" & RsProducto.Fields("F5CODFAB")
            dxDBGrid1.Columns(8).Value = "" & RsProducto.Fields("F5MONEDA")
            dxDBGrid1.Columns(3).Value = "" & RsProducto.Fields("F7CODMED")
            If RsProducto.Fields("F5MONEDA") = "S" Then
                dxDBGrid1.Columns(5).Value = "" & Format(RsStockDet.Fields("F3VALVTA"), "0.00")
            Else
                dxDBGrid1.Columns(5).Value = "" & Format(RsStockDet.Fields("F3VALDOL"), "0.00")
            End If
        End If
                    
        CONT = CONT + 1
        RsStockDet.MoveNext
        If RsStockDet.EOF Then Exit Do
        
    Loop
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Dataset.Post
End If
dxDBGrid1.Dataset.First
dxDBGrid1.Columns.FocusedIndex = 0

'------------------ PARA VER LAS SALIDAS ----------------------

If RsStockDet.State = adStateOpen Then RsStockDet.Close
SQL = "SELECT * FROM IF3VALES WHERE F2CODALM= '" & RsStockCab.Fields("F2CODALM") & "' AND F4NUMVAL = '" & RsStockCab.Fields("F4NUMVAL") & "'"
RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockDet.EOF Then
    RsStockDet.MoveFirst
    CONT = 0
    Do While Not RsStockDet.EOF
        
        dxDBGrid2.Dataset.Append
        dxDBGrid2.Dataset.Edit
        dxDBGrid2.Columns(0).Value = ""
        dxDBGrid2.Columns(1).Value = "" & RsStockDet.Fields("F5CODPRO")
        dxDBGrid2.Columns(4).Value = "" & Format(Val(RsStockDet.Fields("F3CANPRO") & ""), "0.00")
        dxDBGrid2.Columns(5).Value = "" & Format(Val(RsStockDet.Fields("F3MERMA") & ""), "0.00")
        dxDBGrid2.Columns(7).Value = CONT
              
        SSQL = "Select * from IF5PLA where F5CODPRO='" & RsStockDet.Fields("F5CODPRO") & "'"
        If RsProducto.State = adStateOpen Then RsProducto.Close
        RsProducto.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsProducto.EOF Then
            dxDBGrid2.Columns(2).Value = "" & RsProducto.Fields("F5NOMPRO")
            dxDBGrid2.Columns(8).Value = "" & Format(RsProducto.Fields("F5STOCKACT"), "0.00")
            dxDBGrid2.Columns(9).Value = "" & RsProducto.Fields("F5MONEDA")
            dxDBGrid2.Columns(3).Value = "" & RsProducto.Fields("F7CODMED")
            If RsProducto.Fields("F5MONEDA") = "S" Then
                dxDBGrid2.Columns(6).Value = "" & Format(RsStockDet.Fields("F3VALVTA"), "0.00")
            Else
                dxDBGrid2.Columns(6).Value = "" & Format(RsStockDet.Fields("F3VALDOL"), "0.00")
            End If
        End If
                    
        CONT = CONT + 1
        RsStockDet.MoveNext
        If RsStockDet.EOF Then Exit Do
        
    Loop
    dxDBGrid2.Dataset.Edit
    dxDBGrid2.Dataset.Post
End If
dxDBGrid2.Dataset.First
dxDBGrid2.Columns.FocusedIndex = 0

End Sub

Private Sub Elimina_Movimientos()
Dim SSQL, csql As String
Set RsStockCab = New ADODB.Recordset
Set RsStockDet = New ADODB.Recordset

If RsStockCab.State = adStateOpen Then RsStockCab.Close
SQL = "SELECT * FROM IF4VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
RsStockCab.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockCab.EOF Then
    If MsgBox("Está seguro de eliminar los movimientos registrados", vbInformation + vbYesNo, "Sistema de Inventarios") = vbYes Then
        '-----------CABECERA ----------
        SSQL = "DELETE FROM IF4VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        cnn_dbbancos.Execute (SSQL)
        
        '----------DETALLE ---------
        If RsStockDet.State = adStateOpen Then RsStockDet.Close
        SQL = "SELECT * FROM IF3VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsStockDet.EOF Then
            Do While Not RsStockDet.EOF
                Reactualiza_Almacenes Trim(RsStockDet.Fields("F2codalm")), Trim(RsStockDet.Fields("F5codpro")), Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), CVDate(RsStockDet.Fields("F4fecval")), Val(Format(RsStockDet.Fields("F3totite"), "#0.000")), Val(Format(RsStockDet.Fields("F3totdol"), "#0.000")), "I", Val(Format(RsStockDet.Fields("F3valdol"), "#0.000"))
                RsStockDet.MoveNext
                If RsStockDet.EOF Then Exit Do
            Loop
        End If
        
        csql = "DELETE FROM IF3VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        cnn_dbbancos.Execute (csql)
        
        SSQL = "Delete From IF3VALES where F2codalm='" & Txtcodpar.Text & "' and F4numval='" & Txtnumvald.Caption & "'"
        cnn_dbbancos.Execute (SSQL)
        '---------------------------------------------------
        
        If RsStockCab.State = adStateOpen Then RsStockCab.Close
        SQL = "SELECT * FROM IF4VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
        RsStockCab.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsStockCab.EOF Then
            
            '-----------CABECERA ----------
            SSQL = "DELETE FROM IF4VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
            cnn_dbbancos.Execute (SSQL)
            
            '----------DETALLE ---------
            If RsStockDet.State = adStateOpen Then RsStockDet.Close
            SQL = "SELECT * FROM IF3VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
            RsStockDet.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsStockDet.EOF Then
                Do While Not RsStockDet.EOF
                    Reactualiza_Almacenes Trim(RsStockDet.Fields("F2codalm")), Trim(RsStockDet.Fields("F5codpro")), Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), CVDate(RsStockDet.Fields("F4fecval")), Val(Format(RsStockDet.Fields("F3totite"), "#0.000")), Val(Format(RsStockDet.Fields("F3totdol"), "#0.000")), "S", Val(Format(RsStockDet.Fields("F3valdol"), "#0.000"))
                    RsStockDet.MoveNext
                    If RsStockDet.EOF Then Exit Do
                Loop
            End If
            
            csql = "DELETE FROM IF3VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
            cnn_dbbancos.Execute (csql)
            
            SSQL = "Delete From IF3VALES where F2codalm='" & Txtcodalm.Text & "' and F4numval='" & Txtnumvalo.Caption & "'"
            cnn_dbbancos.Execute (SSQL)
            '---------------------------------------------------
            Nuevo
            Txtcodalm.SetFocus
        End If
    End If
Else
    MsgBox "El Registro no ha sido Grabado", vbInformation, "Sistema de Inventarios"
    Exit Sub
End If
End Sub

Private Sub RENUMERARITEMS()
Dim i As Integer

    sw_nuevo_item = True
    dxDBGrid1.Dataset.First
    Do While Not dxDBGrid1.Dataset.EOF
        i = i + 1
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = i
        dxDBGrid1.Dataset.Next
    Loop
    'dxDBGrid1.Dataset.Post
    sw_nuevo_item = False

End Sub

Private Sub RENUMERARITEMS2()
Dim i As Integer

    sw_nuevo_item = True
    dxDBGrid2.Dataset.First
    Do While Not dxDBGrid2.Dataset.EOF
        i = i + 1
        dxDBGrid2.Dataset.Edit
        dxDBGrid2.Columns.ColumnByFieldName("ITEM").Value = i
        dxDBGrid2.Dataset.Next
    Loop
    'dxDBGrid2.Dataset.Post
    sw_nuevo_item = False

End Sub


