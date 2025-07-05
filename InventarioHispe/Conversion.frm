VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Conversion 
   Caption         =   "Conversión"
   ClientHeight    =   6915
   ClientLeft      =   2190
   ClientTop       =   2085
   ClientWidth     =   10665
   Icon            =   "Conversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10665
   Begin Threed.SSPanel SSPanel2 
      Height          =   6630
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   10500
      _Version        =   65536
      _ExtentX        =   18521
      _ExtentY        =   11695
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
      Begin VB.TextBox Txtcodalm 
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
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   0
         Top             =   105
         Width           =   465
      End
      Begin VB.TextBox Txtcodpar 
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
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   1
         Top             =   405
         Width           =   465
      End
      Begin VB.TextBox Txtcodori 
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
         Height          =   285
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   2
         Top             =   720
         Width           =   465
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2145
         Left            =   135
         TabIndex        =   4
         Top             =   1350
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   3784
         _StockProps     =   14
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
            Height          =   1770
            Left            =   135
            OleObjectBlob   =   "Conversion.frx":058A
            TabIndex        =   5
            Top             =   225
            Width           =   9900
         End
      End
      Begin Threed.SSPanel Txtnumvalo 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BevelOuter      =   1
      End
      Begin Threed.SSPanel PnlNomAlm 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   105
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         TabIndex        =   8
         Top             =   720
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         TabIndex        =   9
         Top             =   405
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   2460
         Left            =   135
         TabIndex        =   15
         Top             =   3915
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   4339
         _StockProps     =   14
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
            Height          =   2490
            Left            =   135
            OleObjectBlob   =   "Conversion.frx":343B
            TabIndex        =   16
            Top             =   225
            Width           =   9900
         End
      End
      Begin Threed.SSPanel Txtnumvald 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   3600
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BevelOuter      =   1
      End
      Begin MSComCtl2.DTPicker TxtFecMov 
         Height          =   315
         Left            =   8400
         TabIndex        =   19
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   114425857
         CurrentDate     =   40611
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vale Ingreso:"
         DataField       =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vale Salida:"
         DataField       =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1155
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   7560
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén Origen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   285
         TabIndex        =   12
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label LabPar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén Destino:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   285
         TabIndex        =   11
         Top             =   465
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   285
         TabIndex        =   10
         Top             =   810
         Width           =   735
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   15
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "Conversion.frx":62F2
      ToolBars        =   "Conversion.frx":AEFF
   End
End
Attribute VB_Name = "Conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------
Dim It_em_form                       As Integer
Dim sw_mensaje                    As Boolean
Dim obtiene_codigo As String
Dim X As Integer
Dim wprodcompara As String
Dim sw_ayuda As Boolean
Dim sw_ayudaO As Boolean
Dim sw_nuevo_item, sw_detalle, sw_cabecera As Boolean
Dim Temp As ADODB.Connection
Dim Temp1 As ADODB.Connection
Dim Ubica     As Integer
Dim amovs_vale(0 To 0)  As a_grabacion
Dim amovs_cab(0 To 11) As a_grabacion
Dim amovs_det(0 To 10) As a_grabacion
Dim sw_new_item As Boolean
Dim flag    As Boolean
Dim CANTI   As Double
Dim Cantidad As Double
Dim wfactor As Double

Function Costo_Unitario(Codigo As String, palmacen As String) As Double
Dim CosUni As ADODB.Recordset
Dim sql As String
    
    sql = ""
    sql = sql + "SELECT IF3VALES.F5CODPRO, Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1)= 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, [VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
    sql = sql + "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F2CODALM = IF3VALES.F2CODALM) AND (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) "
    sql = sql + "Where IF3VALES.F4FECVAL <= CVDATE('" & Format(TxtFecMov.value, "DD/MM/YYYY") & "') And IF3VALES.F5CODPRO = '" & Codigo & "' AND "
    sql = sql + "IF4VALES.F2CODALM ='" & palmacen & "' "
    sql = sql + "GROUP BY IF3VALES.F5CODPRO;"
    
    Set CosUni = New ADODB.Recordset
    CosUni.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not CosUni.EOF Then
        Costo_Unitario = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
    End If
    CosUni.Close
        
End Function
 
Private Function Calcula_Numero(pcodalm As String, ptipo As String)
Dim WCONT   As String
Dim WNUMERO As String
    
    Set RsAlmacenes = New ADODB.Recordset
    sql = "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & pcodalm & "'"
    If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
    RsAlmacenes.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsAlmacenes.EOF Then
        wmes = Format(Month(CVDate(TxtFecMov.value)), "00")
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
    RsAlmacenes.Close
    Calcula_Numero = WNUMERO
    
End Function

Private Sub dxDBGrid1_LostFocus()
If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
        dxDBGrid1.Dataset.Post
        dxDBGrid1.Dataset.Refresh
 End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    If KeyCode = 115 Then
        If MsgBox("Desea Eliminar el registro Actual ", vbQuestion + vbYesNo, "Inventario") = vbYes Then
            If dxDBGrid1.Dataset.RecNo = 1 Then
               dxDBGrid1.Dataset.Delete
               AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
                RENUMERARITEMS dxDBGrid1
            End If
        End If
    End If
    
End Sub

Private Sub dxDBGrid2_LostFocus()
 If dxDBGrid2.Dataset.State = dsEdit Or dxDBGrid2.Dataset.State = dsInsert Then
        dxDBGrid2.Dataset.Post
        dxDBGrid2.Dataset.Refresh
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
                RENUMERARITEMS dxDBGrid2
            End If
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Height = 7550
    'Me.Width = 10530
    Me.Width = 10990
    Me.left = 1500
    Me.top = 1050
    
    X = 0
    sw_nuevo_doc = True
    sw_detalle = False
    sw_nuevo_item = False
    sw_new_item = False
    wprodcompara = ""
    
    sw_ayudaO = False
        
    BASE_TEMPORAL "TEMPLUS.MDB"
    DELETEREC_LOG "DETCONVERSION", Temp
    DELETEREC_LOG "DETFORMULA", Temp
    
    dxDBGrid1.Dataset.Refresh
    dxDBGrid2.Dataset.Refresh
    Conf_Grid dxDBGrid1
    Conf_Grid dxDBGrid2
    
    dxDBGrid1.Columns.ColumnByFieldName("Stock").Visible = False
    dxDBGrid1.Columns.ColumnByFieldName("COSTO").Visible = False
    dxDBGrid1.Columns.ColumnByFieldName("CANTOTAL").Visible = False
    
    dxDBGrid2.Columns.ColumnByFieldName("Stock").Visible = False
    dxDBGrid2.Columns.ColumnByFieldName("COSTO").Visible = False
    dxDBGrid2.Columns.ColumnByFieldName("CANTOTAL").Visible = False
    
    AdicionaItem
    AdicionaItem2
    
    wtipcam = gtipcam
    If wtipcam = 0# Then wtipcam = 2.65
    
    TxtFecMov.value = Format(Date, "dd/mm/yyyy")
    SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
    
End Sub

Public Sub BASE_TEMPORAL(Base As String)
Dim CON As String
    
    Set Temp = New ADODB.Connection
    If Temp.State = adStateOpen Then Temp.Close
    CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & Base & "; Persist Security Info=False"
    Temp.Open CON

End Sub

Private Sub Conf_Grid(pgrid As Control)
    
    With pgrid.Options
        .Set (egoAutoExpandOnSearch)
        .Set (egoAutoSort)
        '.Set (egoAutoWidth)
        .Set (egoBandHeaderWidth)
        .Set (egoBandMoving)
        .Set (egoBandSizing)
        .Set (egoCanAppend)
        .Set (egoCancelOnExit)
        .Set (egoCanDelete)
        .Set (egoCanInsert)
        .Set (egoCanNavigation)
        .Set (egoColumnMoving)
        .Set (egoColumnSizing)
        .Set (egoConfirmDelete)
        .Set (egoDragScroll)
        '.Set (egoDynamicLoad)
        .Set (egoEditing)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoExactScrollBar)
        .Set (egoExpandOnDblClick)
        .Set (egoHorzThrough)
        .Set (egoImmediateEditor)
        .Set (egoLoadAllRecords)
        .Set (egoNameCaseInsensitive)
        '.Set (egoShowBands)
        .Set (egoShowBorder)
        .Set (egoShowButtonAlways)
        .Set (egoShowButtons)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowHeader)
        .Set (egoShowHourGlass)
        .Set (egoShowIndicator)
        .Set (egoShowPreviewGrid)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoTabThrough)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoVertThrough)

    End With
     
End Sub

Private Sub AdicionaItem()
Dim i As Integer
Dim sw_nuevo_temp   As Boolean

    dxDBGrid1.Dataset.Active = False
    BASE_TEMPORAL "TEMPLUS.MDB"
    DELETEREC_LOG "DETCONVERSION", Temp
    dxDBGrid1.Dataset.Refresh
    
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
            .FieldValues("CODIGO") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UM") = ""
            .FieldValues("CANTIDAD") = Format(0, "###,##0.00")
            .FieldValues("STOCK") = Format(0, "###,##0.00")
            .FieldValues("PRECIO") = Format(0, "###,##0.00")
'            .FieldValues("PESO") = Format(0, "###,##0.00")
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

End Sub

Private Sub AdicionaItem2()
Dim i As Integer
Dim sw_nuevo_temp   As Boolean

    dxDBGrid2.Dataset.Active = False
    BASE_TEMPORAL "TEMPLUS.MDB"
    DELETEREC_LOG "DETFORMULA", Temp
    dxDBGrid2.Dataset.Refresh
    
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
            .FieldValues("CODIGO") = ""
            .FieldValues("DESCRIPCION") = ""
            .FieldValues("UM") = ""
            .FieldValues("CANTIDAD") = Format(0, "###,##0.00")
            .FieldValues("STOCK") = Format(0, "###,##0.00")
'            .FieldValues("PESO") = Format(0, "###,##0.00")
        Next
        .Post
        sw_nuevo_item = False
    End With
    dxDBGrid2.Dataset.Close
    dxDBGrid2.Dataset.Open

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Nuevo"
            nuevo
        Case "ID_Salir"
            Me.MousePointer = vbHourglass
            
            If sw_cabecera = True Or sw_detalle = True Then
                If MsgBox("Desea Grabar el Movimiento?", vbQuestion + vbYesNo, "Atenciòn") = vbYes Then
                    Verifica_Datos
                    Graba_Datos
                    sw_nuevo_doc = False
                    sw_detalle = False
                End If
            End If
            Me.MousePointer = vbDefault
            Unload Me
        Case "ID_Grabar"
            Me.MousePointer = vbHourglass
            Verifica_Datos
            If sw_mensaje = True Then
                If sw_cabecera = True Or sw_detalle = True Then
                     Graba_Datos
                End If
            End If
            Me.MousePointer = vbDefault
        Case "ID_Imprimir"
            Me.MousePointer = vbHourglass
            imprimir
            Me.MousePointer = vbDefault
        Case "ID_Eliminar"
            Me.MousePointer = vbHourglass
            Elimina_Movimientos
            Me.MousePointer = vbDefault
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
                    'SSFrame1.Caption = "Productos a Transferir del Almacén " & Trim(PnlNomAlm.Caption)
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
        wtipmov = "S"
        ayuda_conceptos.Show 1
        sw_ayudaO = False
        If Len(Trim(wconcepto)) > 0 Then
            Txtcodori.Text = wconcepto
            PnlNomOri.Caption = wnomconcepto
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
            If VALIDA_ORIGEN(Txtcodori.Text) = True Then
                wcodori = Txtcodori.Text
                PnlNomOri.Caption = wnomconcepto
            Else
                MsgBox "Código de Origen no existe. Verifique.", vbInformation, "Atención"
                Txtcodori.SetFocus
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
        ayuda_almacen.Show 1
        sw_ayuda = False
        If Len(Trim(wcod_alm)) > 0 Then
            Txtcodpar.Text = wcod_alm
            PnlNomPar.Caption = wnomalmacen
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
'                If Txtcodalm.Text = Txtcodpar.Text Then
'                    'MsgBox "El Almacén de Origen y el De Destino no Pueden ser los Mismos", vbInformation, "Sistema de Inventario"
'
'                    If MsgBox("El Almacén de Origen y el De Destino es el Mismos,Desea Continuar?", vbYesNo + vbInformation, "Sistema de Inventario") = vbNo Then
'
'                        Exit Sub
'                   End If
'                End If
                If VALIDA_ALMACENO(Txtcodpar.Text) = True Then
                    WcodPar = Trim(Txtcodpar.Text)
                    PnlNomPar.Caption = WNomPar
                    'SSFrame2.Caption = "Insumos de Productos a Ingresar al Almacén " & PnlNomPar.Caption
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
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").value = dxDBGrid1.Dataset.RecordCount + 1
            dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.ColumnByFieldName("STOCK").value = Format(0, "###,##0.00")
            dxDBGrid1.Columns.FocusedIndex = 0
        End If
    End If

End Sub

Private Sub dxDBGrid1_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid1.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid1.Columns(1).value & "")) = 0 Then
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
Dim cumventa    As String

    wcod_alm = Txtcodalm.Text
    wcodproducto = ""
    sw_ayuda_prod = True
    wtipoguia = "S"
    ayuda_productos.Show 1
    If Len(Trim(wcodproducto)) > 0 Then
        dxDBGrid1.Dataset.Edit
        'dxDBGrid1.Columns.ColumnByFieldName("CODFORMULA").Value = dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value
        dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value = wcodproducto
        dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
        dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value = 0#
        cumventa = ""
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "SELECT IF5PLA.F7CODMED,IF5PLA.F5FACTOR,IF5PLA.F5PRECOS FROM IF5PLA WHERE IF5PLA.F5CODPRO='" & wcodproducto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            cumventa = Trim(Rs.Fields("F7CODMED") & "")
            dxDBGrid1.Columns.ColumnByFieldName("PRECIO").value = Rs.Fields("F5PRECOS")
            dxDBGrid1.Columns.ColumnByFieldName("FACTOR").value = Rs.Fields("F5FACTOR")
'            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value = Costo_Unitario(wcodproducto, Txtcodalm.Text) / rs.Fields("F5FACTOR")
'            dxDBGrid1.Columns.ColumnByFieldName("COSTO").Value = (Costo_Unitario(wcodproducto, Txtcodalm.Text) / rs.Fields("F5FACTOR")) * 0#
        Else
            dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value = Costo_Unitario(wcodproducto, Txtcodalm.Text)
            dxDBGrid1.Columns.ColumnByFieldName("COSTO").value = Costo_Unitario(wcodproducto, Txtcodalm.Text) * 0#
        End If
        Rs.Close
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "SELECT F7SIGMED FROM EF7MEDIDAS WHERE F7CODMED='" & cumventa & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            dxDBGrid1.Columns.ColumnByFieldName("UM").value = "" & Rs.Fields("F7SIGMED")
        Else
            dxDBGrid1.Columns.ColumnByFieldName("UM").value = ""
        End If
        Rs.Close
        dxDBGrid1.Columns.ColumnByFieldName("STOCK").value = 0#
        dxDBGrid1.Dataset.Post
    End If

End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Set RSCONSULTA = New ADODB.Recordset

    If dxDBGrid1.Dataset.State <> 0 And dxDBGrid1.Dataset.State <> 1 Then
        If dxDBGrid1.Columns.FocusedColumn.FieldName = "CODIGO" Then
            wcodproducto = dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value
            If Len(Trim(wcodproducto)) > 0 Then
                If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
                sql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB,IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO,IF5PLA.F5VALVTA, IF5PLA.F7CODMED FROM IF5PLA,IF6ALMA WHERE  (IF6ALMA.F2CODALM='" & Trim(Txtcodpar.Text) & "') AND ((IF5PLA.F5CODPRO='" & wcodproducto & "')) ORDER BY IF5PLA.F5CODPRO;"
                RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RSCONSULTA.EOF Then
                    wcodproducto = "" & Trim(RSCONSULTA.Fields("F5CODPRO"))
                    wcodfab = "" & RSCONSULTA.Fields("F5CODFAB")
                    wdesproducto = "" & RSCONSULTA.Fields("F5NOMPRO")
                    wmedida = "" & RSCONSULTA.Fields("F7CODMED")
                    WMONEDAX = "" & RSCONSULTA.Fields("F5MONEDA")
                    'wstock = Val("" & rsconsulta.Fields("F6STOCKACT"))
                    wvalvta = Val("" & RSCONSULTA.Fields("F5VALVTA"))
                Else
                    MsgBox "Producto no existe en el almacen indicado", vbInformation + vbDefaultButton1, "Sistema de Logistica"
                    wcodproducto = "": wdesproducto = "": wmedida = "": wcodfab = ""
                    wvalvta = 0#: wstock = 0#: WMONEDAX = ""
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value = wcodproducto
                    dxDBGrid1.Dataset.Post
                    Exit Sub
                End If
                RSCONSULTA.Close
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("CODIGO").value = wcodproducto
                dxDBGrid1.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
                
                If Rs.State = adStateOpen Then Rs.Close
                Rs.Open "SELECT F7SIGMED FROM EF7MEDIDAS WHERE F7CODMED='" & wmedida & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not Rs.EOF Then
                    dxDBGrid1.Columns.ColumnByFieldName("UM").value = "" & Rs.Fields("F7SIGMED")
                Else
                    dxDBGrid1.Columns.ColumnByFieldName("UM").value = ""
                End If
                Rs.Close
        
                dxDBGrid1.Dataset.Post
                dxDBGrid1.Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex - 1
            End If
        Else
            If dxDBGrid1.Columns.FocusedColumn.FieldName = "CANTIDAD" Or dxDBGrid1.Columns.FocusedColumn.FieldName = "COSTOUNI" Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("COSTOTAL").value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").value
                dxDBGrid1.Dataset.Post
            End If
            
        End If
    End If

End Sub

Private Sub dxDBGrid2_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid2.Columns.ColumnByFieldName("ITEM").value = dxDBGrid2.Dataset.RecordCount + 1
            dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").value = Format(0, "###,##0.00")
            dxDBGrid2.Columns.ColumnByFieldName("STOCK").value = Format(0, "###,##0.00")
            dxDBGrid2.Columns.FocusedIndex = 0
        End If
    End If

End Sub

Private Sub dxDBGrid2_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If dxDBGrid2.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid2.Columns(1).value & "")) = 0 Then
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
Dim cumventa    As String

    wcod_alm = Txtcodalm.Text
    wcodproducto = ""
    sw_ayuda_prod = True
    wtipoguia = "S"
    ayuda_productos.Show 1
    If Len(Trim(wcodproducto)) > 0 Then
        dxDBGrid2.Dataset.Edit
        'dxDBGrid2.Columns.ColumnByFieldName("CODFORMULA").Value = dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value
        dxDBGrid2.Columns.ColumnByFieldName("CODIGO").value = wcodproducto
        dxDBGrid2.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
        dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").value = 0#
        cumventa = ""
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "SELECT IF5PLA.F7CODMED,IF5PLA.F5PRECOS FROM IF5PLA WHERE IF5PLA.F5CODPRO='" & wcodproducto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            cumventa = Trim(Rs.Fields("F7CODMED") & "")
            'dxDBGrid2.Columns.ColumnByFieldName("PRECIO").Value = RS.Fields("F5PRECOS")
            'dxDBGrid2.Columns.ColumnByFieldName("FACTOR").Value = RS.Fields("F5FACTOR")
            dxDBGrid2.Columns.ColumnByFieldName("COSTOUNI").value = Costo_Unitario(wcodproducto, Txtcodpar.Text)
            'dxDBGrid2.Columns.ColumnByFieldName("COSTO").Value = (Costo_Unitario(wcodproducto, Txtcodpar.Text) / RS.Fields("F5FACTOR")) * 0#
        Else
            dxDBGrid2.Columns.ColumnByFieldName("COSTOUNI").value = Costo_Unitario(wcodproducto, Txtcodpar.Text)
            'dxDBGrid2.Columns.ColumnByFieldName("COSTO").Value = Costo_Unitario(wcodproducto, Txtcodpar.Text) * 0#
        End If
        Rs.Close
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "SELECT F7SIGMED FROM EF7MEDIDAS WHERE F7CODMED='" & cumventa & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Rs.EOF Then
            dxDBGrid2.Columns.ColumnByFieldName("UM").value = "" & Rs.Fields("F7SIGMED")
        Else
            dxDBGrid2.Columns.ColumnByFieldName("UM").value = ""
        End If
        Rs.Close
        dxDBGrid2.Columns.ColumnByFieldName("STOCK").value = 0#
        dxDBGrid2.Dataset.Post
        
'        listarSubProductos wcodproducto
    End If
    
End Sub

Private Sub listarSubProductos(cod As String)
Dim rstemp       As New ADODB.Recordset
Dim csqlTemp     As String
Dim sCod         As String
Dim nNom         As String
Dim sUM          As String
Dim sStock       As String
Dim sPrecio     As String

 
    
'rstemp.Open "Select * from IF3SUBPRODUCTOS where F4CODPRO = '" & cod & "'", cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
BASE_TEMPORAL "TEMPLUS.MDB"

'If Not rstemp.EOF Then
  dxDBGrid1.Dataset.Active = False
  BASE_TEMPORAL "TEMPLUS.MDB"
  DELETEREC_LOG "DETCONVERSION", Temp
  dxDBGrid1.Dataset.Refresh
'Else
'    MsgBox "No hay Sub Productos asignados para  el Producto seleccionado", vbInformation, "Sistema"
'    Exit Sub
'End If

'Do While Not rstemp.EOF
    sCod = wcodproducto
    nNom = wdesproducto & ""
    sUM = wmedida & ""
    sStock = "0" 'rstemp.Fields("") & ""
    sPrecio = "0" 'rstemp.Fields("") & ""
    
    csqlTemp = "insert into DETCONVERSION (CODIGO,DESCRIPCION,UM,STOCK,PRECIO)" & _
                " values ('" & sCod & "','" & nNom & "','" & sUM & "'," & sStock & "," & sPrecio & ")"
    Temp.Execute csqlTemp '"insert into DETCONVERSION values (CODIGO,DESCRIPCION,UM,STOCK,PRECIO)"
'   AlmacenaQuery_sql csqlTemp, Temp
'              .FieldValues("CODIGO") = ""
'            .FieldValues("DESCRIPCION") = ""
'            .FieldValues("UM") = ""
'            .FieldValues("CANTIDAD") = Format(0, "###,##0.00")
'            .FieldValues("STOCK") = Format(0, "###,##0.00")
'            .FieldValues("PRECIO") = Format(0, "###,##0.00")
'            .FieldValues("PESO") = Format(0, "###,##0.00")
'    rstemp.MoveNext
'Loop

BASE_TEMPORAL "TEMPLUS.MDB"

dxDBGrid1.Dataset.Active = False
BASE_TEMPORAL "TEMPLUS.MDB"
dxDBGrid1.Dataset.Refresh

dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
dxDBGrid1.Dataset.Active = True
dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open
'dxDBGrid1.Dataset.Refresh

End Sub



Private Sub dxDBGrid2_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Set RSCONSULTA = New ADODB.Recordset
'    If dxDBGrid2.Dataset.State = dsEdit Or dxDBGrid2.Dataset.State = dsInsert Then
'        dxDBGrid2.Dataset.Post
'        dxDBGrid2.Dataset.Refresh
'    End If

    If dxDBGrid2.Dataset.State <> 0 And dxDBGrid2.Dataset.State <> 1 Then
        If dxDBGrid2.Columns.FocusedColumn.FieldName = "CODIGO" Then
            wcodproducto = dxDBGrid2.Columns.ColumnByFieldName("CODIGO").value
            If Len(Trim(wcodproducto)) > 0 Then
                If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
                sql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB,IF5PLA.F5MONEDA, IF5PLA.F5NOMPRO, IF5PLA.F5VALVTA, IF5PLA.F7CODMED FROM IF5PLA,IF6ALMA WHERE  (IF6ALMA.F2CODALM='" & Trim(Txtcodalm.Text) & "') AND ((IF5PLA.F5CODPRO='" & wcodproducto & "')) ORDER BY IF5PLA.F5CODPRO;"
                RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not RSCONSULTA.EOF Then
                    wcodproducto = "" & Trim(RSCONSULTA.Fields("F5CODPRO"))
                    wcodfab = "" & RSCONSULTA.Fields("F5CODFAB")
                    wdesproducto = "" & RSCONSULTA.Fields("F5NOMPRO")
                    wmedida = "" & RSCONSULTA.Fields("F7CODMED")
                    WMONEDAX = "" & RSCONSULTA.Fields("F5MONEDA")
                    'wstock = Val("" & rsconsulta.Fields("F6STOCKACT"))
                    wvalvta = Val("" & RSCONSULTA.Fields("F5VALVTA"))
                Else
                    MsgBox "Producto no existe en el almacen indicado", vbInformation + vbDefaultButton1, "Atención"
                    wcodproducto = "": wdesproducto = "": wmedida = "": wcodfab = ""
                    wvalvta = 0#: wstock = 0#: WMONEDAX = ""
                    dxDBGrid2.Dataset.Edit
                    dxDBGrid2.Columns.ColumnByFieldName("CODIGO").value = wcodproducto
                    dxDBGrid2.Dataset.Post
                    Exit Sub
                End If
                RSCONSULTA.Close
                dxDBGrid2.Dataset.Edit
                dxDBGrid2.Columns.ColumnByFieldName("CODIGO").value = wcodproducto
                dxDBGrid2.Columns.ColumnByFieldName("DESCRIPCION").value = wdesproducto
                
                If Rs.State = adStateOpen Then Rs.Close
                Rs.Open "SELECT F7SIGMED FROM EF7MEDIDAS WHERE F7CODMED='" & wmedida & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not Rs.EOF Then
                    dxDBGrid2.Columns.ColumnByFieldName("UM").value = "" & Rs.Fields("F7SIGMED")
                Else
                    dxDBGrid2.Columns.ColumnByFieldName("UM").value = ""
                End If
                Rs.Close
        
                dxDBGrid2.Dataset.Post
                dxDBGrid2.Columns.FocusedIndex = dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").ColIndex - 1
            End If
        Else
            If dxDBGrid2.Columns.FocusedColumn.FieldName = "CANTIDAD" Or dxDBGrid2.Columns.FocusedColumn.FieldName = "COSTOUNI" Then
                dxDBGrid2.Dataset.Edit
                dxDBGrid2.Columns.ColumnByFieldName("COSTOTAL").value = dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").value * dxDBGrid2.Columns.ColumnByFieldName("COSTOUNI").value
                dxDBGrid2.Dataset.Post
            End If
            
        End If
    End If

End Sub

Private Sub nuevo()
    On Error GoTo error33
    Me.MousePointer = vbHourglass
    sw_nuevo_doc = False
    sw_detalle = False
    Txtcodalm.Text = "": PnlNomAlm.Caption = ""
    Txtcodpar.Text = "": PnlNomPar.Caption = ""
    Txtcodori.Text = "": PnlNomOri.Caption = ""
    Txtnumvalo.Caption = "": Txtnumvald.Caption = ""
    
    BASE_TEMPORAL "TEMPLUS.MDB"
    DELETEREC_LOG "DETCONVERSION", Temp
    DELETEREC_LOG "DETFORMULA", Temp
    
    dxDBGrid1.Dataset.Refresh
    dxDBGrid2.Dataset.Refresh
    Conf_Grid dxDBGrid1
    Conf_Grid dxDBGrid2
    AdicionaItem
    AdicionaItem2
    
    Txtcodalm.SetFocus
    sw_nuevo_doc = True
    Me.MousePointer = vbDefault
    SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
   Exit Sub
error33:
   Resume Next
   
End Sub

Private Sub Verifica_Datos()
Dim Mensaje As String

    If Trim(Txtcodalm.Text) = "" Or Trim(Txtcodpar.Text) = "" Then
        MsgBox "Debe Ingresar Almacen de Origen Y de Destino", vbInformation, "Sistema de Logística"
        sw_mensaje = False
        Exit Sub
    End If
    
'    If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOTAL").SummaryFooterValue, "0.00")) <> Val(Format(dxDBGrid2.Columns.ColumnByFieldName("COSTOTAL").SummaryFooterValue, "0.00")) Then
'        MsgBox "Los costos no coinciden. Verifique.", vbInformation, "Atención"
'        sw_mensaje = False
'        Exit Sub
'    End If
    
    If sw_nuevo_doc = False Then
        Mensaje = "Los vales de conversión no han sido grabados. Desea Grabar ?"
    Else
        Mensaje = "Desea grabar los vales de conversión ?"
    End If
    
    If MsgBox(Mensaje, vbYesNo + vbInformation, "Atención") = vbYes Then
        sw_mensaje = True
    Else
        sw_mensaje = False
    End If

End Sub

Private Sub Graba_Datos()

'    'If dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").SummaryFooterValue = 0 Then
'    '    MsgBox "Debe Ingresar la Cantidad a Producir", vbInformation, "AVISO"
'    '    Exit Sub
'    'End If
    
    
    
    If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
        dxDBGrid1.Dataset.Post
        dxDBGrid1.Dataset.Refresh
        sw_nuevo_item = True
    End If
    
    If dxDBGrid2.Dataset.State = dsEdit Or dxDBGrid2.Dataset.State = dsInsert Then
        dxDBGrid2.Dataset.Post
        dxDBGrid1.Dataset.Refresh
        sw_nuevo_item = True
    End If
    
'    If Val(Format(dxDBGrid1.Columns.ColumnByFieldName("PESO").SummaryFooterValue, "0.000")) <> Val(Format(dxDBGrid2.Columns.ColumnByFieldName("PESO").SummaryFooterValue, "0.000")) Then
'        MsgBox "Los Pesos no coinciden. Verifique.", vbInformation, "Atención"
'        sw_nuevo_item = True
'        Exit Sub
'    End If
    
    If Len(Txtcodori.Text) = 0 Then
        MsgBox "Ingrese Concepto...", vbCritical, "Atención"
        Exit Sub
    Else
        REACTUALIZAR_ALMACENES "S", Txtcodalm.Text, Txtcodpar.Text, "DETFORMULA"
        REACTUALIZAR_ALMACENES "I", Txtcodpar.Text, Txtcodalm.Text, "DETCONVERSION"
        SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
        SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
    End If
    
    sw_nuevo_doc = False
    sw_detalle = False
    sw_cabecera = False
    MsgBox "Vales de conversión, grabados.", vbInformation, "Sistema de Logística"
    
End Sub

Private Sub Elimina_Movimientos()
Dim SSQL, csql As String
Set RsStockCab = New ADODB.Recordset
Set RsStockDet = New ADODB.Recordset

If RsStockCab.State = adStateOpen Then RsStockCab.Close
sql = "SELECT * FROM IF4VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
RsStockCab.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockCab.EOF Then
    If MsgBox("Está seguro de eliminar los movimientos registrados", vbInformation + vbYesNo, "Sistema de Logística") = vbYes Then
        '------------------------------------------------------------------------------------------------------------------------------
        If RsStockDet.State = adStateOpen Then RsStockDet.Close
        sql = "SELECT * FROM IF3VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        RsStockDet.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsStockDet.EOF Then
            Do While Not RsStockDet.EOF
                Reactualiza_Almacenes Trim(RsStockDet.Fields("F2codalm")), Trim(RsStockDet.Fields("F5codpro")), Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), CVDate(RsStockDet.Fields("F4fecval")), Val(Format(RsStockDet.Fields("F3totite"), "#0.000")), Val(Format(RsStockDet.Fields("F3totdol"), "#0.000")), "S", Val(Format(RsStockDet.Fields("F3valdol"), "#0.000"))
                RsStockDet.MoveNext
                If RsStockDet.EOF Then Exit Do
            Loop
        End If
        
        SSQL = "DELETE FROM IF4VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        cnn_dbbancos.Execute (SSQL)
         'AlmacenaQuery_sql SSQL, cnn_dbbancos
         
        csql = "DELETE FROM IF3VALES WHERE F2CODALM = '" & Txtcodpar.Text & "' AND F4NUMVAL = '" & Txtnumvald.Caption & "'"
        cnn_dbbancos.Execute (csql)
         'AlmacenaQuery_sql csql, cnn_dbbancos
        
        If RsStockCab.State = adStateOpen Then RsStockCab.Close
        sql = "SELECT * FROM IF4VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
        RsStockCab.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsStockCab.EOF Then
            
            '----------DETALLE ---------
            If RsStockDet.State = adStateOpen Then RsStockDet.Close
            sql = "SELECT * FROM IF3VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
            RsStockDet.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsStockDet.EOF Then
                Do While Not RsStockDet.EOF
                    Reactualiza_Almacenes Trim(RsStockDet.Fields("F2codalm")), Trim(RsStockDet.Fields("F5codpro")), Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), CVDate(RsStockDet.Fields("F4fecval")), Val(Format(RsStockDet.Fields("F3totite"), "#0.000")), Val(Format(RsStockDet.Fields("F3totdol"), "#0.000")), "I", Val(Format(RsStockDet.Fields("F3valdol"), "#0.000"))
                    RsStockDet.MoveNext
                    If RsStockDet.EOF Then Exit Do
                Loop
            End If
            
            '-----------CABECERA ----------
            SSQL = "DELETE FROM IF4VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
            cnn_dbbancos.Execute (SSQL)
             'AlmacenaQuery_sql SSQL, cnn_dbbancos
             
            csql = "DELETE FROM IF3VALES WHERE F2CODALM = '" & Txtcodalm.Text & "' AND F4NUMVAL = '" & Txtnumvalo.Caption & "'"
            cnn_dbbancos.Execute (csql)
             'AlmacenaQuery_sql csql, cnn_dbbancos
             
            nuevo
            Txtcodalm.SetFocus
        End If
    End If
Else
    MsgBox "El Registro no ha sido Grabado", vbInformation, "Sistema de Inventarios"
    Exit Sub
End If
End Sub

Private Sub RENUMERARITEMS(pgrid As Control)
Dim i As Integer

    sw_nuevo_item = True
    pgrid.Dataset.First
    Do While Not pgrid.Dataset.EOF
        i = i + 1
        pgrid.Dataset.Edit
        pgrid.Columns.ColumnByFieldName("ITEM").value = i
        pgrid.Dataset.Next
    Loop
    sw_nuevo_item = False

End Sub

Private Sub ACTUALIZA_ALMA_VALE(pnumvale As String, pcampo As String, palmacen As String)
Dim csql    As String
        
    csql = "UPDATE EF2ALMACENES SET " & pcampo & " =  '" & pnumvale & "' WHERE '" & pnumvale & "' > " & pcampo & " AND F2CODALM='" & palmacen & "'"
    cnn_dbbancos.Execute csql
     'AlmacenaQuery_sql csql, cnn_dbbancos
    
End Sub

Private Sub imprimir()
    
    With Acr_Vale_Salida
        .DataControl1.ConnectionString = cnn_dbbancos
        '.DataControl1.Source = "SELECT *,IF5PLA.F5CODPRO,F5NOMPRO,F7CODMED,f5codfab,f2desmar FROM IF3VALES,IF5PLA,ef2marcas WHERE IF5PLA.F5CODPRO=IF3VALES.F5CODPRO and ef2marcaS.f2codmar=if5pla.f5marca AND F2CODALM='" & Txtcodalm.Text & "' AND F4NUMVAL='" & Txtnumvalo.Caption & "'"
        ',B.F5NOMMARCA
        sql = "SELECT A.F5CODPRO, B.F5NOMPRO,B.F5CODFAB, B.F7CODMED, Sum(A.F3CANPRO) AS F3CANPRO FROM IF3VALES AS A INNER JOIN IF5PLA AS B ON A.F5CODPRO = B.F5CODPRO Where ((A.F4NUMVAL = '" & Txtnumvalo.Caption & "') And (A.F2CODALM = '" & Txtcodalm.Text & "')) GROUP BY A.F5CODPRO, B.F5NOMPRO,B.F5CODFAB, B.F7CODMED ORDER BY A.F5CODPRO;"
        .DataControl1.Source = sql
        .fldempresa.Text = wnomcia
        .fldFecha.Text = TxtFecMov.value
        .Lbl_vale.Caption = "VALE DE SALIDA"
        .fldalma.Text = Txtcodalm.Text
        .fldAlmacen.Text = PnlNomAlm.Caption
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
        sql = "SELECT A.F5CODPRO, B.F5NOMPRO, B.F5CODFAB,B.F7CODMED, Sum(A.F3CANPRO) AS F3CANPRO FROM IF3VALES AS A INNER JOIN IF5PLA AS B ON A.F5CODPRO = B.F5CODPRO Where ((A.F4NUMVAL = '" & Txtnumvald.Caption & "') And (A.F2CODALM = '" & Txtcodpar.Text & "')) GROUP BY A.F5CODPRO, B.F5NOMPRO, B.F5CODFAB,B.F7CODMED ORDER BY A.F5CODPRO"
        .DataControl1.Source = sql
        .fldempresa.Text = wnomcia
        .fldFecha.Text = TxtFecMov.value
        .Lbl_vale.Caption = "VALE DE INGRESO"
        .fldalma.Text = Txtcodpar.Text
        .fldAlmacen.Text = PnlNomPar.Caption
        .fldvale.Text = Txtnumvald.Caption
        .fldcon.Text = Txtcodori.Text
        .F1NOMORI.Text = PnlNomOri.Caption
        .flddoc.Visible = True
        .NUMDOC.Visible = True
        .flddoc.Text = Txtnumvalo.Caption
        .Show 1
    End With
End Sub

Private Sub Txtfecmov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtcodalm.SetFocus
End Sub

Private Sub TxtFecMov_LostFocus()
If IsDate(TxtFecMov.value) = False Then
    MsgBox "Ingrese Correctamente la Fecha", vbCritical, "ATENCION"
    TxtFecMov.value = Format(Date, "DD/MM/YYYY")
    TxtFecMov.SetFocus
End If
End Sub

Private Sub REACTUALIZAR_ALMACENES(ptipo As String, pcodalm As String, pcodpar As String, ptabla As String)
Dim RsAnulaciones   As New ADODB.Recordset
Dim RsCalculaCosto  As New ADODB.Recordset
Dim XConcepto       As String
Dim XNumeroVale     As String
Dim XCampo          As String
Dim nitems          As Integer
Dim nfila           As Integer
Dim XCospro         As Double
Dim XCosproD        As Double
Dim XNumeroDoc      As String
Dim xtipo           As String
Dim sw_adicionaitem As Boolean

    If ptipo = "I" Then xtipo = "S" Else xtipo = "I"
    If sw_nuevo_doc = True Then
        ctipo = "A"
    Else
        ctipo = "M"
    End If
    XNumeroVale = Calcula_Numero(Trim(pcodalm), ptipo)
    XNumeroDoc = Calcula_Numero(Trim(pcodpar), xtipo)
    
    If ptipo = "S" Then
        XConcepto = "CV1"
        XCampo = "F1VALSAL" & Format(Month(TxtFecMov.value), "00")
        Txtnumvalo.Caption = XNumeroVale
        Txtnumvald.Caption = XNumeroDoc
    Else
        XConcepto = "CV0"
        XCampo = "F1VALING" & Format(Month(TxtFecMov.value), "00")
    End If
    ACTUALIZA_ALMA_VALE XNumeroVale, XCampo, pcodalm
    '---------------------ASIGNA DATOS A IF4VALES -------------------------------------
    amovs_cab(0).campo = "F2CODALM": amovs_cab(0).valor = pcodalm: amovs_cab(0).Tipo = "T"
    amovs_cab(1).campo = "F4NUMVAL": amovs_cab(1).valor = XNumeroVale: amovs_cab(1).Tipo = "T"
    amovs_cab(2).campo = "F4FECVAL": amovs_cab(2).valor = TxtFecMov.value: amovs_cab(2).Tipo = "F"
    amovs_cab(3).campo = "F2CODPAR": amovs_cab(3).valor = pcodpar: amovs_cab(3).Tipo = "T"
    amovs_cab(4).campo = "F2CODPROV": amovs_cab(4).valor = "0": amovs_cab(4).Tipo = "T"
    amovs_cab(5).campo = "F1CODORI": amovs_cab(5).valor = XConcepto: amovs_cab(5).Tipo = "T"
    amovs_cab(6).campo = "F1CODDOC": amovs_cab(6).valor = "": amovs_cab(6).Tipo = "T"
    amovs_cab(7).campo = "F4NUMDOC": amovs_cab(7).valor = XNumeroDoc: amovs_cab(7).Tipo = "T"
    amovs_cab(8).campo = "F4MONEDA": amovs_cab(8).valor = wmoneda_productos: amovs_cab(8).Tipo = "T"
    amovs_cab(9).campo = "F4TIPCAM": amovs_cab(9).valor = wtipcam: amovs_cab(9).Tipo = "T"
    amovs_cab(10).campo = "F4FECULT": amovs_cab(10).valor = Format(Date, "DD/MM/YYYY"): amovs_cab(10).Tipo = "F"
    amovs_cab(11).campo = "F2CODUSE": amovs_cab(11).valor = wusuario: amovs_cab(11).Tipo = "T"
    
    '*-----------------------------DETALLE------------------------------------
    amovs_det(0).campo = "F4NUMVAL": amovs_det(0).valor = "": amovs_det(0).Tipo = "T"
    amovs_det(1).campo = "F5CODPRO": amovs_det(1).valor = "": amovs_det(1).Tipo = "T"
    amovs_det(2).campo = "F3CANPRO": amovs_det(2).valor = "": amovs_det(2).Tipo = "N"
    amovs_det(3).campo = "F3VALVTA": amovs_det(3).valor = "": amovs_det(3).Tipo = "N"
    amovs_det(4).campo = "F2CODALM": amovs_det(4).valor = "": amovs_det(4).Tipo = "T"
    amovs_det(5).campo = "F4FECVAL": amovs_det(5).valor = "": amovs_det(5).Tipo = "F"
    amovs_det(6).campo = "F3VALDOL": amovs_det(6).valor = "": amovs_det(6).Tipo = "N"
    amovs_det(7).campo = "F3TOTITE": amovs_det(7).valor = "": amovs_det(7).Tipo = "N"
    amovs_det(8).campo = "F3TOTDOL": amovs_det(8).valor = "": amovs_det(8).Tipo = "N"
    amovs_det(9).campo = "F3GRUPO": amovs_det(9).valor = "": amovs_det(9).Tipo = "T"
'    amovs_det(10).campo = "PESO": amovs_det(10).valor = "": amovs_det(10).TIPO = "N"
    
    '---------------------Calcula el Numero de Filas
    nitems = 0
    
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    sql = "Select count(ITEM) as NITEM from " & ptabla & " Where LEN(TRIM(ITEM))> 0 AND LEN(TRIM(CODIGO))>0 "
    RSCONSULTA.Open sql, Temp, adOpenDynamic, adLockOptimistic
    If Not RSCONSULTA.EOF Then
        nitems = Val("" & RSCONSULTA.Fields("NITEM"))
    End If
    RSCONSULTA.Close
    
    ReDim Values(9, nitems)
    
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open "Select * from " & ptabla & " WHERE LEN(TRIM(ITEM))>0 AND LEN(TRIM(CODIGO))>0", Temp
    If Not RSCONSULTA.EOF Then
    nfila = 0
    
    RSCONSULTA.MoveFirst
    Do While Not RSCONSULTA.EOF
         If Len(Trim(RSCONSULTA.Fields("CODIGO") & "")) > 0 Then
         
            sw_adicionaitem = False
            'If ptipo = "S" Then
                If Val(RSCONSULTA.Fields("CANTIDAD") & "") > 0# Then
                    sw_adicionaitem = True
                End If
            'Else
            '    If Val(rsconsulta.Fields("CANTIDAD") & "") > 0# Then
            '        sw_adicionaitem = True
            '    End If
            'End If
         
            If sw_adicionaitem = True Then
                Values(0, nfila) = XNumeroVale
                Values(1, nfila) = "" & RSCONSULTA.Fields("CODIGO")
                'Values(2, nfila) = Val("" & rsconsulta.Fields("CANTIDAD"))
                Values(2, nfila) = Val("" & RSCONSULTA.Fields("CANTIDAD"))
                If wmoneda_productos = "S" Then
                    Values(3, nfila) = Val("" & RSCONSULTA.Fields("COSTOUNI"))
                    Values(6, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("COSTOUNI") / wtipcam), "0.000"))
                    Values(7, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("CANTIDAD") * Val(RSCONSULTA.Fields("COSTOUNI"))), "0.000"))
                    Values(8, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("CANTIDAD") * (Val(RSCONSULTA.Fields("COSTOUNI")) / wtipcam)), "0.000"))
                    XCospro = Val("" & RSCONSULTA.Fields("COSTOUNI"))
                    XCosproD = Val("" & Format(Val(RSCONSULTA.Fields("COSTOUNI") / wtipcam), "0.000"))
                Else
                    Values(3, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("COSTOUNI") * wtipcam), "0.000"))
                    Values(6, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("COSTOUNI")), "0.000"))
                    Values(7, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("CANTIDAD") * Val(RSCONSULTA.Fields("COSTOUNI") * wtipcam)), "0.000"))
                    Values(8, nfila) = Val("" & Format(Val(RSCONSULTA.Fields("CANTIDAD")) * Val(RSCONSULTA.Fields("COSTOUNI")), "0.000"))
                    XCospro = Val("" & Format(Val(RSCONSULTA.Fields("COSTOUNI") * wtipcam), "0.000"))
                    XCosproD = Val("" & Format(Val(RSCONSULTA.Fields("COSTOUNI")), "0.000"))
                End If
                Values(4, nfila) = "" & pcodalm
                Values(5, nfila) = "" & CVDate(TxtFecMov.value)
                Values(9, nfila) = "" 'Left(rsconsulta.Fields("CODIGO"), 2)
                'Values(10, nfila) = Val("" & rsconsulta.Fields("PESO"))
'                Values(10, nfila) = Val("" & rsconsulta.Fields("CANTIDAD"))
                Rem NSE Vales_Detalle XNumeroVale, rsconsulta.Fields("CODIGO"), rsconsulta.Fields("CANTIDAD"), XCosproD, pcodalm, CVDate(TxtFecMov.Value), XCospro
                nfila = nfila + 1
            End If
        End If
        RSCONSULTA.MoveNext
    Loop
    End If
    RSCONSULTA.Close
    sw_GRABA_REGISTRO_logistica = True
    
    If ctipo = "A" Then '---Nuevo
        '-----Graba Cabecera
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", "A", 11, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 9, cnn_dbbancos, "", Values(), nfila - 1, "11111111111", "", ""
        End If
    Else    '----------Modificacion
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", "M", 11, cnn_dbbancos, "F4NUMVAL = '" & XNumeroVale & "' AND F2CODALM = '" & pcodalm & "'"
        
        '------- GRABA DETALLE
        
        csql = ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & XNumeroVale & "' AND F2CODALM = '" & pcodalm & "'")
        cnn_dbbancos.Execute csql
         'AlmacenaQuery_sql csql, cnn_dbbancos
        
        GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 9, cnn_dbbancos, "", Values(), nfila - 1, "1111111111", "", ""
    End If
    
    
    '''graba envio
    If wIndEnvia = "*" Then 'cnn_dbEnvia
'        If ctipo = "A" Then '---Nuevo
            '-----Graba Cabecera
            csql = ("DELETE * FROM IF4VALES WHERE F4NUMVAL = '" & XNumeroVale & "' AND F2CODALM = '" & pcodalm & "'")
            cnn_dbEnvia.Execute csql
             'AlmacenaQuery_sql csql, cnn_dbEnvia
            
            GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", "A", 11, cnn_dbEnvia, ""
            
            If sw_GRABA_REGISTRO_logistica = True Then
                '------- GRABA DETALLE
                GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 9, cnn_dbEnvia, "", Values(), nfila - 1, "11111111111", "", ""
            End If
'        Else    '----------Modificacion
'            '------- GRABA CABECERA
'            GRABA_REGISTRO_logistica amovs_cab(), "IF4VALES", "M", 11, cnn_dbEnvia, "F4NUMVAL = '" & XNumeroVale & "' AND F2CODALM = '" & pcodalm & "'"
'
'            '------- GRABA DETALLE
'            cnn_dbEnvia.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & XNumeroVale & "' AND F2CODALM = '" & pcodalm & "'")
'            GRABA_REGISTRO_logistica_DET amovs_det(), "IF3VALES", "A", 9, cnn_dbEnvia, "", Values(), nfila - 1, "1111111111", "", ""
'        End If
    End If

End Sub
