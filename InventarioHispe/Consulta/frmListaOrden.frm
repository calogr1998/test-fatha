VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListaOrden 
   Caption         =   "Lista de Ordenes"
   ClientHeight    =   8265
   ClientLeft      =   2880
   ClientTop       =   1770
   ClientWidth     =   10185
   Icon            =   "frmListaOrden.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   10185
   WindowState     =   2  'Maximized
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   870
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10005
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   9600
      End
   End
   Begin MSComDlg.CommonDialog cmdlgOrden 
      Left            =   0
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgOrden 
      Height          =   4515
      Left            =   120
      OleObjectBlob   =   "frmListaOrden.frx":058A
      TabIndex        =   2
      Top             =   1080
      Width           =   9975
   End
   Begin MSComctlLib.ImageList imgLstEstado 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":B1C8
            Key             =   "Estado 1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":B762
            Key             =   "Estado 2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":BCFC
            Key             =   "Estado 3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":C296
            Key             =   "Estado 4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":C830
            Key             =   "Estado 5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":CDCA
            Key             =   "Estado 6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":D364
            Key             =   "Estado 7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaOrden.frx":D8FE
            Key             =   "Estado 8"
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars tlbOrden 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   16
      ShowShortcutsInToolTips=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "frmListaOrden.frx":DE98
      ToolBars        =   "frmListaOrden.frx":1A949
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dbgOrdenDetalle 
      Height          =   2250
      Left            =   120
      OleObjectBlob   =   "frmListaOrden.frx":1AC46
      TabIndex        =   3
      Top             =   5640
      Width           =   9945
   End
End
Attribute VB_Name = "frmListaOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolAyuda        As Boolean

Private strFichero          As String

Dim d As Double
Dim nSaveRecNo As Double
'Private cImgInfo            As cImageInfo

Public Property Let Ayuda(ByVal value As Boolean)
    bolAyuda = value
End Property

Public Property Get Ayuda() As Boolean
    Ayuda = bolAyuda
End Property



Private Sub listarTipoOrden()
    tlbOrden.Tools("TipoOrden").ComboBox.Clear
    
    tlbOrden.Tools("TipoOrden").ComboBox.AddItem "(*) Todos" & Space(100)
    tlbOrden.Tools("TipoOrden").ComboBox.AddItem "Orden de Compra" & Space(100) & "OC"
    tlbOrden.Tools("TipoOrden").ComboBox.AddItem "Orden de Servicio" & Space(100) & "OS"
    tlbOrden.Tools("TipoOrden").ComboBox.AddItem "Oferta" & Space(100) & "OF"
    
    tlbOrden.Tools("TipoOrden").ComboBox.ListIndex = 1
End Sub


Private Sub dbgDocumento_RowColChange()
    If dbgOrden.Dataset.RecordCount > 0 Then
        If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
            listarOrdenDetalleSql
        Else
            listarOrdenDetalle
        End If
    End If
End Sub

Private Sub dbgOrden_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Select Case dbgOrden.Columns.FocusedColumn.FieldName
        Case "F4LOCAL", "F4NUMORD"
            dbgDocumento_RowColChange
    End Select
End Sub

Private Sub dbgOrden_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
    Select Case Column.FieldName
        Case "F4VB1"
            With dbgOrden
                .Dataset.Edit
                
                With objAyudaOrden
                    .inicializarEntidades
                    
                    .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
                    .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
                    
                    .obtenerConfigOrden
                    
                    If .Estado = 8 Then
                        MsgBox "Imposible realizar esta acción, registro Anulado.", vbInformation + vbOKOnly, App.ProductName
                        
                        dbgOrden.Dataset.Cancel
                        
                        Exit Sub
                    End If
                    
                    If .Estado >= 3 Then
                        MsgBox "Imposible realizar esta acción, registro en etapa superior.", vbInformation + vbOKOnly, App.ProductName
                        
                        dbgOrden.Dataset.Cancel
                        
                        Exit Sub
                    End If
                End With
                
                If Not CBool(objAyudaOrden.VB1) Then
                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0006'") = "0006" Then
                        If MsgBox("¿Desea DAR su VºBº al registro seleccionado?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            .Columns.ColumnByFieldName("F4ESTADO").value = 2
                            .Columns.ColumnByFieldName("F4VB1").value = True
                            .Columns.ColumnByFieldName("F4VBUSER1").value = UCase(wusuario)
                            .Columns.ColumnByFieldName("F4VBFECHA1").value = Now
                            
                            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                With objSqlAyudaOrden
                                    .inicializarEntidades
                                    
                                    .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
                                    .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
                                    .Estado = 2
                                    .VB1 = True
                                    .VB1Usuario = UCase(wusuario)
                                    .VB1Fecha = Now
                                    
                                    If .aprobarOrden Then
                                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                                    End If
                                    
                                    .inicializarEntidades
                                End With
                            End If
                            
                        Else
                            .Dataset.Cancel
                        
                            Exit Sub
                        End If
                    Else
                        MsgBox "Ud. no cuenta con permisos para realizar esta acción.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                Else
                    If ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2CODTAREA", "EF2TAREAUSERS", "F2CODUSER", wusuario, "T", "AND F2CODTAREA = '0011'") = "0011" Then
                        If MsgBox("¿Desea QUITAR su VºBº al registro seleccionado?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                            .Columns.ColumnByFieldName("F4ESTADO").value = 1
                            .Columns.ColumnByFieldName("F4VB1").value = False
                            .Columns.ColumnByFieldName("F4VBUSER1").value = vbNullString
                            .Columns.ColumnByFieldName("F4VBFECHA1").value = Null
                            
                            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                                With objSqlAyudaOrden
                                    .inicializarEntidades
                                    
                                    .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
                                    .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
                                    .Estado = 1
                                    .VB1 = False
                                    .VB1Usuario = vbNullString
                                    .VB1Fecha = vbNullString
                                    
                                    If .aprobarOrden Then
                                        Actualiza_Log " < Replicación BD > " & .SQLSelectAlter, StrConexDbBancos
                                    End If
                                    
                                    .inicializarEntidades
                                End With
                            End If
                        Else
                            .Dataset.Cancel
                        
                            Exit Sub
                        End If
                    Else
                        MsgBox "Ud. no cuenta con permisos para realizar esta acción.", vbInformation + vbOKOnly, App.ProductName
                        
                        .Dataset.Cancel
                        
                        Exit Sub
                    End If
                End If
                
                .Dataset.Post
                
                With objAyudaOrden
                    .Estado = Val(dbgOrden.Columns.ColumnByFieldName("F4ESTADO").value & "")
                    .VB1 = CBool(dbgOrden.Columns.ColumnByFieldName("F4VB1").value)
                    .VB1Usuario = UCase(Trim(dbgOrden.Columns.ColumnByFieldName("F4VBUSER1").value & ""))
                    .VB1Fecha = IIf(Not IsNull(dbgOrden.Columns.ColumnByFieldName("F4VBFECHA1").value), Trim(dbgOrden.Columns.ColumnByFieldName("F4VBFECHA1").value & ""), vbNullString)
                    
                    If .aprobarOrden Then
                        Actualiza_Log .SQLSelectAlter, StrConexDbBancos
                        
                        
                    End If
                End With
            End With
    End Select
End Sub

Private Sub dbgOrden_OnClick()
    Select Case dbgOrden.Columns.FocusedColumn.FieldName
        Case "F4LOCAL", "F4NUMORD", "F4ESTADO"
            dbgDocumento_RowColChange
    End Select
End Sub

Private Sub dbgOrden_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "F4SOLES"
            If Val(Text) = 0 Then
                FontColor = &HC0FFFF
            End If
            
            Text = Format(Val(Text), "#,0.00")
        Case "F4DOLARES"
            If Val(Text) = 0 Then
                FontColor = &HC0FFC0
            End If
            
            Text = Format(Val(Text), "#,0.00")
    End Select
End Sub

Private Sub dbgOrden_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case Column.FieldName
        Case "F4SOLES", "F4DOLARES"
            Text = Format(Val(Text), "#,0.00")
    End Select
End Sub

Private Sub dbgOrden_OnDblClick()
    Me.MousePointer = vbHourglass

    If bolAyuda Then
        objAyudaOrden.TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
        objAyudaOrden.NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
        
        Unload Me
    Else
        With ordendecompra
            .Ayuda = bolAyuda
            
            .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
            .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
            
            .Show
            
            Me.Hide
            
        End With
    End If

    Me.MousePointer = vbDefault
End Sub

Private Sub dbgOrden_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    Select Case KeyCode
        Case vbKeyReturn
            dbgOrden_OnDblClick
        Case vbKeyUp
            If dbgOrden.Dataset.RecNo = 1 Then
                txtBusqueda.SetFocus
            End If
    End Select
End Sub

Private Sub dbgOrden_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, TipText As String, l As Single, t As Single, R As Single, b As Single, NeedShowTip As Boolean)
    Select Case Column.FieldName
        Case "F2NOMPROV", "F4OBSERVA"
            NeedShowTip = True
        Case "F4ESTADO"
            NeedShowTip = True
            
            Select Case Val(TipText & "")
                Case 1
                    TipText = "Orden en Edición"
                Case 2
                    TipText = "Orden Aprobada"
                Case 3
                    TipText = "Orden Enviada"
                Case 4
                    TipText = "Orden Recepcionada"
                Case 5
                    TipText = "Atención Parcial"
                Case 6
                    TipText = "Atención Total"
                Case 7
                    TipText = "Orden Cerrada"
                Case 8
                    TipText = "Orden Anulada"
            End Select
        Case Else
            NeedShowTip = False
    End Select
End Sub

Private Sub dbgOrdenDetalle_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
    Select Case UCase(Column.FieldName)
        Case "COD_SOLICITUD"
            If Trim(Text) <> "STOCK LIBRE" Then
                Font.Bold = True
                FontColor = vbWhite
                Color = RGB(79, 129, 189)
            End If
        Case "F3CANPRO"
            Text = Format(Text, "#,0.00;(#,0.00)")
        Case "F3CANFAL"
            If Val(Text) > 0 Then
                Font.Bold = True
                FontColor = vbRed
                Color = vbYellow
            End If
            
            Text = Format(Text, "#,0.00;(#,0.00)")
    End Select
End Sub

Private Sub Form_Load()
    If Not bolAyuda Then
        Me.top = 1000
        Me.left = 1250
    End If
    
    listarTipoOrden
    
    listarOrden
    
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
End Sub

Public Sub listarOrden()
    Screen.MousePointer = vbHourglass
    
    Select Case Trim(right(tlbOrden.Tools.ITEM("TipoOrden").ComboBox.Text, 2))
        Case "OC"
            tlbOrden.Tools.ITEM("NuevaOC").Visible = True
            tlbOrden.Tools.ITEM("NuevaOS").Visible = False
            tlbOrden.Tools.ITEM("NuevaOF").Visible = False
        Case "OS"
            tlbOrden.Tools.ITEM("NuevaOC").Visible = False
            tlbOrden.Tools.ITEM("NuevaOS").Visible = True
            tlbOrden.Tools.ITEM("NuevaOF").Visible = False
            tlbOrden.Tools.ITEM("Movimiento").Visible = False
        Case "OF"
            tlbOrden.Tools.ITEM("NuevaOC").Visible = False
            tlbOrden.Tools.ITEM("NuevaOS").Visible = False
            tlbOrden.Tools.ITEM("NuevaOF").Visible = True
            tlbOrden.Tools.ITEM("Movimiento").Visible = False
        Case Else
            tlbOrden.Tools.ITEM("NuevaOC").Visible = True
            tlbOrden.Tools.ITEM("NuevaOS").Visible = True
            tlbOrden.Tools.ITEM("Movimiento").Visible = True
    End Select
    
    dbgOrden.Dataset.Close
    
    objAyudaOrden.listarGrillaOrden dbgOrden, Trim(right(tlbOrden.Tools.ITEM("TipoOrden").ComboBox.Text, 2)), txtBusqueda.Text, False, imgLstEstado
    
    dbgDocumento_RowColChange
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub listarOrdenDetalle()
    With dbgOrdenDetalle
        .Dataset.Active = False
        .Dataset.ADODataset.ConnectionString = StrConexDbBancos
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "IIF(DET.F4LOCAL = 'OC', IIF(DET.COD_SOLICITUD <> '', DET.COD_SOLICITUD & ' / ' & REQ.CS_NOMREF, 'STOCK LIBRE'), DET.COD_SOLICITUD) AS COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "DET.F3CODFAB, "
        SqlCad = SqlCad & "DET.F5NOMPRO_ING AS F5NOMPRO, "
        SqlCad = SqlCad & "MED.F7SIGMED, "
        SqlCad = SqlCad & "IIF(DET.F4LOCAL = 'OC', (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))), DET.F3CANPRO) AS F3CANPRO, "
        SqlCad = SqlCad & "IIF(DET.F4LOCAL = 'OC', VAL(FORMAT( (DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100))) - (VAL(INGRESOS.CANTIDAD & '') / IIF(ISNULL(MEDALTER.F5FACTOR), 1, MEDALTER.F5FACTOR)) , '#0.0000')), 0) AS F3CANFAL, "
        SqlCad = SqlCad & "DET.F3PRENETO AS F3PRECOS "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(((IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD) "
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        
        SqlCad = SqlCad & "SUM(DET.F3CANPRO * IIF(TIPO = 'S', -1, 1)) AS CANTIDAD "
        
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "") & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
        
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = DET.F4NUMORD AND INGRESOS.COD_SOLICITUD = DET.COD_SOLICITUD AND INGRESOS.F5CODPROORIGINAL = DET.F3CODPRO "
        
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4LOCAL = '" & Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "") & "' AND "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "") & "' "
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "DET.F5NOMPRO_ING"
        
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        
        .KeyField = "F3CODPRO"
    End With
    
    SqlCad = vbNullString
End Sub

Public Sub listarOrdenDetalleSql()
    With dbgOrdenDetalle
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "OCS.COD_SOLICITUD, "
        SqlCad = SqlCad & "OCS.CODPRODUCTO AS F3CODPRO, "
        SqlCad = SqlCad & "OCS.F3CODFAB, "
        SqlCad = SqlCad & "OCS.F5NOMPRO, "
        SqlCad = SqlCad & "OCS.F7SIGMED, "
        SqlCad = SqlCad & "OCS.F3CANPRO, "
        SqlCad = SqlCad & "OCS.FACTOR, "
        SqlCad = SqlCad & "CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) AS INGR, "
        SqlCad = SqlCad & "(CASE WHEN OCS.F3CANPRO < CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) THEN 0 ELSE OCS.F3CANPRO - CONVERT(DECIMAL(10, 2), ISNULL(INGRESOS.CANTIDAD, 0) / ISNULL(OCS.FACTOR, 1) ) END) AS F3CANFAL, "
        SqlCad = SqlCad & "OCS.F3PRECOS "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD AS OC, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD AS NROPEDIDO, "
        SqlCad = SqlCad & "(CASE WHEN DET.F4LOCAL = 'OC' THEN (CASE WHEN DET.COD_SOLICITUD <> '' THEN DET.COD_SOLICITUD + ' / ' + REQ.CS_NOMREF ELSE 'STOCK LIBRE' END) ELSE DET.COD_SOLICITUD END) AS COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F3CODPRO AS CODPRODUCTO, "
        SqlCad = SqlCad & "DET.F3CODFAB, "
        SqlCad = SqlCad & "DET.F5NOMPRO_ING AS F5NOMPRO, "
        SqlCad = SqlCad & "MED.F7SIGMED, "
        SqlCad = SqlCad & "SUM(CASE WHEN DET.F4LOCAL = 'OC' THEN "
        SqlCad = SqlCad & "DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) "
        SqlCad = SqlCad & "ELSE "
        SqlCad = SqlCad & "DET.F3CANPRO "
        SqlCad = SqlCad & "END) AS F3CANPRO, "
        SqlCad = SqlCad & "ISNULL(MEDALTER.F5FACTOR, 1) AS FACTOR, "
        SqlCad = SqlCad & "DET.F3PRENETO AS F3PRECOS "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "PROCESOS.IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN MAESTROS.EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD "
        SqlCad = SqlCad & "LEFT JOIN MAESTROS.MEDIVENTAS AS MEDALTER ON MEDALTER.F5CODPRO = DET.F3CODPRO AND MEDALTER.F7CODMED = DET.UNIDAD "
        SqlCad = SqlCad & "LEFT JOIN PROCESOS.TB_CABSOLICITUD AS REQ ON REQ.COD_SOLICITUD = DET.COD_SOLICITUD "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4LOCAL = '" & Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "") & "' AND "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "") & "' AND "
        SqlCad = SqlCad & "(CASE WHEN DET.F4LOCAL = 'OC' THEN "
        SqlCad = SqlCad & "DET.F3CANPRO * (1 + (DET.F3PORCDEMASIA/100)) "
        SqlCad = SqlCad & "ELSE "
        SqlCad = SqlCad & "DET.F3CANPRO "
        SqlCad = SqlCad & "END) > 0 "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "(CASE WHEN DET.F4LOCAL = 'OC' THEN (CASE WHEN DET.COD_SOLICITUD <> '' THEN DET.COD_SOLICITUD + ' / ' + REQ.CS_NOMREF ELSE 'STOCK LIBRE' END) ELSE DET.COD_SOLICITUD END), "
        SqlCad = SqlCad & "DET.F3CODPRO, "
        SqlCad = SqlCad & "DET.F3CODFAB, "
        SqlCad = SqlCad & "DET.F5NOMPRO_ING, "
        SqlCad = SqlCad & "MED.F7SIGMED, "
        SqlCad = SqlCad & "ISNULL(MEDALTER.F5FACTOR, 1), "
        SqlCad = SqlCad & "DET.F3PRENETO) AS OCS "
        
        SqlCad = SqlCad & "LEFT JOIN "
        SqlCad = SqlCad & "(SELECT "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL, "
        SqlCad = SqlCad & "CONVERT(DECIMAL(10,2), SUM(DET.F3CANPRO * (CASE WHEN DET.TIPO = 'I' THEN 1 ELSE -1 END)) ) AS CANTIDAD "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "PROCESOS.IF3VALES AS DET "
        SqlCad = SqlCad & "LEFT JOIN PROCESOS.IF4VALES AS CAB ON CAB.F4NUMVAL = DET.F4NUMVAL AND CAB.F2CODALM = DET.F2CODALM "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "CAB.F1CODORI IN ('XC0') AND "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "") & "' "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F4NUMORD, "
        SqlCad = SqlCad & "DET.COD_SOLICITUD, "
        SqlCad = SqlCad & "DET.F5CODPROORIGINAL) AS INGRESOS "
        SqlCad = SqlCad & "ON INGRESOS.F4NUMORD = OCS.OC AND INGRESOS.COD_SOLICITUD = OCS.NROPEDIDO AND INGRESOS.F5CODPROORIGINAL = OCS.CODPRODUCTO "
        
        SqlCad = SqlCad & "ORDER BY "
        SqlCad = SqlCad & "OCS.F5NOMPRO"
        
        
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = strCadenaConexioBdCPlus
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandType = cmdText
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.ADODataset.CursorType = ctStatic
        .Dataset.ADODataset.LockType = ltReadOnly
        .Dataset.ADODataset.CommandText = SqlCad
        .Dataset.Active = True
        .Dataset.Refresh
        .KeyField = "F3CODPRO"
    End With
    
    SqlCad = vbNullString
End Sub

Private Sub imprimeOrdenV2(ByVal strTipoOrden As String, _
                            ByVal strNumeroOrden As String, _
                            Optional ByVal imprimirPDFparaEnvioMail As Boolean)
    
    On Error GoTo errImprimeOrdenV2
    
    Dim nAnchoHoja As Double
    
    Dim rpt As New Acr_OrdenCompra
    
    With objAyudaOrden
        .TipoOrden = strTipoOrden
        .NumeroOrden = strNumeroOrden
        
        .obtenerConfigOrden
    End With
    
    With rpt
        .DescargarReporte = imprimirPDFparaEnvioMail
        
        If imprimirPDFparaEnvioMail Then
            If Dir(wrutatemp & "\ParaAtencionDeOrden.pdf", vbArchive) <> vbNullString Then
                Kill wrutatemp & "\ParaAtencionDeOrden.pdf"
            End If
            
            .TipoOrden = strTipoOrden
            .NumeroOrden = strNumeroOrden
        End If
        
        Select Case objAyudaOrden.CodMoneda
            Case "S"
                .LblTotF.Caption = "Total " & "S/"
            Case "D"
                .LblTotF.Caption = "Total " & "US$"
        End Select
        
        .flddirec1.Text = wf1direc1
        .FldTelf.Text = "Teléfono: " & wtelefono & " // Fax: " & wfax
        .LblCentroCosto.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F3DESCRIP", "CENTROS", "F3COSTO", objAyudaOrden.CentroCosto, "T")
        
        nAnchoHoja = (.PageSettings.PaperWidth - .PageSettings.LeftMargin - .PageSettings.RightMargin)
        .fldruc.Text = "R.U.C. " & wrucempresa
        
        If FileExist(App.Path & "\Logo" & left(wempresa, 5) & ".bmp") = True Then
'            .fldempresa.Visible = False
            With cImgInfo
                .ReadImageInfo App.Path & "\Logo" & left(wempresa, 5) & ".bmp"
            End With
        Else
'            .fldempresa.Visible = True
'            .fldempresa.Text = wnomcia
        End If
        
            
        Select Case objAyudaOrden.TipoOrden
            Case "OC"
                .LblTitle.Caption = "ORDEN DE COMPRA"
            Case "OS"
'                .Field334.Text = "Formato LOG-F-08"
        End Select
        
        .LblNroOC.Caption = "N° " & objAyudaOrden.NumeroOrden
        
        .fldsolicitud.Text = objAyudaOrden.generarCadenaSolicitud
        
        With objAyudaProveedor
            .Codigo = objAyudaOrden.CodProveedor
            
            .obtenerProveedor
        End With
        
        .F2NOMPROV.Text = objAyudaProveedor.NombreProveedor
        .F2DIRPROV.Text = objAyudaProveedor.DireccionProveedor
        .F2CONTACTO.Text = objAyudaProveedor.Contacto
        .F2CONTACTO.Visible = True
        .F2TELPROV.Text = objAyudaProveedor.Telefono
        .F2FAXPROV.Text = objAyudaProveedor.Fax
        
        .F4FECEMI.Text = objAyudaOrden.FechaEmision
        .FldFchEntrega = objAyudaOrden.FechaEntrega
        .FldTipCam.Text = Format(objAyudaOrden.TipoCambio, "0.000")
        
        .FldTipDoc.Text = UCase(ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESDOC", "DOCUMENTOS", "F2CODDOC", objAyudaOrden.CodTipoComprobante, "T"))
        
        Select Case objAyudaOrden.CodTipoComprobante
            Case "02"
                .LblImp.Caption = "Reten."
            Case Else
                .LblImp.Caption = "I.G.V."
        End Select
        
        .FldObservaAll.Text = objAyudaOrden.Observacion
        .F4CODPRV.Text = objAyudaOrden.RucProveedor
        .FldSon.Text = CADENANUM(objAyudaOrden.TotalFacturado, objAyudaOrden.CodMoneda, vbNullString)
        
        .F2DESPAG.Text = ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F2DESPAG", "EF2FORPAG", "F2FORPAG", objAyudaOrden.CodFormaPago, "T")
                        
        .F4COTIZACION.Text = objAyudaOrden.NumeroCotizacion
        .REMITIR.Text = objAyudaOrden.LugarEntrega
        
        .DataControl1.ConnectionString = cnn_dbbancos
        
        SqlCad = vbNullString
        SqlCad = SqlCad & "SELECT "
        SqlCad = SqlCad & "DET.F3CODFAB AS F3CODPRO, "
        SqlCad = SqlCad & "DET.F5NOMPRO, "
        SqlCad = SqlCad & "COLOR.DESCRIPCION AS DESCOLOR, "
        SqlCad = SqlCad & "MED.F7SIGMED AS F3MEDIDA, "
        SqlCad = SqlCad & "SUM(DET.F3CANPRO * (1 + (F3PORCDEMASIA/100))) AS F3CANPRO, "
        SqlCad = SqlCad & "VAL(FORMAT(DET.F3PRENETO, '#0.00')) AS F3PREUNI, "
        SqlCad = SqlCad & "SUM(DET.F5VALVTA) AS F5VALVTA, "
        SqlCad = SqlCad & "SUM(DET.F3IGV) AS F3IGV, "
        SqlCad = SqlCad & "SUM(DET.F3PORDCT) AS F3PORDCT, "
        SqlCad = SqlCad & "SUM(DET.F3TOTAL) AS F3TOTAL "
        SqlCad = SqlCad & "FROM "
        SqlCad = SqlCad & "((IF3ORDEN AS DET "
        SqlCad = SqlCad & "LEFT JOIN CENTROS AS CC ON CC.F3COSTO = DET.F3CENCOS) "
        SqlCad = SqlCad & "LEFT JOIN EF7MEDIDAS AS MED ON MED.F7CODMED = DET.UNIDAD) "
        SqlCad = SqlCad & "LEFT JOIN EF2BIENCOLOR AS COLOR ON COLOR.CODIGO = DET.CODCOLOR "
        SqlCad = SqlCad & "WHERE "
        SqlCad = SqlCad & "DET.F4NUMORD = '" & strNumeroOrden & "' AND "
        SqlCad = SqlCad & "DET.F4LOCAL = '" & strTipoOrden & "' AND "
        SqlCad = SqlCad & "DET.F3CANPRO <> 0 "
        SqlCad = SqlCad & "GROUP BY "
        SqlCad = SqlCad & "DET.F3CODFAB,  "
        SqlCad = SqlCad & "DET.F5NOMPRO, "
        SqlCad = SqlCad & "COLOR.DESCRIPCION, "
        SqlCad = SqlCad & "MED.F7SIGMED, "
        SqlCad = SqlCad & "VAL(FORMAT(DET.F3PRENETO, '#0.00'))"
        
        .DataControl1.Source = SqlCad
        
        .Caption = "ORDEN NACIONAL"
        
        .Show
    End With
    
    Exit Sub
errImprimeOrdenV2:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub verificarEstadoOC()
    On Error GoTo errVerificarEstadoOC
    
    Me.MousePointer = vbHourglass
    
    With objAyudaOrden
        .inicializarEntidades
        
        .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
        .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
        
        If .obtenerOrden Then
            If .Estado <> 7 And .Estado <> 8 Then
                For d = 0 To 25
                    nSaveRecNo = dbgOrden.Dataset.RecNo
                Next
                
                dbgOrdenDetalle.Dataset.Close
                dbgOrden.Dataset.Close
                
                abrirCnnDbBancos
                
                .atencionOrden
                
                abrirCnnDbBancos
                
                listarOrden
                
                If dbgOrden.Dataset.RecordCount >= nSaveRecNo Then
                    dbgOrden.Dataset.RecNo = nSaveRecNo
                End If
            End If
        End If
    End With
    
    Me.MousePointer = vbDefault
    
    Exit Sub
errVerificarEstadoOC:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub verificarEstadoOCSql()
    On Error GoTo errVerificarEstadoOCSql
    
    Me.MousePointer = vbHourglass
    
    With objSqlAyudaOrden
        .inicializarEntidades
        
        .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
        .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
        
        If .obtenerOrden Then
            If .Estado <> 7 And .Estado <> 8 Then
                For d = 0 To 25
                    nSaveRecNo = dbgOrden.Dataset.RecNo
                Next
                
                dbgOrdenDetalle.Dataset.Close
                dbgOrden.Dataset.Close
                
                .atencionOrden
                
                abrirCnnDbBancos
                
                SqlCad = vbNullString
                SqlCad = SqlCad & "UPDATE "
                    SqlCad = SqlCad & "IF4ORDEN "
                SqlCad = SqlCad & "SET "
                    SqlCad = SqlCad & "F4ESTADO = " & .Estado & " "
                SqlCad = SqlCad & "WHERE "
                    SqlCad = SqlCad & "F4LOCAL = '" & .TipoOrden & "' AND "
                    SqlCad = SqlCad & "F4NUMORD = '" & .NumeroOrden & "'"
                
                cnn_dbbancos.Execute SqlCad
                
                Actualiza_Log SqlCad, StrConexDbBancos
                
                abrirCnnDbBancos
                
                listarOrden
                
                If dbgOrden.Dataset.RecordCount >= nSaveRecNo Then
                    dbgOrden.Dataset.RecNo = nSaveRecNo
                End If
            End If
        End If
    End With
    
    Me.MousePointer = vbDefault
    
    Exit Sub
errVerificarEstadoOCSql:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
    
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bolAyuda = False
End Sub

Private Sub tlbOrden_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "TipoOrden"
            listarOrden
    End Select
End Sub

Private Sub tlbOrden_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "NuevaOC"
            
            With ordendecompra
                .Ayuda = bolAyuda
                
                .TipoOrden = "OC"
                .NumeroOrden = vbNullString
                
                .Show
                
                Me.Hide
            End With
        Case "NuevaOS"
            With ordendecompra
                .Ayuda = bolAyuda
                
                .TipoOrden = "OS"
                .NumeroOrden = vbNullString
                
                .Show
                
                Me.Hide
            End With
        Case "Filtrar"
            If Tool.State = ssChecked Then
                dbgOrden.Filter.FilterActive = True
            Else
                dbgOrden.Filter.FilterActive = False
            End If
        Case "Agrupar"
            If Tool.State = ssChecked Then
                dbgOrden.Options.Set (egoShowGroupPanel)
            Else
                dbgOrden.Options.Unset (egoShowGroupPanel)
            End If
        Case "Excel"
            Screen.MousePointer = vbHourglass
            
            With cmdlgOrden
                .DialogTitle = "Guardar como..."
                .Filter = "Archivos de MS Excel | *.xls"
                .FileName = vbNullString
                
                .ShowSave
                
                If .FileName <> vbNullString Then
                    dbgOrden.M.ExportToXLS .FileName
                    
                    If Dir(.FileName) <> vbNullString Then
                        MsgBox "Exportación terminada.", vbInformation, App.ProductName
                    Else
                        MsgBox "Exportación fallida.", vbInformation, App.ProductName
                    End If
                End If
            End With
            
            Screen.MousePointer = vbDefault
        Case "Movimiento"
            If Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "") <> "OC" Then
                MsgBox "Opción disponible solo para Ordenes de Compra.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            Dim rpt As New rptOrdenMovimiento
            
            With rpt
                .TipoOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "")
                .NumeroOrden = Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
                
                .Show
            End With
        Case "Imprimir"
            imprimeOrdenV2 Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & ""), Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "")
        Case "VerificarEstado"
            If Trim(dbgOrden.Columns.ColumnByFieldName("F4LOCAL").value & "") <> "OC" Then
                MsgBox "Opción disponible solo para Ordenes de Compra.", vbInformation + vbOKOnly, App.ProductName
                
                Exit Sub
            End If
            
            If MsgBox("¿Desea ejecutar el proceso de Verificación de Estado de Orden N° " & Trim(dbgOrden.Columns.ColumnByFieldName("F4NUMORD").value & "") & "?" & vbNewLine & _
                        "ATENCIÓN: Esto puede tomar algunos minutos.", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
                
                Exit Sub
            End If
            
            If sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "UsarReplicacionBdCPlus", "l") = "1" Then
                verificarEstadoOCSql
            Else
                verificarEstadoOC
            End If
        Case "Salir"
            objAyudaOrden.inicializarEntidades
            
            Unload Me
    End Select
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Dim dblPromH As Double
    
    dblPromH = Me.ScaleHeight / 6
    
    
    FraBusqueda.Move 0, 0, Me.ScaleWidth, 815
    
    dbgOrden.Move 0, FraBusqueda.Height, Me.ScaleWidth, (dblPromH * 4) - 400
    
    dbgOrdenDetalle.Move 0, FraBusqueda.Height + dbgOrden.Height, Me.ScaleWidth, (dblPromH * 2) - 400
    
    FraBusqueda.Width = dbgOrden.Width
    txtBusqueda.Width = dbgOrden.Width - 350
End Sub

Private Sub txtbusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            listarOrden
        Case vbKeyDown
            dbgOrden.SetFocus
    End Select
End Sub


