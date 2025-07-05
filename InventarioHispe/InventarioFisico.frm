VERSION 5.00
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form InventarioFisico 
   Caption         =   "Toma de Inventario Físico"
   ClientHeight    =   7155
   ClientLeft      =   1185
   ClientTop       =   1215
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   10680
   Begin VB.Frame Frame1 
      Height          =   6630
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   10476
      Begin VB.CheckBox ChkVSF 
         Caption         =   "Visualizar Stock Fisico"
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   1050
         Width           =   2325
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5055
         Left            =   165
         OleObjectBlob   =   "InventarioFisico.frx":0000
         TabIndex        =   12
         Top             =   1440
         Width           =   10140
      End
      Begin VB.TextBox Txtcodalm 
         Height          =   285
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   0
         Top             =   225
         Width           =   465
      End
      Begin aBoxCtl.aBox abofecha 
         Height          =   315
         Left            =   1125
         TabIndex        =   1
         Top             =   630
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
         Text            =   "16/07/2007"
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
         ButtonPicture   =   "InventarioFisico.frx":0C7B
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
      Begin aBoxCtl.aBox AboInventa 
         Height          =   315
         Left            =   9000
         TabIndex        =   2
         Top             =   630
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
         Locked          =   -1  'True
         ApplyTextFormat =   -1  'True
         TextFormat      =   "dd/mm/yyyy"
         Text            =   "16/07/2007"
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
         ButtonPicture   =   "InventarioFisico.frx":0FCD
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
      Begin Threed.SSPanel PnlNomalm 
         Height          =   285
         Left            =   1635
         TabIndex        =   4
         Top             =   225
         Width           =   8700
         _Version        =   65536
         _ExtentX        =   15346
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
         Left            =   1125
         TabIndex        =   8
         Top             =   1080
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   16761024
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
      Begin Threed.SSPanel Txtnumvald 
         Height          =   285
         Left            =   9000
         TabIndex        =   9
         Top             =   1080
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   -2147483640
         BackColor       =   16761024
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Vale Salida:"
         DataField       =   "<"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1080
         Width           =   840
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
         Left            =   7830
         TabIndex        =   10
         Top             =   1125
         Width           =   930
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fec. Ult. Inventario :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   7335
         TabIndex        =   7
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha    :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Almacén :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   345
         TabIndex        =   5
         Top             =   270
         Width           =   705
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "InventarioFisico.frx":131F
      ToolBars        =   "InventarioFisico.frx":5F6D
   End
End
Attribute VB_Name = "InventarioFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTOMADET As New ADODB.Recordset
Dim RsTomaCab   As New ADODB.Recordset
Dim sw_diferencia As Boolean
Dim wcierre As String
Dim avales_cab(0 To 11) As a_grabacion
Dim avales_det(0 To 9) As a_grabacion
Dim amovs_det(0 To 8) As a_grabacion
Dim Values()            As Variant
Dim sw_ayuda_prod       As Boolean

Sub llenar_inv()
If Len(Trim(wcod_alm)) > 0 Then
    Txtcodalm.Text = wcod_alm
    PnlNomalm.Caption = wnomalmacen
    AboInventa.Value = Format(wultinv, "DD/MM/YYYY")
    Txtcodalm_KeyPress 13
End If
End Sub


Private Sub ChkVSF_Click()
Call Llena_Grid
End Sub

Private Sub dxDBGrid1_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Select Case dxDBGrid1.Columns.FocusedAbsoluteIndex
    Case 0:
        wcodproducto = ""
        sw_ayuda_prod = True
        hlp_productos.Show 1
        If Len(Trim(wcodproducto)) > 0 Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
            dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
            dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
            dxDBGrid1.Columns.ColumnByFieldName("F7CODMED").Value = wmedida
            If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
            RsMovAlmacen.Open "Select F6STOCKACT,F5COSPRO from IF6ALMA where F2CODALM = '" & wcod_alm & "' AND F5CODPRO = '" & wcodproducto & "'", cnn_dbbancos, adOpenDynamic
            If Not RsMovAlmacen.EOF Then
                dxDBGrid1.Columns(4).Value = Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.00000")
                dxDBGrid1.Columns(5).Value = Format(0, "###,##0.00000")
                dxDBGrid1.Columns(7).Value = Format(RsMovAlmacen.Fields("F5COSPRO"), "###,##0.00")
                dxDBGrid1.Columns(6).Value = Format(0, "###,##0.00000")
                dxDBGrid1.Columns(9).Value = Format(0, "###,##0.00")
            End If
            RsMovAlmacen.Close
            dxDBGrid1.Dataset.Post
            dxDBGrid1.Columns.FocusedIndex = 4
        End If
End Select
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

If dxDBGrid1.Columns.FocusedIndex = 4 Then
    dxDBGrid1.Dataset.Edit
    calcula
    If dxDBGrid1.Dataset.State = dsEdit Or dsInsert Then
       dxDBGrid1.Dataset.Post
       dxDBGrid1.Dataset.Refresh
    End If
End If
End Sub

Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
    If KeyCode = 115 Then
        If MsgBox("¿Desea Eliminar el registro Actual? ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            If dxDBGrid1.Dataset.RecNo = 1 Then
                dxDBGrid1.Dataset.Delete
                AdicionaItem
            Else
                dxDBGrid1.Dataset.Delete
                If dxDBGrid1.Dataset.RecordCount = 0 Then AdicionaItem
            End If
            sw_nuevo_item = False
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Left = 1625
Me.Top = 1220
Me.Height = 7560
Me.Width = 10815

'BASE_TEMPORAL "TEMPFAC.MDB"
'DELETEREC_N "TMPINVENTARIO", Temp
'dxDBGrid1.Dataset.Refresh
'Conf_Grid

SSActiveToolBars1.Tools("ID_CerrarInventario").Enabled = False
SSActiveToolBars1.Tools("ID_Grabar").Enabled = False
SSActiveToolBars1.Tools("ID_Imprimir").Enabled = False
SSActiveToolBars1.Tools("ID_ImprimirDetalle").Enabled = False

abofecha.Value = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Select Case Tool.ID
    Case "ID_Procesar"
        Me.MousePointer = 11
        Actualiza_Grid
        Me.MousePointer = 1
        SSActiveToolBars1.Tools("ID_Grabar").Enabled = True
        SSActiveToolBars1.Tools("ID_Imprimir").Enabled = True
        SSActiveToolBars1.Tools("ID_ImprimirDetalle").Enabled = True
    Case "ID_Grabar"
        Me.MousePointer = 11
        If Val(wultinv) = CDate("12:00:00") Then
            wultinv = abofecha.Value
        End If

        grabar
        Me.MousePointer = 1
    Case "ID_CerrarInventario"
        If MsgBox("El Sistema Bloqueará las modificaciones de los Vales Emitidos hasta el " & Format(abofecha.Value, "dd/mm/yyyy") & ", Está seguro de realizar el Cierre de Inventario", 36, "Cierre de Inventario") = 6 Then
            Me.MousePointer = 11
            CERRAR
            dxDBGrid1.Enabled = False
            SSActiveToolBars1.Tools("ID_CerrarInventario").Enabled = False
            Me.MousePointer = 1
        End If
    Case "ID_Imprimir"
        Me.MousePointer = 11
        sw_diferencia = False
        'Imprimir_Saldo
        Me.MousePointer = 1
    Case "ID_ImprimirDetalle"
        Me.MousePointer = 11
        sw_diferencia = True
        'Imprimir_Saldo
        Me.MousePointer = 1
    Case "ID_Salir"
        Unload Me
End Select
End Sub

Private Sub Txtcodalm_Change()
If Trim(Txtcodalm.Text) <> "" And sw_cabecera = False Then sw_cabecera = True
End Sub

Private Sub Txtcodalm_DblClick()
Txtcodalm_KeyDown 113, 0
End Sub

Private Sub txtcodalm_GotFocus()
Txtcodalm.SelStart = 0: Txtcodalm.SelLength = Len(Txtcodalm.Text)
End Sub

Public Sub Txtcodalm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        wcod_alm = ""
        ayuda_almacen.Show 1
        sw_ayuda = False
        Call llenar_inv
    End If
    
End Sub

Private Sub Txtcodalm_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then abofecha.SetFocus
End Sub

Private Sub txtcodalm_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(Txtcodalm.Text)) > 0 Then
            wnomalmacen = ""
            
            If Txtcodalm.Text <> "" Then
                If VALIDA_ALMACEN(Txtcodalm.Text) = True Then
                    wcod_alm = Txtcodalm.Text
                    PnlNomalm.Caption = wnomalmacen
                    AboInventa.Value = Format(wultinv, "DD/MM/YYYY")
                    
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

Public Sub Actualiza_Grid()
On Error GoTo HNDERR
Set RsTomaCab = New ADODB.Recordset

'sql = "SELECT * FROM H4TOMAINV WHERE CVDATE(F4FECTOM) = '" & CVDate(wultinv) & "' AND F2CODALM='" & wcod_alm & "'"
'If RsTomaCab.State = adStateOpen Then RsTomaCab.Close
'RsTomaCab.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'If Not RsTomaCab.EOF Then
'    sw_nuevo_doc = False
'    Mostrar_Datos
'    If RsTomaCab.Fields("F4cierre") = "1" Then
'        Txtnumvalo.Caption = RsTomaCab.Fields("F4NUMVALS")
'        Txtnumvald.Caption = RsTomaCab.Fields("F4NUMVALI")
'        SSActiveToolBars1.Tools("ID_CerrarInventario").Enabled = False
'        dxDBGrid1.Enabled = False
'        Txtnumvalo.Visible = True
'        Txtnumvald.Visible = True
'        Label1(0).Visible = True
'        Label1(1).Visible = True
'    Else
        SSActiveToolBars1.Tools("ID_CerrarInventario").Enabled = True
        dxDBGrid1.Enabled = True
        Txtnumvalo.Visible = False
        Txtnumvald.Visible = False
        Label1(0).Visible = False
        Label1(1).Visible = False
 '   End If
        
'Else
    sw_nuevo_doc = True
    dxDBGrid1.Enabled = True
    Llena_Grid
'End If
Exit Sub
HNDERR:
MsgBox "Ha Ocurrido el Siguiente Error: " & Err.Description, vbExclamation, "Sistema de Logística"

Exit Sub
End Sub

Private Sub Mostrar_Datos()
Dim var As Integer
Set RsProducto = New Recordset
dxDBGrid1.Dataset.Close
BASE_TEMPORAL "TEMPFAC.MDB"
DELETEREC_N "TMPINVENTARIO", Temp
If RSTOMADET.State = adStateOpen Then RSTOMADET.Close
RSTOMADET.Open "SELECT F5CODPRO,F3STOCKSIS,F3STOCKFIS,F4FECTOM,F2CODALM,F3DIFERENCIA,F3COSTOSIS,F3COSTOACT FROM H3TOMAINV WHERE CVDATE(F4FECTOM) = '" & CVDate(wultinv) & "' AND F2CODALM = '" & wcod_alm & "' ORDER BY F5CODPRO ", cnn_dbbancos, adOpenDynamic
If Not RSTOMADET.EOF Then
    With dxDBGrid1.Dataset
        Do While Not RSTOMADET.EOF
            
            If RsProducto.State = adStateOpen Then RsProducto.Close
            RsProducto.Open "SELECT F5CODFAB,F5NOMPRO,F7CODMED FROM IF5PLA WHERE F5CODPRO = '" & RSTOMADET.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic
            If Not RsProducto.EOF Then
                SSQL = "INSERT INTO TMPINVENTARIO(F3ITEM,F5CODPRO,F5CODFAB,F5NOMPRO,F3STOCKSIS,F7CODMED,F3STOCKFIS,F3DIFERENCIA,F3COSTOSIS,F3COSTOACT) VALUES(" & var & ",'" & RSTOMADET.Fields("F5CODPRO") & "', " & _
                       "'" & IIf(IsNull(RsProducto.Fields("F5CODFAB")), "", RsProducto.Fields("F5CODFAB")) & "','" & Left(RsProducto.Fields("F5NOMPRO"), 50) & "'," & Format(RSTOMADET.Fields("F3STOCKSIS"), "0.00000") & ", " & _
                       "'" & Trim(RsProducto.Fields("F7CODMED")) & "'," & RSTOMADET.Fields("F3STOCKFIS") & "," & RSTOMADET.Fields("F3DIFERENCIA") & "," & RSTOMADET.Fields("F3COSTOSIS") & "," & RSTOMADET.Fields("F3COSTOACT") & ")"
                Temp.Execute (SSQL)
                var = var + 1
            End If
            RsProducto.Close
            RSTOMADET.MoveNext
            If RSTOMADET.EOF Then Exit Do
        Loop
    End With
    dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Dataset.Open
    dxDBGrid1.Dataset.First
    dxDBGrid1.Columns.FocusedIndex = 0
End If
RSTOMADET.Close
End Sub

Private Sub AdicionaItem()
Dim I As Integer
Dim sw_nuevo_temp   As Boolean

'dxDBGrid1.Dataset.Active = False

If sw_nuevo_doc = False Then
    BASE_TEMPORAL "TEMPFAC.MDB"
    DELETEREC_N "TMPINVENTARIO", Temp
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Refresh
End If

dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
dxDBGrid1.Dataset.Active = True
dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open

With dxDBGrid1.Dataset

sw_nuevo_temp = False
sw_nuevo_item = True
For I = 1 To 1

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

    .FieldValues("F3ITEM") = I
    .FieldValues("F5CODPRO") = ""
    .FieldValues("F5NOMPRO") = ""
    .FieldValues("F7SIGMED") = ""
    .FieldValues("F5STOCKSIST") = Format(0, "###,##0.00")
    .FieldValues("F5STOCKREAL") = Format(0, "###,##0.00")
    .FieldValues("F5DIFERENCIA") = Format(0, "###,##0.00")
    .FieldValues("F5COSTOSIST") = Format(0, "###,##0.00")
    .FieldValues("F5COSTOACT") = Format(0, "###,##0.00")
    .FieldValues("F5CODFAB") = ""
    
Next
    
    .Post
    sw_nuevo_item = False

End With

dxDBGrid1.Dataset.Close
dxDBGrid1.Dataset.Open
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
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
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

dxDBGrid1.Columns.ColumnByFieldName("F3ITEM").Visible = False
dxDBGrid1.Columns.ColumnByFieldName("F5CODFAB").Visible = False
       
End Sub

Private Sub Llena_Grid()
'Dim i, var As Integer
'Dim csql, SSQL As String
'Dim Dif As Double
'
'Set RsMovAlmacen = New ADODB.Recordset
'Set RsProducto = New ADODB.Recordset
'dxDBGrid1.Dataset.Close
'var = 1
'sql = "SELECT F5CODPRO,F6STOCKACT,F5COSPRO FROM IF6ALMA WHERE F2CODALM='" & wcod_alm & "' ORDER BY F5CODPRO"
'If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
'RsMovAlmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'If Not RsMovAlmacen.EOF Then
'    RsMovAlmacen.MoveFirst
'    BASE_TEMPORAL "TEMPFAC.MDB"
'    DELETEREC_N "TMPINVENTARIO", Temp
'    With dxDBGrid1.Dataset
'        Do While Not RsMovAlmacen.EOF
'            If RsProducto.State = adStateOpen Then RsProducto.Close
'            RsProducto.Open "SELECT F5CODFAB,F5NOMPRO,F7CODMED FROM IF5PLA WHERE F5CODPRO='" & RsMovAlmacen.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'            If Not RsProducto.EOF Then
'                wstock = 0
'                SSQL = "INSERT INTO TMPINVENTARIO(F3ITEM,F5CODPRO,F5CODFAB,F5NOMPRO,F3STOCKSIS,F7CODMED,F3STOCKFIS,F3DIFERENCIA,F3COSTOSIS,F3COSTOACT) VALUES(" & var & ",'" & RsMovAlmacen.Fields("F5CODPRO") & "', " & _
'                       "'" & IIf(IsNull(RsProducto.Fields("F5CODFAB")), "", RsProducto.Fields("F5CODFAB")) & "','" & Left(RsProducto.Fields("F5NOMPRO"), 50) & "'," & Format(RsMovAlmacen.Fields("F6STOCKACT"), "0.00000") & ", " & _
'                       "'" & Trim(RsProducto.Fields("F7CODMED")) & "'," & wstock & "," & RsMovAlmacen.Fields("F6STOCKACT") & "," & RsMovAlmacen.Fields("F5COSPRO") & ",'0.00')"
'                Temp.Execute (SSQL)
'                var = var + 1
'            End If
'            RsProducto.Close
'            RsMovAlmacen.MoveNext
'            If RsMovAlmacen.EOF Then Exit Do
'        Loop
'    End With
'    dxDBGrid1.Dataset.ADODataset.ConnectionString = Temp
'    dxDBGrid1.Dataset.Active = True
'    dxDBGrid1.Dataset.Open
'    dxDBGrid1.Dataset.First
'    dxDBGrid1.Columns.FocusedIndex = 0
'
'End If
'RsMovAlmacen.Close

cnombase = "TEMPFAC.mdb"

If cnn_form.State = adStateOpen Then cnn_form.Close
cnn_form.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "" & cnombase & ";Persist Security Info=False"
cnn_dbbancos.Execute "DELETE FROM TMPINVENTARIO"

    sql = "Insert Into TMPINVENTARIO SELECT DISTINCT Iif(isnull(Consulta3.CANTIDAD),0,Consulta3.CANTIDAD) as " & _
    " F3STOCKSIS, consulta3.StockFis as F3STOCKFIS, (0 - Consulta3.CANTIDAD) As F3DIFERENCIA, IF5PLA.F5CODPRO, IF5PLA.F5CODFAB AS CODFAB,  IF5PLA.F5NOMPRO, IF5PLA.F7CODMED "
    sql = sql + " FROM ([SELECT IF3VALES.F5CODPRO, IF3VALES.F2CODALM, Sum(IIf(Left(IF3VALES.F4NUMVAL,1)='I', " & _
    "IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, 0 as StockFis FROM IF3VALES in '" & wrutabancos & "\db_bancos.mdb' GROUP BY IF3VALES.F5CODPRO, " & _
    " IF3VALES.F2CODALM HAVING ((IF3VALES.F2CODALM) = '" & wcod_alm & "')]. AS Consulta3"
    
    sql = sql + " RIGHT JOIN IF5PLA ON Consulta3.F5CODPRO = IF5PLA.F5CODPRO) "
    
    sql = sql + " LEFT JOIN  [SELECT IF3VALES.F2CODALM, IF3VALES.F5CODPRO FROM IF3VALES in '" & wrutabancos & "\db_bancos.mdb' GROUP BY IF3VALES.F2CODALM, " & _
    " IF3VALES.F5CODPRO]. AS Consulta2 ON IF5PLA.F5CODPRO = Consulta2.F5CODPRO"
    
    sql = sql + " in '" & wrutabancos & "\db_bancos.mdb' " & IIf(ChkVSF.Value = 1, " WHERE consulta3.StockFis > 0", "") & " GROUP BY Consulta3.CANTIDAD,IF5PLA.F5CODPRO, IF5PLA.F5CODFAB,  IF5PLA.F5NOMPRO, IF5PLA.F7CODMED, " & _
    " IF5PLA.F5VALVTA, consulta3.StockFis"
    
    sql = sql & " ORDER BY IF5PLA.F5CODPRO;"
'MsgBox sql
'sql = "Insert Into TMPINVENTARIO SELECT DISTINCTROW TBPRODUCTOS.CODPRO As F5CODPRO, TBPRODUCTOS.NOMPRO As F5NOMPRO, TBPRODUCTOS.F5FACTOR, TBPRODUCTOS.DESMAR, " & _
'" TBPRODUCTOS.F7CODMED, TBPRODUCTOS.CODFAB, TBPRODUCTOS.F5PREVTA, TBPRODUCTOS.NIVEL01, IIF(ISNULL " & _
'" (TBEXISTENCIAS.CANTIDAD),0,TBEXISTENCIAS.CANTIDAD) AS F3STOCKSIS, 0 AS F3STOCKFIS, 0 - IIF(ISNULL " & _
'" (TBEXISTENCIAS.CANTIDAD),0,TBEXISTENCIAS.CANTIDAD) As F3DIFERENCIA, TBPRODUCTOS.COSTO_UNITARIO FROM TBPRODUCTOS " & _
'" LEFT JOIN TBEXISTENCIAS ON TBPRODUCTOS.CODPRO = TBEXISTENCIAS.F5CODPRO WHERE  Val(IIf(IsNull( " & _
'" [TBEXISTENCIAS].[CANTIDAD]),0,[TBEXISTENCIAS].[CANTIDAD]))> -10000000 AND TBEXISTENCIAS.F2CODALM = '" & Txtcodalm.Text & "' ORDER BY TBPRODUCTOS.CODPRO"


cnn_form.Execute sql
'dxDBGrid1.DefaultFields = True
dxDBGrid1.Dataset.Active = False
dxDBGrid1.DefaultFields = True
dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
dxDBGrid1.Dataset.ADODataset.CommandType = cmdText

dxDBGrid1.Dataset.ADODataset.CommandText = "Select * From TMPINVENTARIO"

dxDBGrid1.Dataset.Active = True
'dxDBGrid1.Dataset.ADODataset.Requery
dxDBGrid1.KeyField = "F5CODPRO"

dxDBGrid1.Columns(0).Width = 150
dxDBGrid1.Columns(1).Width = 150
dxDBGrid1.Columns(2).Visible = False
dxDBGrid1.Columns(3).Visible = False
dxDBGrid1.Columns(4).Width = 60
dxDBGrid1.Columns(5).Visible = False
dxDBGrid1.Columns(6).Visible = False
dxDBGrid1.Columns(7).Visible = False
dxDBGrid1.Columns(11).Visible = False
dxDBGrid1.Columns(0).Caption = "Código"
dxDBGrid1.Columns(1).Caption = "Nombre del Producto"
dxDBGrid1.Columns(2).Caption = "Medida"
dxDBGrid1.Columns(8).Caption = "Stk. Sis."
dxDBGrid1.Columns(9).Caption = "Stk. Fis"
dxDBGrid1.Columns(10).Caption = "Dif."

If ChkVSF.Value = 1 Then
    dxDBGrid1.Columns(8).Visible = False
    dxDBGrid1.Columns(10).Visible = False
Else
    dxDBGrid1.Columns(8).Visible = True
    dxDBGrid1.Columns(9).Visible = True
    dxDBGrid1.Columns(10).Visible = True
End If

    

End Sub

Private Sub calcula()
With dxDBGrid1
    If IsNull(.Columns(4).Value) Then .Columns(4).Value = Format(0, "###,##0.00000")
    If IsNull(.Columns(5).Value) Then .Columns(5).Value = Format(0, "###,##0.00000")
    .Columns.ColumnByFieldName("F3DIFERENCIA").Value = Format(Val(.Columns.ColumnByFieldName("F3STOCKSIS").Value) - Val(.Columns.ColumnByFieldName("F3STOCKFIS").Value), "0.00000")
End With
End Sub

Private Sub grabar()
On Error GoTo HNDERR
If RsTomaCab.State = adStateOpen Then RsTomaCab.Close
RsTomaCab.Open "Select F2CODUSER from H4tomainv where CVDATE(F4fectom) = '" & CVDate(wultinv) & "' and F2codalm = '" & wcod_alm & "'", cnn_dbbancos
If RsTomaCab.EOF Then
    '--------nuevo
    sql = "INSERT INTO H4TOMAINV(F4CIERRE,F4FECTOM,F2CODUSER,F2CODALM) VALUES('0','" & CVDate(abofecha.Value) & "','" & wusuario & "','" & wcod_alm & "')"
Else
    '---------editar
    sql = "UPDATE H4TOMAINV SET F2CODUSER =  '" & wusuario & "' WHERE CVDATE(F4FECTOM) = '" & CVDate(wultinv) & "' AND F2CODALM = '" & wcod_alm & "'"
End If
cnn_dbbancos.Execute (sql)
RsTomaCab.Close
Graba_Producto
Exit Sub
HNDERR:
MsgBox "Ha Ocurrido el Siguiente Error : " & Err.Description, vbExclamation, "Sistema de Logística"
Resume
Exit Sub
End Sub

 Sub Graba_Producto()
    Dim nitems As Integer
    Dim nfil       As Integer
    Dim cvalores As String
    Dim RsInventario As New ADODB.Recordset
    
    amovs_det(0).campo = "F4FECTOM": amovs_det(0).valor = "": amovs_det(0).TIPO = "T"
    amovs_det(1).campo = "F2CODALM": amovs_det(1).valor = "": amovs_det(1).TIPO = "T"
    amovs_det(2).campo = "F5CODPRO": amovs_det(2).valor = "": amovs_det(2).TIPO = "T"
    amovs_det(3).campo = "F7CODMED": amovs_det(3).valor = "": amovs_det(3).TIPO = "T"
    amovs_det(4).campo = "F3STOCKSIS": amovs_det(4).valor = "": amovs_det(4).TIPO = "N"
    amovs_det(5).campo = "F3STOCKFIS": amovs_det(5).valor = "": amovs_det(5).TIPO = "N"
    amovs_det(6).campo = "F3DIFERENCIA": amovs_det(6).valor = "": amovs_det(6).TIPO = "N"
'    amovs_det(7).campo = "F3COSTOSIS": amovs_det(7).valor = "": amovs_det(7).TIPO = "N"
'    amovs_det(8).campo = "F3COSTOACT": amovs_det(8).valor = "": amovs_det(8).TIPO = "N"

    nitems = 0

    If RsInventario.State = adStateOpen Then RsInventario.Close
    Call BASE_TEMPORAL("tempfac.MDB")
    RsInventario.Open "Select Count(*) as Total from TmpInventario", Temp, adOpenDynamic
    If Not RsInventario.EOF Then
        nitems = RsInventario.Fields("Total")
    End If
    RsInventario.Close
    ReDim Values(6, nitems)
    
    If RsInventario.State = adStateOpen Then RsInventario.Close
    RsInventario.Open "Select * from TmpInventario", Temp, adOpenDynamic
    If Not RsInventario.EOF Then
        nfil = 0
        RsInventario.MoveFirst
        Do While Not RsInventario.EOF
            Values(0, nfil) = Format(abofecha.Value, "DD/MM/YYYY")
            Values(1, nfil) = wcod_alm
            Values(2, nfil) = "" & RsInventario.Fields("F5codpro")
            Values(3, nfil) = "" & RsInventario.Fields("F7codmed")
            Values(4, nfil) = 0 + Format(RsInventario.Fields("F3StockSis"), "0.00")
            Values(5, nfil) = 0 + Format(RsInventario.Fields("F3StockFis"), "0.00")
            Values(6, nfil) = 0 + Format(IIf(IsNull(RsInventario.Fields("F3Diferencia")), 0, RsInventario.Fields("F3Diferencia")), "0.00")
'            Values(7, nfil) = 0 + Format(RsInventario.Fields("F3CostoSis"), "0.00")
'            Values(8, nfil) = 0 + Format(RsInventario.Fields("F3CostoAct"), "0.00")
            RsInventario.MoveNext
            nfil = nfil + 1
        Loop
    End If
    RsInventario.Close
            
    cvalores = "111111111"
    ctipo = "A"
    cnn_dbbancos.Execute ("DelEte * from H3TomaInv Where cvdate(F4FecTom) = '" & CVDate(wultinv) & "' and F2Codalm = '" & wcod_alm & "'")
    GRABA_REGISTRO_DET amovs_det(), "H3TOMAINV", "A", 6, cnn_dbbancos, "", Values(), nfil - 1, cvalores, "", ""
    
End Sub

Private Sub CERRAR()
On Error GoTo HNDERR
    Dim rstemporal      As New ADODB.Recordset
    Dim RsInventario   As New ADODB.Recordset
    Dim wnumval        As String
    Dim cvalores        As String
    Dim csql             As String
    Dim nitems            As Integer
    Dim nfil                 As Integer
    Dim I                     As Integer
    Dim Dif                 As Double
    Dim ctipo              As String
    Dim cnumvale      As String
    Dim ccampo         As String
    Dim mes                 As String
    
    Dif = 0#
    If RsTomaCab.State = adStateOpen Then RsTomaCab.Close
    RsTomaCab.Open "Select F4NumVali,F4NumVals,F2CodUser,F4Cierre From H4TOMAINV Where cvdate(F4Fectom) = '" & CVDate(abofecha.Value) & "' and F2CodAlm = '" & wcod_alm & "'", cnn_dbbancos, adOpenDynamic
    If Not RsTomaCab.EOF Then
        For I = 1 To 2
            If rstemporal.State = adStateOpen Then rstemporal.Close
            If I = 1 Then
                sql = "Select * from TmpInventario Where F3Diferencia < " & Dif & ""
            Else
                sql = "Select * from TmpInventario Where F3Diferencia > " & Dif & ""
            End If
            rstemporal.Open sql, Temp
            If Not rstemporal.EOF Then
                If I = 1 Then
                    Rem MARLY cnumvale = GENERA_NUMVALE(wcod_alm, Format(Month(abofecha.Value), "00"), "I")
                    wnumval = cnumvale
                Else
                    Rem MARLY cnumvale = GENERA_NUMVALE(wcod_alm, Format(Month(abofecha.Value), "00"), "S")
                End If
                              
               '--------------------------------ASIGNA DATOS A LA CABECERA ------------------------------------
                avales_cab(0).campo = "F4NUMVAL": avales_cab(0).valor = cnumvale: avales_cab(0).TIPO = "T"
                avales_cab(1).campo = "F2CODPROV": avales_cab(1).valor = "": avales_cab(1).TIPO = "T"
                avales_cab(2).campo = "F4NUMDOC": avales_cab(2).valor = cnumvale: avales_cab(2).TIPO = "T"
                avales_cab(3).campo = "F1CODDOC": avales_cab(3).valor = "35": avales_cab(3).TIPO = "T"
                avales_cab(4).campo = "F4FECVAL": avales_cab(4).valor = abofecha.Value: avales_cab(4).TIPO = "F"
                avales_cab(5).campo = "F2CODALM": avales_cab(5).valor = wcod_alm: avales_cab(5).TIPO = "T"
                avales_cab(6).campo = "F1CODORI": avales_cab(6).valor = "XJ0": avales_cab(6).TIPO = "T"
                avales_cab(7).campo = "F4TIPCAM": avales_cab(7).valor = wtipcam: avales_cab(7).TIPO = "N"
                avales_cab(8).campo = "F4MONEDA": avales_cab(8).valor = wmoneda_productos: avales_cab(8).TIPO = "T"
                avales_cab(9).campo = "F4CENTRO": avales_cab(9).valor = vcodigocentro: avales_cab(9).TIPO = "T"
                avales_cab(10).campo = "F4FECGRA": avales_cab(10).valor = abofecha.Value: avales_cab(10).TIPO = "F"
                avales_cab(11).campo = "F4USEGRA": avales_cab(11).valor = wusuario: avales_cab(11).TIPO = "T"
                
                '-----------------------------ASIGNA DATOS AL DETALLE ------------------------------------------------
                
                avales_det(0).campo = "F4NUMVAL": avales_det(0).valor = "": avales_det(0).TIPO = "T"
                avales_det(1).campo = "F2CODALM": avales_det(1).valor = "": avales_det(1).TIPO = "T"
                avales_det(2).campo = "F5CODPRO": avales_det(2).valor = "": avales_det(2).TIPO = "T"
                avales_det(3).campo = "F4FECVAL": avales_det(3).valor = "": avales_det(3).TIPO = "F"
                avales_det(4).campo = "F3CANPRO": avales_det(4).valor = "": avales_det(4).TIPO = "N"
                avales_det(5).campo = "F3VALVTA": avales_det(5).valor = "": avales_det(5).TIPO = "N"
                avales_det(6).campo = "F3VALDOL": avales_det(6).valor = "": avales_det(6).TIPO = "N"
                avales_det(7).campo = "F3TOTITE": avales_det(7).valor = "": avales_det(7).TIPO = "N"
                avales_det(8).campo = "F3TOTDOL": avales_det(8).valor = "": avales_det(8).TIPO = "N"
                avales_det(9).campo = "F3GRUPO": avales_det(9).valor = "": avales_det(9).TIPO = "T"

                If I = 1 Then
                    csql = "Select Count(F3Item) as NTOTAL From TmpInventario Where F3Diferencia < " & Dif & ""
                Else
                    csql = "Select Count(F3Item) as NTOTAL From TmpInventario Where F3Diferencia > " & Dif & ""
                End If
                If RsInventario.State = adStateOpen Then RsInventario.Close
                RsInventario.Open csql, Temp
                If Not RsInventario.EOF Then
                    nitems = RsInventario.Fields("NTOTAL")
                End If
                RsInventario.Close
                ReDim Values(9, nitems)
                If wtipcam = 0 Then wtipcam = 3.6
                If RsInventario.State = adStateOpen Then RsInventario.Close
                RsInventario.Open sql, Temp
                If Not RsInventario.EOF Then
                    nfil = 0
                    Do While Not RsInventario.EOF
                        Values(0, nfil) = cnumvale
                        Values(1, nfil) = wcod_alm
                        Values(2, nfil) = "" & RsInventario.Fields("F5CODPRO")
                        Values(3, nfil) = abofecha.Value
                        Values(4, nfil) = 0 + Format(RsInventario.Fields("F3STOCKFIS"), "0.00")
                        Values(5, nfil) = 0 + Format(RsInventario.Fields("F3COSTOACT"), "0.00")
                        Values(6, nfil) = 0 + Format(Val(RsInventario.Fields("F3COSTOACT") / wtipcam), "0.00")
                        Values(7, nfil) = 0 + Format(Val(RsInventario.Fields("F3STOCKFIS") * RsInventario.Fields("F3COSTOACT")), "0.00")
                        Values(8, nfil) = 0 + Format(Val(RsInventario.Fields("F3STOCKFIS") * (RsInventario.Fields("F3COSTOACT") / wtipcam)), "0.00")
                        Values(9, nfil) = "" & Left(RsInventario.Fields("F5CODPRO"), 1)
                        RsInventario.MoveNext
                        nfil = nfil + 1
                    Loop
                End If
                RsInventario.Close
            End If
            rstemporal.Close
            cvalores = "1111111111"
            mes = Format(Month(abofecha.Value), "00")
            ctipo = "A"
                If ctipo = "A" Then     '--- Nuevo
                    '------- GRABA CABECERA
                    GRABA_REGISTRO avales_cab(), "IF4VALES", ctipo, 11, cnn_dbbancos, ""
                    
                    If sw_graba_registro = True Then
                        '------- GRABA DETALLE
                        sw_inventario = True
                        GRABA_REGISTRO_DET avales_det(), "IF3VALES", ctipo, 9, cnn_dbbancos, "", Values(), nfil - 1, cvalores, mes, "A"
                    End If
                    If I = 1 Then
                        If MsgBox("Desea Imprimir el Vale de Ingreso", vbYesNo + vbInformation, "AVISO") = vbYes Then
                            Rem MARLY Imprime_Vale wcod_alm, wnumval, 0
                        End If
                    Else
                        If MsgBox("¿Desea Imprimir el Vale de Salida?", vbYesNo + vbInformation, "AVISO") = vbYes Then
                            Rem MARLY Imprime_Vale wcod_alm, cnumvale, 0
                        End If
                    End If
                    If I = 1 Then
                        ccampo = "F1VALING" & mes
                    Else
                        ccampo = "F1VALSAL" & mes
                    End If
                    ACTUALIZA_ALMA_VALE cnumvale, ccampo, wcod_alm
                    
                Else
                    '------- GRABA CABECERA
                    GRABA_REGISTRO avales_cab(), "IF4VALES", ctipo, 11, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM='" & wcod_alm & "'"
                    '-------------------------------------------------------
                    '------- RESTA LOS SALDOS
                    If rsif3vales.State = adStateOpen Then rsif3vales.Close
                    rsif3vales.Open "SELECT * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & wcod_alm & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsif3vales.EOF Then
                        Do While Not rsif3vales.EOF
                            GRABA_SALDO_ALM rsif3vales.Fields("F5CODPRO") & "", rsif3vales.Fields("F3CANPRO"), rsif3vales.Fields("F3TOTITE"), mes, "I", cnn_dbbancos, wcod_alm, rsif3vales.Fields("F3TOTDOL"), "R"
                            rsif3vales.MoveNext
                        Loop
                    End If
                    rsif3vales.Close
                    '------- GRABA DETALLE
                    cnn_dbbancos.Execute ("DELETE * FROM IF3VALES WHERE F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & wcod_alm & "'")
                    GRABA_REGISTRO_DET amovs_det(), "IF3VALES", "A", 9, cnn_dbbancos, "F4NUMVAL = '" & cnumvale & "' AND F2CODALM = '" & wcod_alm & "'", Values(), nfil - 1, cvalores, mes, "A"
                End If
        Next I
        cnn_dbbancos.Execute ("Update H4TomaInv Set F4Numvali = '" & wnumval & "',F4Numvals = '" & cnumvale & "',F2CodUser = '" & wusuario & "',F4Cierre = '1' Where cvdate(F4Fectom) = '" & CVDate(abofecha.Value) & "' and F2CodAlm = '" & wcod_alm & "'")
        cnn_dbbancos.Execute ("Update EF2ALMACENES Set F1ULTINV = '" & CVDate(abofecha.Value) & "' WHERE F2CODALM = '" & wcod_alm & "'")
        wultinv = abofecha.Value
        sw_inventario = False
    End If
    RsTomaCab.Close

Exit Sub
HNDERR:
MsgBox "Ha Ocurrido el Siguiente Error: " & Err.Description, vbExclamation, "Sistema de Logística"
Exit Sub
End Sub

Private Sub ACTUALIZA_ALMA_VALE(pnumvale As String, pcampo As String, palmacen As String)
Dim csql    As String
        
    csql = "UPDATE EF2ALMACENES SET " & pcampo & " =  '" & pnumvale & "' WHERE '" & pnumvale & "' > " & pcampo & " AND F2CODALM='" & palmacen & "'"
    cnn_dbbancos.Execute csql
    
End Sub

''''''''Sub Imprimir_Saldo()
''''''''    Dim RsInventario As New ADODB.Recordset
''''''''
''''''''    Printer.ScaleMode = 4
''''''''    Printer.FontName = "Courier New"
''''''''    Imprime_Titulo
''''''''    wfila = 13
''''''''
''''''''    If RsInventario.State = adStateOpen Then RsInventario.Close
''''''''    RsInventario.Open "Select * from TmpInventario", Temp, adOpenDynamic
''''''''    If Not RsInventario.EOF Then
''''''''        RsInventario.MoveFirst
''''''''        Do While Not RsInventario.EOF
''''''''            writexy RsInventario.Fields("F5codpro"), wfila, 1, 0
''''''''            writexy Left(RsInventario.Fields("F5NOMPRO"), 25), wfila, 12, 0
''''''''            writexy Format(RsInventario.Fields("F3StockSis"), "0.00000"), wfila, 40, 2
''''''''            writexy RsInventario.Fields("F7codmed"), wfila, 52, 0
''''''''            If wcierre = "0" Then
''''''''                If sw_diferencia = True Then
''''''''                    writexy Format(RsInventario.Fields("F3StockFis"), "0.00000"), wfila, 68, 2
''''''''                    writexy Format(RsInventario.Fields("F3Diferencia"), "0.00"), wfila, 70, 2
''''''''                    writexy Format(RsInventario.Fields("F3CostoAct"), "0.00"), wfila, 82, 2
''''''''                Else
''''''''                    writexy "....................", wfila, 68, 0
''''''''                End If
''''''''            Else
''''''''                writexy Format(RsInventario.Fields("F3StockFis"), "0.00000"), wfila, 58, 2
''''''''                writexy Format(RsInventario.Fields("F3Diferencia"), "0.00"), wfila, 70, 2
''''''''                writexy Format(RsInventario.Fields("F3CostoAct"), "0.00"), wfila, 82, 2
''''''''            End If
''''''''            wfila = wfila + 1
''''''''            If wfila >= 60 Then
''''''''                Printer.NewPage
''''''''                Imprime_Titulo
''''''''                wfila = 13
''''''''            End If
''''''''            RsInventario.MoveNext
''''''''        Loop
''''''''    End If
''''''''    RsInventario.Close
''''''''    Printer.Line (1, wfila)-(95, wfila)
''''''''    Printer.EndDoc
''''''''
''''''''End Sub
''''''''
Sub Imprime_Titulo()

    Rem MARLY CABECERA
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.FontUnderline = True
    If RsTomaCab.State = adStateOpen Then RsTomaCab.Close
    RsTomaCab.Open "SELECT F4CIERRE FROM H4TOMAINV WHERE CVDATE(F4FECTOM) = '" & CVDate(wultinv) & "' AND F2CODALM = '" & wcod_alm & "'", cnn_dbbancos
    If Not RsTomaCab.EOF Then
        wcierre = RsTomaCab.Fields("F4CIERRE")
        If RsTomaCab.Fields("F4CIERRE") = "0" Then
            writexy "TOMA DE INVENTARIO FISICO", 5, 32, 0
        Else
            writexy "SALDO DE DIFERENCIAS DEL INVENTARIO", 5, 32, 0
            writexy "Diferencia", 11, 80, 0
        End If
        Printer.FontUnderline = False
        Printer.FontSize = 10
        writexy "Almacén:", 8, 1, 0
        writexy Trim(wcod_alm & " - " & wnomalmacen), 8, 12, 0
        Printer.FontSize = 9
        Printer.Line (1, 10)-(95, 10)
        writexy "Código", 11, 1, 0
        writexy "Descripción", 11, 12, 0
        writexy "Stock-Sist.", 11, 50, 0
        writexy "Und.", 11, 62, 0
        writexy "Stock-Real", 11, 68, 0
        If sw_diferencia = True Then
             writexy "Diferencia", 11, 80, 0
        End If
        
        Printer.Line (1, 12)-(95, 12)
        Printer.FontBold = False
    End If
    RsTomaCab.Close
End Sub


