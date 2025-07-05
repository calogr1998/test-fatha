VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Begin VB.Form importar_solicitud 
   Caption         =   "Importar Requerimientos"
   ClientHeight    =   5250
   ClientLeft      =   405
   ClientTop       =   1605
   ClientWidth     =   10560
   LinkTopic       =   "Form2"
   ScaleHeight     =   5250
   ScaleWidth      =   10560
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1710
      TabIndex        =   2
      Top             =   135
      Width           =   2055
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   270
      TabIndex        =   1
      Top             =   135
      Width           =   1455
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   1770
      Left            =   120
      OleObjectBlob   =   "importar_solicitud.frx":0000
      TabIndex        =   0
      Top             =   3360
      Width           =   10380
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2610
      Left            =   120
      OleObjectBlob   =   "importar_solicitud.frx":2DCD
      TabIndex        =   3
      Top             =   480
      Width           =   10260
   End
End
Attribute VB_Name = "importar_solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wnorden, StrSql As String
Dim wnumi As Integer
Dim csql     As String
Dim rsOrden As ADODB.Recordset

Private Sub Checkagrupar_Click()
    If Checkagrupar.Value = 1 Then
      dxDBGrid1.Options.Set (egoShowGroupPanel)
    Else
      dxDBGrid1.Options.Unset (egoShowGroupPanel)
    End If
   
End Sub

Private Sub CheckFiltro_Click()
    If CheckFiltro.Value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub dxDBGrid1_OnBackgroundDraw(ByVal hdc As Long, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single, Done As Boolean)
    Dim X As Integer, y As Integer
    Dim IsClipRgnExists As Boolean
    Dim PrevClipRgn As Long, Rgn As Long
    Dim s, OldFont As Long
    Dim Font1 As IFont
    
    'If dxDBGrid1.Ex.GroupColumnCount < 1 Then
    ' s = "Arrastre una columna aquí para agrupar información"
    ' SetBkMode hdc, TRANSPARENT
    ' Set Font1 = dxDBGrid1.Columns.HeaderFont
    ' OldFont = SelectObject(hdc, Font1.hFont)
    ' DrawText hdc, s, Len(s), R, DT_SINGLELINE + DT_VCENTER
    ' Call SelectObject(hdc, OldFont)
    'End If
End Sub

Private Sub dxDBGrid1_OnClick()
proceso2 dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value
End Sub

'Private Sub dxDBGrid1_OnChangeNodeEx()
'    Dim valor As String
'    valor = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value
'    valor = dxDBGrid1.Columns(0).Value
'    proceso2 (valor)
'End Sub

Private Sub dxDBGrid1_OnDblClick()
wnumi = 0
    wnorden = dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value
    vale_salida.TxtNumsol.Text = wnorden
    Set rslista = New ADODB.Recordset
    StrSql = "select * from tb_detsolicitud where cod_solicitud='" & wnorden & "'"
    rslista.Open StrSql, cnn_dbbancos, adOpenDynamic, adLockReadOnly
    With vale_salida.dxDBGrid1
        Do While Not rslista.EOF
                wnumi = wnumi + 1
                If wnumi = 1 Then
                  .Dataset.Edit
                Else
                  .Dataset.Insert
                End If
                .Columns.ColumnByFieldName("ITEM").Value = wnumi
                .Columns.ColumnByFieldName("CODPROD").Value = rslista!COD_PRODUCTO
                .Columns.ColumnByFieldName("CODFAB").Value = rslista!F5CODFAB
                .Columns.ColumnByFieldName("DESCRIPCION").Value = rslista!ds_descripcion
                .Columns.ColumnByFieldName("UMEDIDA").Value = rslista!ds_unidmed
                .Columns.ColumnByFieldName("MARCA").Value = rslista!F5MARCA
                .Columns.ColumnByFieldName("CANTIDAD").Value = rslista!ds_cantidad
                .Columns.ColumnByFieldName("STOCKACTUAL").Value = rslista!stok
                .Columns.ColumnByFieldName("PUNIT").Value = rslista!Precio
                .Columns.ColumnByFieldName("TOTAL").Value = rslista!f5valvta
                .Dataset.Post
                .Columns.FocusedIndex = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").ColIndex
       
          rslista.MoveNext
        Loop
    End With
    rslista.Close
'    dxDBGrid1.Dataset.Active = False
'    dxDBGrid2.Dataset.Active = False
'    dxDBGrid1.Dataset.Close
'    dxDBGrid2.Dataset.Close
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.AutoRedraw = False
    Me.Left = 1450
    Me.Top = 1050
    
    'sw_nuevo_documento = True
    Me.AutoRedraw = True
    
    With dxDBGrid1
        .Dataset.Close
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
         csql = "SELECT COD_SOLICITUD,CS_FECHA,CS_CODAREA,CS_ESTADO,CS_TOTAL,CS_MONEDA,NUMORDEN FROM TB_CABSOLICITUD WHERE CS_CODAREA='002' AND num_val is null"
       ' .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        '.Dataset.Active = True
        
        
    End With
    
    proceso
    
    proceso2 dxDBGrid1.Columns.ColumnByFieldName("COD_SOLICITUD").Value


'    With dxDBGrid1
'        .Options.Unset (egoShowGroupPanel)
'        .Filter.FilterActive = False
'    End With
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Nuevo"
            sw_nuevo_documento = True
            Me.MousePointer = 11
            If wtipoguia = "I" Then
                vale_ingreso.Show 1
            Else
                vale_salida.Show 1
            End If
            Me.MousePointer = 1
        Case "ID_Salir"
            Unload Me
    End Select
End Sub

Public Sub proceso()
    
    With dxDBGrid1
        '.Dataset.Close
        'csql = "SELECT COD_SOLICITUD,CS_FECHA,CS_CODAREA,CS_ESTADO,CS_TOTAL,CS_MONEDA,NUMORDEN FROM TB_CABSOLICITUD WHERE num_val is null"
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
         csql = "SELECT COD_SOLICITUD,CS_FECHA,CS_CODAREA,CS_ESTADO,CS_TOTAL,CS_MONEDA,NUMORDEN FROM TB_CABSOLICITUD WHERE CS_CODAREA='002' AND num_val is null"
         
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = csql
        
        .Dataset.Active = True
        '.Dataset.Open
        .KeyField = "COD_SOLICITUD"
        
        .Dataset.ADODataset.Requery
    End With
 
End Sub

Public Sub proceso2(ByVal codigo As String)
    Dim sql As String
    'codigo = "0004"
    With dxDBGrid2
        .Dataset.Close
        .Dataset.ADODataset.ConnectionString = cnn_dbbancos
        'SQL = "SELECT a.f3codpro,a.f3codfab,a.f5nompro,b.f7codmed,a.f3canpro, a.f3canfal FROM if3orden a left join if5pla b on a.f3codpro = b.f5codpro WHERE a.f4numord= " & codigo & " order by a.f3codpro"
        sql = "SELECT COD_SOLICITUD,DS_CANTIDAD,COD_PRODUCTO,DS_UNIDMED,DS_DESCRIPCION,PRECIO,STOK FROM TB_DETSOLICITUD WHERE COD_SOLICITUD='" & codigo & "'"
        
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = sql
        .Dataset.Active = True
        .Dataset.Open
        .KeyField = "COD_PRODUCTO"
        
    End With
End Sub

Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    With dxDBGrid1.Dataset
        If dxDBGrid1.Columns.FocusedColumn.ColumnType = gedLookupEdit Then
            If .State = dsEdit Then
                dxDBGrid1.m.HideEditor
                .Post
                .DisableControls
                .Close
                .Open
                .EnableControls
            End If
        End If
End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
'dxDBGrid1.Dataset.Close
End Sub

