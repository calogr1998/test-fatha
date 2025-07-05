VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Begin VB.Form actualiza_producto_prov 
   Caption         =   "Actualización de Productos / Proveedor"
   ClientHeight    =   7335
   ClientLeft      =   1320
   ClientTop       =   1245
   ClientWidth     =   10245
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
   ScaleHeight     =   7335
   ScaleWidth      =   10245
   Begin VB.Frame Frame1 
      Height          =   6765
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   10005
      Begin VB.TextBox txtproducto 
         Height          =   330
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   0
         Top             =   405
         Width           =   1680
      End
      Begin Threed.SSPanel pnlproducto 
         Height          =   330
         Left            =   2925
         TabIndex        =   3
         Top             =   405
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   582
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
         Height          =   5430
         Left            =   225
         OleObjectBlob   =   "actualiza_producto_prov.frx":0000
         TabIndex        =   4
         Top             =   1170
         Width           =   9555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   210
         Left            =   315
         TabIndex        =   2
         Top             =   450
         Width           =   645
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   225
      Top             =   6885
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "actualiza_producto_prov.frx":2C1D
      ToolBars        =   "actualiza_producto_prov.frx":77F9
   End
End
Attribute VB_Name = "actualiza_producto_prov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cconex_form         As String
Dim cnn_form            As New ADODB.Connection
Dim CadSql              As String
Dim sw_nuevo_item       As Boolean
Dim rstemporal          As New ADODB.Recordset
Dim sw_ayuda_prod       As Boolean
Dim rsproductos         As New ADODB.Recordset

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
        '.Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        '.Set (egoShowFooter)
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
    dxDBGrid1.Columns(4).Visible = False
    
End Sub

Private Sub dxDBGrid1_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)

    If sw_nuevo_item = False Then
        If Action = daInsert Then
            dxDBGrid1.Dataset.Edit
            dxDBGrid1.Columns.ColumnByFieldName("ITEM").Value = dxDBGrid1.Dataset.RecordCount + 1
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
    Select Case dxDBGrid1.Columns.FocusedIndex
        Case 0:
            wrucprov = ""
            'hlp_proveedores.Show 1
            ayuda_proveedores.Show 1
            If Len(Trim(wrucprov)) > 0 Then
                If rstemporal.State = adStateOpen Then rstemporal.Close
                rstemporal.Open "SELECT RUCPROV FROM DETALLE_PROV WHERE RUCPROV='" & wrucprov & "'", cnn_form, adOpenDynamic, adLockOptimistic
                If rstemporal.EOF Then
                    dxDBGrid1.Dataset.Edit
                    dxDBGrid1.Columns.ColumnByFieldName("RUCPROV").Value = wrucprov
                    dxDBGrid1.Columns.ColumnByFieldName("NOMPROV").Value = wnomprov
                Else
                    MsgBox "El proveedor ya fue asignado al producto. Verifique.", vbInformation, "Atención"
                End If
                rstemporal.Close
            End If
        Case 3:
            wcodpag = ""
            hlp_formapago.Show 1
'            ayu_f_p_c.Show 1
            If Len(Trim(wcodpag)) > 0 Then
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("FPAGO").Value = wcodpag
            End If
    End Select
End Sub
Private Sub dxDBGrid1_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    If KeyCode = 115 Then
        If MsgBox("Desea eliminar el registro actual ", vbQuestion + vbYesNo, "Atención") = vbYes Then
            sw_nuevo_item = True
            dxDBGrid1.Dataset.Delete
            sw_nuevo_item = False
        End If
    End If
    
End Sub

Private Sub Form_Load()

    sw_ayuda_prod = False
    sw_nuevo_item = False
    
    'cnombase = wusuario & "PROVPROD" & Format(Time, "hh_mm_ss") & ".MDB"
    'CREATEDATABASE_N wrutatemp & "\", cnombase
    cnombase = "TMP_PROP.MDB"
    cnomtabla = "DETALLE_PROV"
    
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    
    'CadSql = "(ITEM TEXT(6),CODPROD TEXT(20),NOMPROD TEXT(100),UMEDIDA TEXT(3),PRECIO DOUBLE,CODFAB TEXT(20))"
    'CREATETABLE_N cnomtabla, CadSql, cnn_form
    
    CONFIGURA_GRID
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    sw_nuevo_item = True
    dxDBGrid1.Dataset.Close
    cnn_form.Close
    
    If sw_ayuda_prod = True Then
        Unload hlp_productos
    End If
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Grabar":
            Me.MousePointer = 11
            If dxDBGrid1.Dataset.State = dsEdit Or dxDBGrid1.Dataset.State = dsInsert Then
                dxDBGrid1.Dataset.Post
            End If
            GRABA_PROD
            Me.MousePointer = 1
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub

Private Sub LLENA_PRODUCTOS(pproducto As String)
Dim rsprov_prod     As New ADODB.Recordset
Dim nitem           As Integer
Dim csql            As String
    
    DELETEREC_N cnomtabla, cnn_form
    DELETEREC_N cnomtabla, cnn_form
    
    nitem = 0
    If rsprov_prod.State = adStateOpen Then rsprov_prod.Close
    rsprov_prod.Open "SELECT * FROM EF2PROD_PROV WHERE F5CODPRO='" & pproducto & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsprov_prod.EOF Then
        rsprov_prod.MoveFirst
        
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        dxDBGrid1.OptionEnabled = False
        dxDBGrid1.Dataset.DisableControls
        
        With dxDBGrid1.Dataset
            Do While Not rsprov_prod.EOF
                nitem = nitem + 1
                sw_nuevo_item = True
                .Append
                .FieldValues("ITEM") = 1
                .FieldValues("RUCPROV") = "" & rsprov_prod.Fields("F2CODPRV")
                .FieldValues("NOMPROV") = "" & rsprov_prod.Fields("F2NOMPRV")
                .FieldValues("PRECIO") = Format(Val("" & rsprov_prod.Fields("F5VALVTA")), "###,###,##0.00")
                .FieldValues("CODFAB") = "" & rsprov_prod.Fields("F5CODFAB")
                .FieldValues("FPAGO") = "" & rsprov_prod.Fields("F2FORPAG")
                sw_nuevo_item = False
                rsprov_prod.MoveNext
            Loop
        .Post
        End With
        dxDBGrid1.Dataset.EnableControls
        dxDBGrid1.Dataset.Open
        dxDBGrid1.OptionEnabled = True
    Else
        dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
        dxDBGrid1.Dataset.Active = True
        dxDBGrid1.Dataset.Close
        dxDBGrid1.Dataset.Open
        dxDBGrid1.OptionEnabled = False
        dxDBGrid1.Dataset.DisableControls
        With dxDBGrid1.Dataset
            sw_nuevo_item = True
            .Append
            .FieldValues("ITEM") = 1
            .FieldValues("RUCPROV") = ""
            .FieldValues("NOMPROV") = ""
            .FieldValues("PRECIO") = "0.00"
            .FieldValues("CODFAB") = ""
            .FieldValues("FPAGO") = ""
            .Post
            sw_nuevo_item = False
        End With
        dxDBGrid1.Dataset.EnableControls
        dxDBGrid1.Dataset.Open
        dxDBGrid1.OptionEnabled = True
    End If
    rsprov_prod.Close
    
End Sub

Private Sub GRABA_PROD()
Dim csql        As String
Dim ccodigo     As String
Dim cnombre     As String
Dim nprecio     As Double
Dim ccodfab     As String
Dim cmedida     As String
Dim cruc        As String
Dim cnombreprov As String
Dim cformapago  As String

    cnn_dbbancos.Execute ("DELETE * FROM EF2PROD_PROV WHERE F5CODPRO='" & txtproducto.Text & "'")

    If rsproductos.State = adStateOpen Then rsproductos.Close
    rsproductos.Open "SELECT F7CODMED,F5CODFAB FROM IF5PLA WHERE F5CODPRO='" & txtproducto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsproductos.EOF Then
        cmedida = rsproductos.Fields("F7CODMED") & ""
        ccodfab = rsproductos.Fields("F5CODFAB") & ""
    End If
    rsproductos.Close
    
    ccodigo = txtproducto.Text
    cnombre = pnlproducto.Caption

    If rstemporal.State = adStateOpen Then rstemporal.Close
    rstemporal.Open "SELECT * FROM DETALLE_PROV", cnn_form, adOpenDynamic, adLockOptimistic
    If Not rstemporal.EOF Then
        rstemporal.MoveFirst
        Do While Not rstemporal.EOF
            If Len(Trim(rstemporal.Fields("RUCPROV") & "")) > 0 Then
                cruc = rstemporal.Fields("RUCPROV") & ""
                nprecio = Val(rstemporal.Fields("PRECIO") & "")
                cnombreprov = rstemporal.Fields("NOMPROV") & ""
                cformapago = rstemporal.Fields("FPAGO") & ""
                
                csql = "INSERT INTO EF2PROD_PROV (F2CODPRV,F2NOMPRV,F5CODPRO,F5NOMPRO,F5VALVTA,F5CODFAB,F7CODMED,F2FORPAG) " & _
                       " VALUES('" & cruc & "','" & cnombreprov & "','" & ccodigo & "','" & _
                        cnombre & "'," & nprecio & ",'" & ccodfab & "','" & cmedida & "','" & cformapago & "')"
                cnn_dbbancos.Execute (csql)
            End If
            rstemporal.MoveNext
        Loop
    End If
    rstemporal.Close
            
End Sub

Private Sub txtproducto_DblClick()

    txtproducto_KeyDown 113, 0

End Sub

Private Sub txtproducto_GotFocus()

    txtproducto.SelStart = 0: txtproducto.SelLength = Len(txtproducto.Text)
    
End Sub

Private Sub txtproducto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcod_alm = ""
        wcodproducto = ""
        sw_ayuda_prod = True
        'hlp_productos.Show 1
        ayuda_productos.Show 1
        If Len(Trim(wcodproducto)) > 0 Then
            txtproducto.Text = wcodproducto
            pnlproducto.Caption = wdesproducto
            txtproducto_KeyPress 13
        End If
    End If

End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        dxDBGrid1.SetFocus
    End If
    
End Sub

Private Sub txtproducto_LostFocus()

    If Len(Trim(txtproducto.Text)) > 0 Then
        If rsproductos.State = adStateOpen Then rsproductos.Close
        rsproductos.Open "SELECT F5NOMPRO FROM IF5PLA WHERE F5CODPRO='" & txtproducto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsproductos.EOF Then
            pnlproducto.Caption = rsproductos.Fields("F5NOMPRO") & ""
        End If
        rsproductos.Close
        LLENA_PRODUCTOS txtproducto.Text
    End If
    
End Sub
