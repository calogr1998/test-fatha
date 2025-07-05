VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cons_ordenes_compra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Ordenes de Compra"
   ClientHeight    =   8955
   ClientLeft      =   -15
   ClientTop       =   2070
   ClientWidth     =   19020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   19020
   Begin VB.TextBox txtbusqueda 
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   240
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog cmdsave 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   " Rango de fechas "
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin MSComCtl2.DTPicker txtdesde 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   99680257
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker txthasta 
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   99680257
         CurrentDate     =   40611
      End
      Begin VB.Label lblfecemi 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Left            =   465
         TabIndex        =   2
         Top             =   405
         Width           =   465
      End
      Begin VB.Label lblfecven 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Left            =   3120
         TabIndex        =   1
         Top             =   405
         Width           =   420
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   7800
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   18855
      _Version        =   65536
      _ExtentX        =   33258
      _ExtentY        =   13758
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
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   7425
         Left            =   240
         OleObjectBlob   =   "cons_ordenes_compra.frx":0000
         TabIndex        =   4
         Top             =   240
         Width           =   18375
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "cons_ordenes_compra.frx":83FE
      ToolBars        =   "cons_ordenes_compra.frx":10274
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox2 
      Height          =   270
      Left            =   6600
      TabIndex        =   10
      Top             =   600
      Width           =   1335
      _Version        =   65536
      _cx             =   2355
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Solo Servicios"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox1 
      Height          =   270
      Left            =   16680
      TabIndex        =   9
      Top             =   360
      Width           =   1785
      _Version        =   65536
      _cx             =   3149
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Mostrar Presupuesto"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   15790320
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin VB.Label Label1 
      Caption         =   "Filtro / Búsqueda"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "cons_ordenes_compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnombase            As String
Dim cnomtabla           As String
Dim cconex_form         As String
Dim cnn_form            As New ADODB.Connection
Dim GridNum             As Byte
Dim OldValue            As Byte
Dim sw_nuevo_item       As Boolean

Public Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim i As Byte
    
    If Ind > 199 Then
        SaveTo (Ind)
        Exit Sub
    End If
    
End Sub

Public Sub SaveTo(index)
On Error GoTo errHandler
Dim FileName As String

    If GridNum <> 0 Then
        With cmdsave
            .CancelError = True
            .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
            '.DialogTitle = menu.dxSideBar1.StuckLink.Item.Caption
            .DialogTitle = "Ordenes de Compra"
            Select Case index
                Case 204
                    .Filter = "Text Files (*.txt)|*.txt"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    If GetGridByActive().Ex.SelectedCount = 0 Then
                        GetGridByActive().m.SaveAllToTextFile (FileName)
                    Else
                        GetGridByActive().m.SaveSelectedToTextFile (FileName)
                    End If
                Case 245
                    .Filter = "Excel Files (*.xls)|*.xls"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().m.ExportToXLS FileName
                Case 202
                    .Filter = "HTML Files (*.htm)|*.htm"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().m.ExportToHTML FileName
                Case 205
                    .Filter = "XML Files (*.xml)|*.xml"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().m.ExportToXML FileName
                Case 201
                    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Sistema de Logistica") = vbYes Then _
                        GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), False
                Case 255
                    GetGridByActive().m.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), True
            End Select
        End With
    End If
    
errHandler:
    
    Exit Sub
 
End Sub

Public Function GetGridByActive() As dxDBGrid
    
    Set GetGridByActive = dxDBGrid1
    
End Function

Private Sub dxCheckBox1_Click()
    If dxCheckBox1.Checked = 1 Then
        dxDBGrid1.Columns.ColumnByFieldName("Partida").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("DESCPARTIDA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("QPARTIDA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("UMPARTIDA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("UNIPDA").Visible = True
        dxDBGrid1.Columns.ColumnByFieldName("TOTPDA").Visible = True
    Else
        dxDBGrid1.Columns.ColumnByFieldName("Partida").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("DESCPARTIDA").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("QPARTIDA").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("UMPARTIDA").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("UNIPDA").Visible = False
        dxDBGrid1.Columns.ColumnByFieldName("TOTPDA").Visible = False
    End If
End Sub

Private Sub Form_Load()
Dim CadSql      As String
Dim Fecha As String
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    
    Me.left = 1650
    Me.top = 1050

    Me.MousePointer = vbHourglass

    'cnombase = wusuario & "OCOMPRA" & Format(Time, "hh_mm_ss") & ".MDB"
    '--- conexion a la base de datos temporal --------'
    'CREATEDATABASE_N wrutatemp & "\", cnombase
    cnombase = "TEMPLUS.MDB"
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    cnomtabla = "OC_CONSULTA"
    
    'CadSql = "(ITEM TEXT(4),OCOMPRA TEXT(10),SOLSUMINISTRO TEXT(4),USUARIO TEXT(8),FECHA DATE,PROVEEDOR TEXT(100)," _
            & " OBRA TEXT(8),NOMOBRA TEXT(100),OBSERVACIONES TEXT(100))"
    'CREATETABLE_N cnomtabla, CadSql, cnn_form
    Fecha = "01/01/" & Year(Date)
    
    txtdesde.Value = Format(Fecha, "dd/mm/yyyy")
    txthasta.Value = Format(Date, "dd/mm/yyyy")
    
    'LLENA_TEMPORAL
    
    Me.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub LLENA_TEMPORAL()
Dim x       As Integer
    
    DELETEREC_LOG cnomtabla, cnn_form

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.Filter.Clear
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls
    With dxDBGrid1.Dataset
        sw_nuevo_item = True
        x = 1
        If dxCheckBox2.Checked = 1 Then
            csql = "SELECT CENTROS.F3COSTO, CENTROS.F3ABREV, IF3ORDEN.PARTIDA, dbo_Partida.Descripcion, dbo_Partida.Material1, "
            csql = csql & "dbo_Partida.CodUnidad, dbo_Partida.Precio1, dbo_Partida.TotalMaterial1, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4TIPMON, IF3ORDEN.F3CODPRO, IF3ORDEN.F5NOMPRO, IF3ORDEN.F3CANPRO, EF7MEDIDAS.F7SIGMED, IF4ORDEN.F4NOMPROV, val( IIf( IF4ORDEN.F4TIPMON='S',IF3ORDEN.F3PRECOS,Format(IF3ORDEN.F3PRECOS* IF4ORDEN.F4TIPCAM,'0.00'))) AS F3PREUNI,  Val(Format(Val([IF3ORDEN].[F3CANPRO]*[IF3ORDEN].[F3PRECOS])*IIf([IF4ORDEN].[F4TIPMON]='S',1,[IF4ORDEN].[F4TIPCAM]),'0.00')) AS F3TOTAL, REGISDOC.F4FECHA, IIf(REGISDOC.f4moneda='S',REGISDOC.F4TOTAL,0) AS F4TOTAL, PAG_DCTO.nro_comp, PAG_DCTO.total, IIf(REGISDOC.f4moneda='D',REGISDOC.F4TOTAL,0) AS SALDO, IF4ORDEN.F4OBSERVA as cs_observaciones "
            csql = csql & "FROM ((REGISDOC LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) RIGHT JOIN (IF4ORDEN LEFT JOIN CENTROS ON IF4ORDEN.F4CENTRO = CENTROS.F3COSTO) ON REGISDOC.F4OCOMPRA = IF4ORDEN.F4NUMORD) INNER JOIN ((((IF3ORDEN INNER JOIN IF5PLA ON IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO) INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) LEFT JOIN TB_CABSOLICITUD ON IF3ORDEN.COD_SOLICITUD = TB_CABSOLICITUD.cod_solicitud) LEFT JOIN dbo_Partida ON (IF3ORDEN.F3CENCOS = dbo_Partida.CodPresupuesto) AND (IF3ORDEN.PARTIDA = dbo_Partida.CodPartida)) ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
            csql = csql & "WHERE IF4ORDEN.F4LOCAL='OS' AND IF4ORDEN.F4ESTNUL='N' AND CVDATE(F4FECEMI) >= CVDate('" & txtdesde.Value & "') and cvdate(F4FECEMI) <=  CVDate('" & txthasta.Value & "') "
            If Len(txtbusqueda.Text) > 0 Then
                csql = csql & " AND IF5PLA.F5NOMPRO LIKE '*" & txtbusqueda.Text & "*'"
            End If
            csql = csql & " ORDER BY IF4ORDEN.F4NUMORD"
        Else
            csql = "SELECT CENTROS.F3COSTO, CENTROS.F3ABREV, IF3ORDEN.PARTIDA, dbo_Partida.Descripcion, dbo_Partida.Material1, "
            csql = csql & "dbo_Partida.CodUnidad, dbo_Partida.Precio1, dbo_Partida.TotalMaterial1, IF4ORDEN.F4NUMORD, IF4ORDEN.F4FECEMI, IF4ORDEN.F4TIPMON, IF3ORDEN.F3CODPRO, IF3ORDEN.F5NOMPRO, IF3ORDEN.F3CANPRO, EF7MEDIDAS.F7SIGMED, IF4ORDEN.F4NOMPROV, val( IIf( IF4ORDEN.F4TIPMON='S',IF3ORDEN.F3PRECOS,Format(IF3ORDEN.F3PRECOS* IF4ORDEN.F4TIPCAM,'0.00'))) AS F3PREUNI,  Val(Format(Val([IF3ORDEN].[F3CANPRO]*[IF3ORDEN].[F3PRECOS])*IIf([IF4ORDEN].[F4TIPMON]='S',1,[IF4ORDEN].[F4TIPCAM]),'0.00')) AS F3TOTAL, REGISDOC.F4FECHA, IIf(REGISDOC.f4moneda='S',REGISDOC.F4TOTAL,0) AS F4TOTAL, PAG_DCTO.nro_comp, PAG_DCTO.total, IIf(REGISDOC.f4moneda='D',REGISDOC.F4TOTAL,0) AS SALDO, IF4ORDEN.F4OBSERVA as cs_observaciones "
            csql = csql & "FROM ((REGISDOC LEFT JOIN PAG_DCTO ON REGISDOC.F4CORRELA = PAG_DCTO.correla) RIGHT JOIN (IF4ORDEN LEFT JOIN CENTROS ON IF4ORDEN.F4CENTRO = CENTROS.F3COSTO) ON REGISDOC.F4OCOMPRA = IF4ORDEN.F4NUMORD) INNER JOIN ((((IF3ORDEN INNER JOIN IF5PLA ON IF3ORDEN.F3CODPRO = IF5PLA.F5CODPRO) INNER JOIN EF7MEDIDAS ON IF5PLA.F7CODMED = EF7MEDIDAS.F7CODMED) LEFT JOIN TB_CABSOLICITUD ON IF3ORDEN.COD_SOLICITUD = TB_CABSOLICITUD.cod_solicitud) LEFT JOIN dbo_Partida ON (IF3ORDEN.F3CENCOS = dbo_Partida.CodPresupuesto) AND (IF3ORDEN.PARTIDA = dbo_Partida.CodPartida)) ON (IF4ORDEN.F4NUMORD = IF3ORDEN.F4NUMORD) AND (IF4ORDEN.F4LOCAL = IF3ORDEN.F4LOCAL) "
            csql = csql & "WHERE IF4ORDEN.F4ESTNUL='N' AND CVDATE(F4FECEMI) >= CVDate('" & txtdesde.Value & "') and cvdate(F4FECEMI) <=  CVDate('" & txthasta.Value & "') "
            If Len(txtbusqueda.Text) > 0 Then
                csql = csql & " AND IF5PLA.F5NOMPRO LIKE '*" & txtbusqueda.Text & "*'"
            End If
            csql = csql & " ORDER BY IF4ORDEN.F4NUMORD"
        End If
        If rsif4orden.State = 1 Then rsif4orden.Close
        rsif4orden.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif4orden.EOF Then
            Do While Not rsif4orden.EOF
                .Append
                .FieldValues("ITEM") = x
                .FieldValues("F3ABREV") = "" & rsif4orden.Fields("F3ABREV")
                .FieldValues("F4NUMORD") = "" & rsif4orden.Fields("F4NUMORD")
                .FieldValues("F4FECEMI") = "" & Format(rsif4orden.Fields("f4fecemi"), "dd/mm/yyyy")
                .FieldValues("F3CODPRO") = "" & rsif4orden.Fields("F3CODPRO")
                .FieldValues("F5NOMPRO") = "" & rsif4orden.Fields("F5NOMPRO")
                .FieldValues("F3CANPRO") = "" & rsif4orden.Fields("F3CANPRO")
                .FieldValues("F7SIGMED") = "" & rsif4orden.Fields("F7SIGMED")
                .FieldValues("F4NOMPRV") = "" & rsif4orden.Fields("F4NOMPROV")
                .FieldValues("F4TIPMON") = "" & rsif4orden.Fields("F4TIPMON")
                .FieldValues("F3PREUNI") = "" & rsif4orden.Fields("F3PREUNI")
                .FieldValues("F3TOTAL") = "" & rsif4orden.Fields("F3TOTAL")
                .FieldValues("F4FECHA") = "" & Format(rsif4orden.Fields("F4FECHA"), "dd/mm/yyyy")
                .FieldValues("nro_comp") = "" & rsif4orden.Fields("nro_comp")
                .FieldValues("F4TOTAL") = "" & rsif4orden.Fields("F4TOTAL")
                .FieldValues("total") = "" & rsif4orden.Fields("total")
                .FieldValues("saldo") = "" & rsif4orden.Fields("saldo")
                .FieldValues("Partida") = "" & rsif4orden.Fields("Partida")
                .FieldValues("DESCPARTIDA") = "" & rsif4orden.Fields("Descripcion")
                .FieldValues("QPARTIDA") = "" & rsif4orden.Fields("Material1")
                .FieldValues("UMPARTIDA") = "" & rsif4orden.Fields("CodUnidad")
                .FieldValues("UNIPDA") = "" & rsif4orden.Fields("Precio1")
                .FieldValues("TOTPDA") = "" & rsif4orden.Fields("TotalMaterial1")
                .FieldValues("cs_observacion") = "" & rsif4orden.Fields("cs_observaciones")
                rsif4orden.MoveNext
                x = x + 1
                .Post
            Loop
        End If
        rsif4orden.Close

        sw_nuevo_item = False
    End With

    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open

    dxDBGrid1.OptionEnabled = True


End Sub

Private Sub Form_Unload(Cancel As Integer)

    cnn_form.Close
    
    dxDBGrid1.Dataset.Close
    ELIMINA_BD_N wrutatemp & "\", cnombase
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim n       As Byte
Dim IValue  As Byte

    Select Case Tool.ID
        Case "ID_Actualizar"
            Me.MousePointer = vbHourglass
            LLENA_TEMPORAL
            Me.MousePointer = vbDefault
        Case "ID_Preliminar"
'            IValue = SSActiveToolBars1.Tools.ITEM("ID_Preliminar").UseMaskColor
'            GridNum = 1: OldValue = 1
'            GridInit IValue, OldValue
'            OldValue = IValue
            'If Tool.State = ssChecked Then
                dxDBGrid1.Options.Set (egoShowGroupPanel)
            'Else
            '    dxDBGrid1.Options.Unset (egoShowGroupPanel)
            'End If
        Case "ID_ExportExcell"
            IValue = SSActiveToolBars1.Tools.ITEM("ID_ExportExcell").UseMaskColor
            GridNum = 1: OldValue = 1
            GridInit IValue - 10, OldValue
            OldValue = IValue
        Case "ID_Imprimir"
        
        Case "ID_Filtro"
            If dxDBGrid1.Filter.FilterActive = True Then
                dxDBGrid1.Filter.FilterActive = False
                'Me.Toolbar.Buttons.ITEM(11).Image = 3
                'Me.Toolbar.Buttons.ITEM(11).ToolTipText = "Activar Filtro"
            Else
                dxDBGrid1.Filter.FilterActive = True
                'Me.Toolbar.Buttons.ITEM(11).Image = 6
                'Me.Toolbar.Buttons.ITEM(11).ToolTipText = "Desactivar Filtro"
            End If
'            If CheckFiltro.Value = 1 Then
'              dxDBGrid1.Filter.FilterActive = True
'            Else
'              dxDBGrid1.Filter.FilterActive = False
'            End If
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub txtbusqueda_Change()
    If Len(txtbusqueda.Text) > 2 Then
        dxDBGrid1.Dataset.Filtered = True
        'If Len(dxDBGrid1.Filter.FilterText) > 0 Then
        dxDBGrid1.Dataset.Filter = "F5NOMPRO LIKE '*" & txtbusqueda.Text & "*' OR " & "F4NOMPRV LIKE '*" & txtbusqueda.Text & "*' "
        'Else
            'dxDBGrid1.Dataset.Filter = False
        'End If
        
    End If
    If Len(Trim(txtbusqueda.Text)) = 0 Then
                dxDBGrid1.Dataset.Filtered = False
    End If
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txthasta.SetFocus
    End If

End Sub

