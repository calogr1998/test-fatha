VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cons_comprasxproveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Compras x Proveedor"
   ClientHeight    =   7095
   ClientLeft      =   3345
   ClientTop       =   3345
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11730
   Begin VB.TextBox txtrucprov 
      Height          =   285
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   9
      Top             =   450
      Width           =   1335
   End
   Begin VB.CheckBox CheckFiltro 
      Caption         =   "Activar Filtro"
      Height          =   255
      Left            =   405
      TabIndex        =   6
      Top             =   1035
      Width           =   1455
   End
   Begin VB.CheckBox Checkagrupar 
      Caption         =   "Agrupar columnas"
      Height          =   255
      Left            =   1890
      TabIndex        =   5
      Top             =   1035
      Width           =   2055
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   " Datos del Proveedor y rango de fechas "
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
      Begin VB.TextBox txtnomprov 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker txtdesde 
         Height          =   315
         Left            =   8400
         TabIndex        =   10
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96534529
         CurrentDate     =   40611
      End
      Begin MSComCtl2.DTPicker txthasta 
         Height          =   315
         Left            =   9960
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96534529
         CurrentDate     =   40611
      End
      Begin VB.Label Label1 
         Caption         =   "RUC Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
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
         Left            =   8640
         TabIndex        =   2
         Top             =   120
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
         Left            =   10320
         TabIndex        =   1
         Top             =   120
         Width           =   420
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5640
      Left            =   120
      TabIndex        =   3
      Top             =   1395
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   9948
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
         Height          =   5145
         Left            =   90
         OleObjectBlob   =   "cons_comprasxproveedor.frx":0000
         TabIndex        =   4
         Top             =   120
         Width           =   11310
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
      Tools           =   "cons_comprasxproveedor.frx":452F
      ToolBars        =   "cons_comprasxproveedor.frx":C41F
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   810
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save to"
      FileName        =   "GridNum"
   End
End
Attribute VB_Name = "cons_comprasxproveedor"
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
Dim sw_ayuda            As Boolean

Public Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim i As Byte
    
    If Ind > 199 Then
        SaveTo (Ind)
        Exit Sub
    End If
    
End Sub

Public Sub SaveTo(Index)
On Error GoTo errhandler
Dim FileName As String

    If GridNum <> 0 Then
        With cmdSave
            .CancelError = True
            .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNOverwritePrompt
            '.DialogTitle = menu.dxSideBar1.StuckLink.Item.Caption
            .DialogTitle = "Ordenes de Compra"
            Select Case Index
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
    
errhandler:
    
    Exit Sub
 
End Sub

Public Function GetGridByActive() As dxDBGrid
    
    Set GetGridByActive = dxDBGrid1
    
End Function

Private Sub CheckFiltro_Click()
    If CheckFiltro.value = 1 Then
      dxDBGrid1.Filter.FilterActive = True
    Else
      dxDBGrid1.Filter.FilterActive = False
    End If
End Sub

Private Sub Form_Activate()
    txtrucprov.SetFocus
End Sub

Private Sub Form_Load()
Dim CadSql      As String
    
    Me.MousePointer = vbHourglass
    dxDBGrid1.Options.Unset (egoShowGroupPanel)
    dxDBGrid1.Filter.FilterActive = False
    
    Me.left = 100
    Me.top = 1150

    Me.MousePointer = vbHourglass

    'cnombase = wusuario & "OCOMPRA" & Format(Time, "hh_mm_ss") & ".MDB"
    '--- conexion a la base de datos temporal --------'
    'CREATEDATABASE_N wrutatemp & "\", cnombase
    cnombase = "TEMPLUS.MDB"
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_form.Open cconex_form
    cnomtabla = "COMPRASXPROVEEDOR"
    
    'CadSql = "(ITEM TEXT(4),OCOMPRA TEXT(10),SOLSUMINISTRO TEXT(4),USUARIO TEXT(8),FECHA DATE,PROVEEDOR TEXT(100)," _
            & " OBRA TEXT(8),NOMOBRA TEXT(100),OBSERVACIONES TEXT(100))"
    'CREATETABLE_N cnomtabla, CadSql, cnn_form

    TxtDesde.value = Format(Date, "dd/mm/yyyy")
    TxtHasta.value = Format(Date, "dd/mm/yyyy")
    
    LLENA_TEMPORAL
    
    Me.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub LLENA_TEMPORAL()
Dim X       As Integer
Dim csql    As String

    DELETEREC_LOG cnomtabla, cnn_form

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cnn_form
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls
    With dxDBGrid1.Dataset
        sw_nuevo_item = True
        X = 1
        If ctipoadm_bd = "M" Then
            csql = csql + "SELECT IF4VALES.F2CODPROV, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F1CODORI, IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, IF3VALES.F3CANPRO, IF5PLA.F7CODMED, IF4VALES.F4MONEDA, IF3VALES.F3VALVTA, IF3VALES.F3VALDOL "
            csql = csql + "FROM (IF4VALES INNER JOIN EF2PROVEEDORES ON IF4VALES.F2CODPROV = EF2PROVEEDORES.F2NEWRUC) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) "
            csql = csql + "WHERE IF4VALES.F2CODPROV='" & txtrucprov.Text & "'  AND IF4VALES.F1CODORI='XC0' AND (IF4VALES.F4FECVAL) >= '" & TxtDesde.value & "' And (IF4VALES.F4FECVAL) <= '" & TxtHasta.value & "' ;"
        Else
            csql = csql + "SELECT IF4VALES.F2CODPROV, EF2PROVEEDORES.F2NOMPROV, IF4VALES.F1CODORI, IF4VALES.F2CODALM, IF4VALES.F4NUMVAL, IF4VALES.F4FECVAL, IF3VALES.F5CODPRO, IF5PLA.F5NOMPRO, IF3VALES.F3CANPRO, IF5PLA.F7CODMED, IF4VALES.F4MONEDA, IF3VALES.F3VALVTA, IF3VALES.F3VALDOL "
            csql = csql + "FROM (IF4VALES INNER JOIN EF2PROVEEDORES ON IF4VALES.F2CODPROV = EF2PROVEEDORES.F2NEWRUC) INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM) "
            csql = csql + "WHERE IF4VALES.F2CODPROV='" & txtrucprov.Text & "'  AND IF4VALES.F1CODORI='XC0' AND CVDATE(IF4VALES.F4FECVAL) >= CVDATE('" & TxtDesde.value & "') And CVDATE(IF4VALES.F4FECVAL) <= CVDATE('" & TxtHasta.value & "') ;"
                        
            rsif4orden.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        End If
        If Not rsif4orden.EOF Then
            Do While Not rsif4orden.EOF
                .Append
                .FieldValues("ITEM") = X
                .FieldValues("F4FECVAL") = "" & Format(rsif4orden.Fields("F4FECVAL"), "dd/mm/yyyy")
                .FieldValues("F2CODALM") = Format("" & rsif4orden.Fields("F2CODALM"), "00")
                .FieldValues("F4NUMVAL") = "" & rsif4orden.Fields("F4NUMVAL")
                .FieldValues("F5CODPRO") = "" & rsif4orden.Fields("F5CODPRO")
                .FieldValues("F5NOMPRO") = "" & rsif4orden.Fields("F5NOMPRO")
                .FieldValues("F2CODPROV") = rsif4orden.Fields("F2CODPROV") & ""
                .FieldValues("F7CODMED") = "" & rsif4orden.Fields("F7CODMED")
                .FieldValues("F3CANPRO") = rsif4orden.Fields("F3CANPRO")
                .FieldValues("F3VALVTA") = IIf(rsif4orden.Fields("F4MONEDA") = "S", rsif4orden.Fields("F3VALVTA"), Null)
                .FieldValues("F3VALDOL") = IIf(rsif4orden.Fields("F4MONEDA") = "D", rsif4orden.Fields("F3VALDOL"), Null)
                
                rsif4orden.MoveNext
                X = X + 1
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

    Select Case Tool.Id
        Case "ID_Actualizar"
            Me.MousePointer = vbHourglass
            LLENA_TEMPORAL
            Me.MousePointer = vbDefault
        Case "ID_Preliminar"
            IValue = SSActiveToolBars1.Tools.ITEM("ID_Preliminar").UseMaskColor
            GridNum = 1: OldValue = 1
            GridInit IValue, OldValue
            OldValue = IValue
        Case "ID_ExportExcell"
            IValue = SSActiveToolBars1.Tools.ITEM("ID_ExportExcell").UseMaskColor
            GridNum = 1: OldValue = 1
            GridInit IValue - 10, OldValue
            OldValue = IValue
        Case "ID_Imprimir"
        
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TxtHasta.SetFocus
    End If

End Sub

Private Sub txtrucprov_DblClick()
    txtrucprov_KeyDown 113, 0
End Sub

Private Sub txtrucprov_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw_ayuda = True
        wcodprov = "": wrucprov = "": wnomprov = ""
        Ayuda_Proveedores.Show 1
        'If Len(Trim(wrucprov)) > 0 Then
        If Len(Trim(wRucCliProv)) > 0 Then
            txtrucprov.Text = wRucCliProv 'wrucprov
            txtnomprov.Text = wnomcliprov 'wnomprov
        End If
        sw_ayuda = False
    End If
End Sub

Private Sub txtrucprov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
End Sub

Private Sub txtrucprov_LostFocus()
    If sw_ayuda = False Then
        If Len(Trim(txtrucprov.Text)) > 0 Then
            If VALIDA_PROVEEDOR(txtrucprov.Text) = True Then
                txtnomprov.Text = wnomprov
            Else
                MsgBox "El proveedor no existe. Verifique.", vbCritical, "Atención"
                txtrucprov.SetFocus
            End If
        End If
    End If
End Sub
Private Function VALIDA_PROVEEDOR(pproveedor As String)
Dim sw_e    As Boolean

    If RsProveedor.State = adStateOpen Then RsProveedor.Close
    RsProveedor.Open "SELECT F2NOMPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & pproveedor & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProveedor.EOF Then
        wnomprov = Trim(RsProveedor.Fields("F2NOMPROV") & "")
        sw_e = True
    Else
        sw_e = False
    End If
    RsProveedor.Close
    VALIDA_PROVEEDOR = sw_e

End Function
