VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Consulta_PreciosdeProductos 
   Caption         =   "Consulta de precios"
   ClientHeight    =   8115
   ClientLeft      =   615
   ClientTop       =   645
   ClientWidth     =   11085
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
   ScaleHeight     =   8115
   ScaleWidth      =   11085
   Begin Threed.SSFrame SSFrame1 
      Height          =   870
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   10860
      _Version        =   65536
      _ExtentX        =   19156
      _ExtentY        =   1535
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
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   1125
         TabIndex        =   0
         Top             =   315
         Width           =   1380
      End
      Begin Threed.SSPanel pnlnompro 
         Height          =   300
         Left            =   2610
         TabIndex        =   3
         Top             =   315
         Width           =   8130
         _Version        =   65536
         _ExtentX        =   14340
         _ExtentY        =   529
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   210
         Left            =   270
         TabIndex        =   4
         Top             =   360
         Width           =   645
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6540
      Left            =   135
      OleObjectBlob   =   "Consulta_PreciosdeProductos.frx":0000
      TabIndex        =   1
      Top             =   1080
      Width           =   10800
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "Consulta_PreciosdeProductos.frx":3A4C
      ToolBars        =   "Consulta_PreciosdeProductos.frx":B940
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save to"
      FileName        =   "GridNum"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   945
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save to"
      FileName        =   "GridNum"
   End
End
Attribute VB_Name = "Consulta_PreciosdeProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn_temp        As New ADODB.Connection
Dim rs              As New ADODB.Recordset
Dim a               As String
Dim b               As String
Dim c               As String
Dim rs2             As New ADODB.Recordset
Dim rs3             As New ADODB.Recordset
Dim cconex_temp     As String
Dim GridNum         As Byte
Dim OldValue        As Byte

Public Sub GridInit(ByVal Ind As Byte, ByVal IndOld As Byte)
Dim I As Byte
    
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
            .DialogTitle = "Lista de Precios"
            Select Case Index
                Case 204
                    .Filter = "Text Files (*.txt)|*.txt"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    If GetGridByActive().Ex.SelectedCount = 0 Then
                        GetGridByActive().M.SaveAllToTextFile (FileName)
                    Else
                        GetGridByActive().M.SaveSelectedToTextFile (FileName)
                    End If
                Case 245
                    .Filter = "Excel Files (*.xls)|*.xls"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().M.ExportToXLS FileName
                Case 202
                    .Filter = "HTML Files (*.htm)|*.htm"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().M.ExportToHTML FileName
                Case 205
                    .Filter = "XML Files (*.xml)|*.xml"
                    .FileName = ""
                    .ShowSave
                    FileName = .FileName
                    GetGridByActive().M.ExportToXML FileName
                Case 201
                    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then _
                        GetGridByActive().M.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), False
                Case 255
                    GetGridByActive().M.PrintControl GetGridByActive().Options.Contains(egoAutoWidth), True
            End Select
        End With
    End If
    
errhandler:
    
    Exit Sub
 
End Sub

Public Function GetGridByActive() As dxDBGrid
    
    Set GetGridByActive = dxDBGrid1
    
End Function

Private Sub Form_Load()
Dim csql        As String
       
    cnombase = wusuario & "Producto" & Format(Time, "hh_mm_ss") & ".MDB"
    CREATEDATABASE_N wrutatemp & "\", cnombase
    
    cconex_temp = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    cnn_temp.Open cconex_temp
    
    cnomtabla = "Productos"
    csql = "(ITEM Text(4),F4NUMORD Text(15),F4FECEMI Date,F4CODPRV Text(100),F4TIPMON Text(3),F3CANPRO Text(5),F3PREBRUS Double,F3PREBRUD Double)"
    
    CREATETABLE_N cnomtabla, csql, cnn_temp

End Sub

Private Sub Llenado_Registros()
Dim ncont           As Integer
Dim cnomprov        As String
Dim csql            As String
    
    dxDBGrid1.Dataset.Close
    DELETEREC_N cnomtabla, cnn_temp
    DELETEREC_N cnomtabla, cnn_temp

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cconex_temp
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls

    With dxDBGrid1.Dataset
        ncont = 1
        csql = "Select f3codpro,f4numord,f3canpro,f3precos from if3orden where f3codpro='" & a & "'"
        If rs.State = adStateOpen Then rs.Close
        rs.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Do While Not rs.EOF
                csql = "Select * from if4orden where f4numord=" & rs.Fields("f4numord") & ""
                If rs2.State = adStateOpen Then rs2.Close
                rs2.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rs2.EOF Then
                    .Append
                    .FieldValues("ITEM") = ncont
                    .FieldValues("F4NUMORD") = rs2.Fields("F4NUMORD")
                    .FieldValues("F4FECEMI") = Format(rs2.Fields("F4FECEMI"), "DD/MM/YYYY")
                    cnomprov = ""
                    csql = "Select f2nomprov from ef2proveedores where f2newruc='" & rs2.Fields("F4CODPRV") & "'"
                    If rs3.State = adStateOpen Then rs3.Close
                    rs3.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rs3.EOF Then
                        cnomprov = Trim("" & rs3.Fields("f2nomprov"))
                    End If
                    rs3.Close
                    .FieldValues("F4CODPRV") = cnomprov
                    .FieldValues("F4TIPMON") = Trim("" & rs2.Fields("F4TIPMON"))
                    .FieldValues("F3CANPRO") = Format(Val(rs.Fields("F3CANPRO") & ""), "###,###,##0.00")
                    If Trim("" & rs2.Fields("F4TIPMON")) = "S" Then
                        .FieldValues("F3PREBRUS") = Format(Val("" & rs.Fields("F3PRECOS")) * Val("" & rs.Fields("F3CANPRO")), "###,###,##0.00")
                        If Val("" & rs2.Fields("F4TIPCAM")) > 0 Then
                            .FieldValues("F3PREBRUD") = Format((Val("" & rs.Fields("F3PRECOS")) * Val("" & rs.Fields("F3CANPRO"))) / Val("" & rs2.Fields("F4TIPCAM")), "###,###,##0.00")
                        Else
                            .FieldValues("F3PREBRUD") = Format(0, "###,###,##0.00")
                        End If
                    Else
                        .FieldValues("F3PREBRUD") = Format(Val("" & rs.Fields("F3PRECOS")) * Val("" & rs.Fields("F3CANPRO")), "###,###,##0.00")
                        .FieldValues("F3PREBRUS") = Format((Val("" & rs.Fields("F3PRECOS")) * Val("" & rs.Fields("F3CANPRO"))) * Val("" & rs2.Fields("F4TIPCAM")), "###,###,##0.00")
                    End If
                    ncont = ncont + 1
                End If
                rs2.Close
                
                rs.MoveNext
    
            Loop
            .Post
        Else
            MsgBox "No existen registros para ser procesados.", vbCritical, "Atención"
        End If
    End With

    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    cnn_temp.Close
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim n       As Byte
Dim IValue  As Byte

    Select Case Tool.Id
        Case "ID_Preliminar"
            IValue = SSActiveToolBars1.Tools.item("ID_Preliminar").UseMaskColor
            GridNum = 1: OldValue = 1
            GridInit IValue, OldValue
            OldValue = IValue
        Case "ID_ExportExcell"
            IValue = SSActiveToolBars1.Tools.item("ID_ExportExcell").UseMaskColor
            GridNum = 1: OldValue = 1
            GridInit IValue - 10, OldValue
            OldValue = IValue
        Case "ID_Salir"
            Unload Me
    End Select
    
End Sub

Private Sub txtcodigo_DblClick()

    TxtCodigo_KeyDown 113, 0

End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        wcod_alm = ""
        hlp_productos.Show 1
        If Len(Trim(wcodproducto)) > 0 Then
            txtcodigo.Text = wcodproducto
            pnlnompro.Caption = wdesproducto
            txtcodigo_KeyPress 13
        End If
    End If

End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(Trim(txtcodigo.Text)) > 0 Then
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select f5codpro,f5nompro from if5pla where f5codpro='" & txtcodigo.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Len(rs.Fields(0)) > 0 Then
                    a = Trim$(rs.Fields("f5codpro"))
                    pnlnompro.Caption = Trim(rs.Fields("f5nompro") & "")
                    'Llenado_Registros
                    CONSULTA_VALES
                End If
            Else
                MsgBox "Código del producto no existe. Verifique.", vbCritical, "Atención"
            End If
            rs.Close
        End If
    End If

End Sub

Private Sub CONSULTA_VALES()
Dim ncont           As Integer
Dim cnomprov        As String
Dim sw_append       As Boolean
Dim csql            As String
    
    sw_append = False
    
    dxDBGrid1.Dataset.Close
    DELETEREC_N cnomtabla, cnn_temp
    DELETEREC_N cnomtabla, cnn_temp

    dxDBGrid1.Dataset.ADODataset.ConnectionString = cconex_temp
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls

    With dxDBGrid1.Dataset
        ncont = 1
        csql = "Select f5codpro,f4numval,f3canpro,f3valvta,f3valdol,f2codalm from if3vales where f5codpro='" & txtcodigo.Text & "' AND LEFT(F4NUMVAL,1)='I' ORDER BY F4FECVAL DESC"
        If rs.State = adStateOpen Then rs.Close
        rs.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Do While Not rs.EOF
                csql = "SELECT * FROM IF4VALES WHERE F4NUMVAL='" & rs.Fields("F4NUMVAL") & "' AND F2CODALM='" & rs.Fields("F2CODALM") & "' AND F1CODORI='" & wconc_compra & "'"
                If rs2.State = adStateOpen Then rs2.Close
                rs2.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                If Not rs2.EOF Then
                    If sw_append = False Then
                        sw_append = True
                    End If
                    .Append
                    .FieldValues("ITEM") = ncont
                    .FieldValues("F4NUMORD") = rs2.Fields("F4NUMVAL") & ""
                    .FieldValues("F4FECEMI") = Format(rs2.Fields("F4FECVAL"), "DD/MM/YYYY")
                    cnomprov = ""
                    csql = "Select f2nomprov from ef2proveedores where f2newruc='" & rs2.Fields("F2CODPROV") & "'"
                    If rs3.State = adStateOpen Then rs3.Close
                    rs3.Open csql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rs3.EOF Then
                        cnomprov = Trim("" & rs3.Fields("f2nomprov"))
                    End If
                    rs3.Close
                    .FieldValues("F4CODPRV") = cnomprov
                    .FieldValues("F4TIPMON") = Trim("" & rs2.Fields("F4MONEDA"))
                    .FieldValues("F3CANPRO") = Format(Val(rs.Fields("F3CANPRO") & ""), "###,###,##0.00")
                    .FieldValues("F3PREBRUS") = Format(Val(rs.Fields("F3VALVTA") & ""), "###,###,##0.00")
                    .FieldValues("F3PREBRUD") = Format(Val(rs.Fields("F3VALDOL") & ""), "###,###,##0.00")
                    ncont = ncont + 1
                End If
                rs2.Close
                rs.MoveNext
            Loop
            If sw_append = True Then
                .Post
            End If
        Else
            MsgBox "No existen registros para ser procesados.", vbCritical, "Atención"
        End If
    End With

    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = True

End Sub
