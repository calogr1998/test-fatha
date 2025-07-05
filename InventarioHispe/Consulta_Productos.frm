VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Consulta_Productos 
   Caption         =   "Consulta de precios"
   ClientHeight    =   7830
   ClientLeft      =   1800
   ClientTop       =   1950
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11085
   Begin Threed.SSFrame SSFrame1 
      Height          =   870
      Left            =   90
      TabIndex        =   1
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
         TabIndex        =   3
         Top             =   315
         Width           =   1380
      End
      Begin Threed.SSPanel pnlnompro 
         Height          =   300
         Left            =   2610
         TabIndex        =   2
         Top             =   315
         Width           =   6240
         _Version        =   65536
         _ExtentX        =   11007
         _ExtentY        =   529
         _StockProps     =   15
         BackColor       =   12632256
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
         Autosize        =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   360
         Width           =   645
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   6540
      Left            =   135
      OleObjectBlob   =   "Consulta_Productos.frx":0000
      TabIndex        =   0
      Top             =   1080
      Width           =   10800
   End
End
Attribute VB_Name = "Consulta_Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim conexion As ADODB.Connection
Dim conexion_1 As ADODB.Connection
Dim conexion_2 As ADODB.Connection
Dim temp As ADODB.Connection
Dim rs As ADODB.Recordset
Dim a As String
Dim b As String
Dim c As String
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Private Sub Form_Load()
Me.Height = 7890
Me.Width = 10530
Me.Left = 960
Me.Top = 700


Set conexion = New ADODB.Connection
Set conexion_1 = New ADODB.Connection
Set conexion_2 = New ADODB.Connection
Set temp = New ADODB.Connection


With conexion
    .Provider = "Microsoft.Jet.OLEDB.4.0; " & _
    "Data Source=C:\bancowin\Agro01\Inventa.mdb ; " & _
    "Persist Security Info=False"
    .Open
End With

With conexion_1
    .Provider = "Microsoft.Jet.OLEDB.4.0; " & _
    "Data Source=C:\bancowin\Agro01\empresa.mdb; " & _
    "Persist Security Info=False"
    .Open
End With

With conexion_2
    .Provider = "Microsoft.Jet.OLEDB.4.0; " & _
    "Data Source=C:\bancowin\Agro01\compras.mdb; " & _
    "Persist Security Info=False"
    .Open
End With

'Creando la base de datos temporal
usuario = "Pamela"
base_temporal = usuario & "Producto" & Format(Time, "hh_mm_ss") & ".MDB"
CREATEDATABASE_N "C:\", CStr(base_temporal)

CON = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=C:\" & base_temporal & ";Persist Security Info=False"
temp.Open CON

'Creando la tabla temporal
DBTable2 = "Productos"
SQL = "(ITEM Text(4),F4NUMORD Text(4),F4FECEMI Date,F4CODPRV Text(20),F4TIPMON Text(3),F3CANPRO Text(5),F3PREBRUS Double,F3PREBRUD Double)"

CREATETABLE_N DBTable2, CStr(SQL), temp


End Sub

Public Sub Llenado_Registros()

Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
    
    dxDBGrid1.Dataset.Close
    DELETEREC_N DBTable2, temp
    DELETEREC_N DBTable2, temp

    dxDBGrid1.Dataset.ADODataset.ConnectionString = temp
    dxDBGrid1.Dataset.Active = True

    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Dataset.DisableControls

    With dxDBGrid1.Dataset
        X = 1
        
        SQL = "Select f3codpro,f4numord,f3canpro,f3prebru from if3orden where f3codpro='" & a & "'"
        If rs.State = adStateOpen Then rs.Close
        rs.Open SQL, conexion_2, adOpenDynamic, adLockOptimistic
                
        If Not rs.EOF Then
            Do While Not rs.EOF
            
                SQL = "Select * from if4orden where f4numord=" & rs.Fields("f4numord") & ""
                If rs2.State = adStateOpen Then rs2.Close
                rs2.Open SQL, conexion_2, adOpenDynamic, adLockOptimistic
                If Not rs2.EOF Then
                
                    .Append
                    .FieldValues("ITEM") = X
                    .FieldValues("F4NUMORD") = "" & Trim(rs2.Fields("F4NUMORD"))
                    .FieldValues("F4FECEMI") = "" & Trim(rs2.Fields("F4FECEMI"))
                    
                    R = rs2.Fields("F4CODPRV")
                    SQL = "Select f2newruc,f2nomprov from ef2proveedores where f2newruc='" & R & "'"
                    If rs3.State = adStateOpen Then rs3.Close
                    rs3.Open SQL, conexion_1, adOpenDynamic, adLockOptimistic
                    If Not rs3.EOF Then
                        n = "" & Trim(rs3.Fields("f2nomprov"))
                    End If
                    rs3.Close
                    .FieldValues("F4CODPRV") = n
                    .FieldValues("F4TIPMON") = "" & Trim(rs2.Fields("F4TIPMON"))
                    
                    'Almacenando un valor que guarde el tipo de moneda
                    Y = "" & Trim(rs2.Fields("F4TIPMON"))
                    
                    
                    'Llenados de Cantidad
                    .FieldValues("F3CANPRO") = "" & Trim(rs.Fields("F3CANPRO"))
                    
                    'Llenado de Soles y Dolares
                    If Y = "S" Then
                        '.FieldValues("F3PREBRUS") = CStr(Trim(rs.Fields("F3PREBRU"))) * CStr(Trim(rs.Fields("F3CANPRO")))
                        .FieldValues("F3PREBRUS") = "" & Format(Val("" & rs.Fields("F3PREBRU")) * Val("" & rs.Fields("F3CANPRO")), "#0.00")
                        .FieldValues("F3PREBRUD") = "" & Format(Val("" & rs.Fields("F3PREBRU")) * Val("" & rs.Fields("F3CANPRO")) / Val("" & rs2.Fields("F4TIPCAM")), "#0.00")
                        
                       ' .FieldValues("F3PREBRUD") = CStr(Trim(rs.Fields("F3PREBRU"))) * CStr(Trim(rs.Fields("F3CANPRO"))) / CStr(Trim(rs2.Fields("F4TIPCAM")))
                    Else
                        If Y = "D" Then
                            .FieldValues("F3PREBRUS") = "" & Format(Val("" & rs.Fields("F3PREBRU")) * Val("" & rs.Fields("F3CANPRO")), "#0.00")
                            .FieldValues("F3PREBRUD") = "" & Format(Val("" & rs.Fields("F3PREBRU")) * Val("" & rs.Fields("F3CANPRO")) / Val("" & rs2.Fields("F4TIPCAM")), "#0.00")
                           ' .FieldValues("F3PREBRUS") = CStr(Trim(rs.Fields("F3PREBRU"))) * CStr(Trim(rs.Fields("F3CANPRO")))
                           ' .FieldValues("F3PREBRUD") = CStr(Trim(rs.Fields("F3PREBRU"))) * CStr(Trim(rs.Fields("F3CANPRO"))) / CStr(Trim(rs2.Fields("F4TIPCAM")))
                        
                        End If
                    End If
                
                
                    X = X + 1
                    
                End If
                rs2.Close
                
                rs.MoveNext
    
            Loop
        .Post
        End If
        '.Post
    End With

    dxDBGrid1.Dataset.EnableControls
    dxDBGrid1.Dataset.Close
    dxDBGrid1.Dataset.Open
    Rem NSE dxDBGrid1.Dataset.Refresh
    dxDBGrid1.OptionEnabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Unload Ayu_Prod
    Unload Ayud_Prod
End Sub

Private Sub TxtCodigo_DblClick()

TxtCodigo_KeyDown 113, 0

End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 113 Then
        wcod_alm = ""
        hlp_productos.Show 1
        If Len(Trim(wcodproducto)) > 0 Then
            TxtCodigo.Text = wcodproducto
            pnlnompro.Caption = wdesproducto
            TxtCodigo_KeyPress 13
        End If
    End If

End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)

Set rs = New ADODB.Recordset


    If KeyAscii = 13 Then
        If Trim(TxtCodigo.Text) <> "" Then
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "Select f5codpro,f5nompro from if5pla where f5codpro='" & TxtCodigo.Text & "'", conexion, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Len(rs.Fields(0)) > 0 Then
                    a = Trim$(rs.Fields("f5codpro"))
                    pnlnompro.Caption = rs.Fields("f5nompro")
                                        
                    Llenado_Registros
                    
                End If
            Else
                MsgBox "No encontro", vbExclamation, "Aviso"
            End If
        
        End If
    End If

End Sub

