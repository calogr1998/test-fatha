VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "ABOX.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "SSTBARS2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ResuCenCosto2 
   Caption         =   "Resumen por Centro de Costo"
   ClientHeight    =   7005
   ClientLeft      =   1500
   ClientTop       =   1920
   ClientWidth     =   10485
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
   Moveable        =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10485
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00808080&
      Caption         =   "Aceptar"
      Height          =   420
      Left            =   135
      TabIndex        =   20
      Top             =   6480
      Width           =   1185
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1140
      Left            =   90
      TabIndex        =   5
      Top             =   945
      Width           =   10275
      _Version        =   65536
      _ExtentX        =   18124
      _ExtentY        =   2011
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
      Begin Threed.SSFrame SSFrame1 
         Height          =   780
         Left            =   6345
         TabIndex        =   14
         Top             =   180
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1376
         _StockProps     =   14
         Caption         =   "Moneda"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optdolares 
            BackColor       =   &H8000000A&
            Caption         =   "Dolares"
            Height          =   285
            Left            =   2025
            TabIndex        =   16
            Top             =   270
            Width           =   915
         End
         Begin VB.OptionButton optsoles 
            BackColor       =   &H8000000A&
            Caption         =   "Soles"
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.TextBox txtcodigocen 
         Height          =   285
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   3
         Top             =   630
         Width           =   870
      End
      Begin VB.TextBox txtcodigoal 
         Height          =   285
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   2
         Top             =   180
         Width           =   870
      End
      Begin Threed.SSPanel pnldescripcioncen 
         Height          =   285
         Left            =   2115
         TabIndex        =   6
         Top             =   630
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483638
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
         Alignment       =   1
      End
      Begin Threed.SSPanel pnldescripcional 
         Height          =   285
         Left            =   2115
         TabIndex        =   7
         Top             =   180
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   503
         _StockProps     =   15
         BackColor       =   -2147483638
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
         Alignment       =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen:"
         Height          =   285
         Left            =   225
         TabIndex        =   13
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lblcentro 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro:"
         Height          =   285
         Left            =   225
         TabIndex        =   9
         Top             =   630
         Width           =   690
      End
      Begin VB.Label lblalmacen 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         Top             =   2790
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00808080&
      Caption         =   "Cancelar"
      Height          =   420
      Left            =   1305
      TabIndex        =   4
      Top             =   6480
      Width           =   1185
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   870
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   10275
      _Version        =   65536
      _ExtentX        =   18124
      _ExtentY        =   1535
      _StockProps     =   14
      Caption         =   "Rango de Fechas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin aBoxCtl.aBox abodesde 
         Height          =   315
         Left            =   3375
         TabIndex        =   0
         Top             =   315
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
         Text            =   "22/01/2003"
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
         ButtonPicture   =   "NuevaConsultaInventa.frx":0000
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
      Begin aBoxCtl.aBox abohasta 
         Height          =   315
         Left            =   6120
         TabIndex        =   1
         Top             =   315
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
         Text            =   "22/01/2003"
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
         ButtonPicture   =   "NuevaConsultaInventa.frx":0352
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   2565
         TabIndex        =   12
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   5310
         TabIndex        =   11
         Top             =   360
         Width           =   510
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   4020
      Left            =   90
      TabIndex        =   17
      Top             =   2250
      Width           =   10275
      _Version        =   65536
      _ExtentX        =   18124
      _ExtentY        =   7091
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
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
         Height          =   3570
         Left            =   450
         OleObjectBlob   =   "NuevaConsultaInventa.frx":06A4
         TabIndex        =   18
         Top             =   4365
         Width           =   9690
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   3705
         Left            =   135
         OleObjectBlob   =   "NuevaConsultaInventa.frx":1FE5
         TabIndex        =   19
         Top             =   180
         Width           =   10005
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   4005
      Top             =   6390
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "NuevaConsultaInventa.frx":3926
      ToolBars        =   "NuevaConsultaInventa.frx":B816
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   3105
      Top             =   6435
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save to"
      FileName        =   "GridNum"
   End
End
Attribute VB_Name = "ResuCenCosto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn_temp        As New ADODB.Connection
Dim cconex_temp     As String

Dim RS As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim RST As ADODB.Recordset
Dim RST1 As ADODB.Recordset
Dim cadena As String
Dim pcosto As String
Dim cnomcen As String
Dim pcodalm As String
Dim cnomalm As String

Dim GridNum As Byte
Dim OldValue As Byte

Private Sub abodesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        abohasta.SetFocus
    End If

End Sub

Private Sub abohasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtcodigoal.SetFocus
    End If

End Sub

Private Sub Form_Load()

    Me.Left = 1500
    Me.Top = 980
    Me.Height = 7890
    Me.Width = 10530
       
    abodesde.Value = Format(Date, "dd/mm/yyyy")
    abohasta.Value = Format(Date, "dd/mm/yyyy")
    
    BASE_TEMPORAL
    TABLA_TEMPORAL
    MEDIDAS
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    cnn_temp.Close

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim n       As Byte
Dim IValue  As Byte
Dim SQL     As String

    Select Case Tool.Id
        Case "ID_Actualizar"
            Me.MousePointer = 11
            Me.Visible = False
            MEDIDAS_2
            LLENADO
            Me.MousePointer = 1
            Me.Visible = True
            With dxDBGrid1
                .DefaultFields = True
                .Dataset.ADODataset.ConnectionString = cconex_temp
                 SQL = "select * from " & DBTable & ""
                .Dataset.Active = False
                .Dataset.ADODataset.CommandText = SQL
                .Dataset.Active = True
                .KeyField = "ITEM"
            End With
            CABECERA
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
        Case "ID_Salir"
            Me.Hide
            Unload Me
    End Select
    
End Sub
          
Private Sub txtcodigoal_Change()
    
    pnldescripcional.Caption = ""
        
End Sub

Private Sub txtcodigoal_DblClick()
    
    txtcodigoal_KeyDown 113, 0
    
End Sub

Private Sub txtcodigoal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        txtcodigoal.Text = ""
        pnldescripcional.Caption = ""
        wtipoayuda = "A"
        Ayudas.Top = 3800
        Ayudas.Left = 6000
        Ayudas.Show 1
        txtcodigoal.Text = wcodigos
        pnldescripcional.Caption = wdescripcion
        Ayudas.Hide
        Unload Ayudas
    End If
    
End Sub

Private Sub txtcodigoal_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtcodigocen.SetFocus
    End If
    
End Sub

Private Sub txtcodigocen_Change()
    
    pnldescripcioncen.Caption = ""
    
End Sub

Private Sub txtcodigocen_DblClick()

    txtcodigocen_KeyDown 113, 0
    
End Sub

Private Sub txtcodigocen_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
         txtcodigocen.Text = ""
         pnldescripcioncen.Caption = ""
         wtipoayuda = "CE"
         Ayudas.Top = 3800
         Ayudas.Left = 6000
         Ayudas.Show 1
         txtcodigocen.Text = wcodigos
         pnldescripcioncen.Caption = wdescripcion
         Ayudas.Hide
         Unload Ayudas
    End If
    
End Sub

Public Sub BASE_TEMPORAL()

    cnombase = wusuario & "Centros" & Format(Time, "hh_mm_ss") & ".MDB"
    CREATEDATABASE_N wrutatemp & "\", cnombase
    cconex_temp = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & cnombase & "; Persist Security Info=False"
    cnn_temp.Open cconex_temp

End Sub

Public Sub TABLA_TEMPORAL()
Dim SQL     As String

    cnomtabla = "Centros"
    SQL = "(ITEM Text(5),CODALMACEN Text(20),NUMVALE Text(15),FECHA DATE,CODPRODUCTO Text(20),NOMPRODUCTO Text(100),CANTIDAD Double,COSTO Double,TOTAL Double)"
    CREATETABLE_N cnomtabla, CStr(SQL), cnn_temp

End Sub

Public Sub LLENADO()
Dim X               As Integer
Dim SQL             As String
Dim FECHA           As Date
Dim codigoalmacen   As String
Dim codigocentro    As String
Dim vale            As String
Dim Moneda          As String
Dim tipocambio      As Double
Dim codigoproducto  As String
Dim cantidad        As Double
Dim costo           As Double
Dim nombreproducto  As String
Dim valorventa      As Double
Dim acumulador      As Double
Dim TOTAL           As Double
Dim csql            As String

    Set RS = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    Set rs4 = New ADODB.Recordset
    Set rs5 = New ADODB.Recordset
    
    dxDBGrid1.Dataset.Close
    DELETEREC_N cnomtabla, cnn_temp
    
    vcodigoalmacen = txtcodigoal.Text
    vcodigocentro = txtcodigocen.Text
    X = 1
    SQL = "Select * from if4vales where f4fecval >= CVDate('" & abodesde.Value & "') and f4fecval <= CVDate('" & abohasta.Value & "') and f4centro='" & vcodigocentro & "'"
    If vcodigoalmacen <> "" Then
        SQL = SQL & " and f2codalm='" & vcodigoalmacen & "'"
    End If
    
    If rs1.State = adStateOpen Then RS.Close
    rs1.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        Do While Not rs1.EOF
            FECHA = "" & Trim(rs1.Fields("f4fecval"))
            codigoalmacen = "" & Trim(rs1.Fields("f2codalm"))
            codigocentro = "" & Trim(rs1.Fields("f4centro"))
            vale = "" & Trim(rs1.Fields("f4numval"))
            Moneda = "" & Trim(rs1.Fields("f4moneda"))
            tipocambio = VAL("" & rs1.Fields("f4tipcam"))
            
            SQL = "Select * from if3vales where f2codalm='" & codigoalmacen & "' and f4numval='" & vale & "'"
            If rs2.State = adStateOpen Then rs2.Close
            rs2.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not rs2.EOF Then
                Do While Not rs2.EOF
                    codigoproducto = "" & Trim(rs2.Fields("f5codpro"))
                    cantidad = VAL("" & rs2.Fields("f3canpro"))
                    costo = VAL("" & rs2.Fields("f3valvta"))
                    
                    SQL = "Select f5nompro from if5pla where f5codpro='" & codigoproducto & "'"
                    If rs3.State = adStateOpen Then rs3.Close
                    rs3.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rs3.EOF Then
                        nombreproducto = "" & Trim(rs3.Fields("f5nompro"))
                    End If
                    rs3.Close
                    
                    'Evaluacion de la moneda y del tipo que se pide
                    
                    If optsoles.Value = True Then
                        If rs1.Fields("f4moneda") = "S" Then
                            valorventa = Format(costo, "0.00")
                        Else
                            valorventa = Format(costo * tipocambio, "0.00")
                        End If
                    Else
                        If optdolares.Value = True Then
                            If rs1.Fields("f4moneda") = "D" Then
                                valorventa = Format(costo, "0.00")
                            Else
                                If tipocambio > 0 Then
                                    valorventa = Format(costo / tipocambio, "0.00")
                                Else
                                    valorventa = Format(0, "0.00")
                            
                                End If
                            End If
                        End If
                    End If
                    'Calculo de totales teniendo en cuenta si el vale empieza en S
                    If Left(vale, 1) = "S" Then
                        cantidad = cantidad * -1
                        valorventa = valorventa * -1
                        acumulador = cantidad * valorventa
                        TOTAL = Format(acumulador * -1, "0.00")
                    Else
                        TOTAL = Format(cantidad * valorventa, "0.00")
                    End If
                    'Sentecia de tabla temporal
                    csql = "INSERT INTO " & DBTable & " (ITEM,CODALMACEN,NUMVALE,FECHA,CODPRODUCTO,NOMPRODUCTO,CANTIDAD,COSTO,TOTAL)" & _
                    "VALUES (" & X & " ,'" & codigoalmacen & "','" & vale & "','" & FECHA & "','" & codigoproducto & "','" & nombreproducto & "'," & cantidad & "," & valorventa & "," & TOTAL & ")"
                    cnn_temp.Execute (csql)
                    X = X + 1
                    rs2.MoveNext
                Loop
            End If
            rs2.Close
            rs1.MoveNext
        Loop
    End If
    
    rs1.Close
    
End Sub

Public Sub CABECERA()

    'dxDBGrid1.Columns.HeaderFontColor = &HC00000
    dxDBGrid1.Columns.HeaderFont.Italic = True
    
    dxDBGrid1.HighlightColor = &HC0FFFF
    dxDBGrid1.HighlightColor = &HC000000
       
    dxDBGrid1.Columns(0).Caption = "Item"
    dxDBGrid1.Columns(1).Caption = "Almacen"
    dxDBGrid1.Columns(2).Caption = "No Vale"
    dxDBGrid1.Columns(3).Caption = "Fecha"
    dxDBGrid1.Columns(4).Caption = "Cod.Producto"
    dxDBGrid1.Columns(5).Caption = "Producto"
    dxDBGrid1.Columns(6).Caption = "Cantidad"
    dxDBGrid1.Columns(7).Caption = "Costo"
    dxDBGrid1.Columns(8).Caption = "Total"
    dxDBGrid1.Columns(7).DecimalPlaces = 2
    dxDBGrid1.Columns(8).DecimalPlaces = 2
    dxDBGrid1.Columns(0).Width = 20
    dxDBGrid1.Columns(1).Width = 50
    dxDBGrid1.Columns(2).Width = 65
    dxDBGrid1.Columns(3).Width = 65
    dxDBGrid1.Columns(4).Width = 80
    dxDBGrid1.Columns(5).Width = 180
    dxDBGrid1.Columns(6).Width = 65
    dxDBGrid1.Columns(7).Width = 65
    dxDBGrid1.Columns(8).Width = 75
    dxDBGrid1.Columns(8).SummaryFooterType = cstSum
    dxDBGrid1.Columns(0).DisableEditor = True
    dxDBGrid1.Columns(1).DisableEditor = True
    dxDBGrid1.Columns(2).DisableEditor = True
    dxDBGrid1.Columns(3).DisableEditor = True
    dxDBGrid1.Columns(4).DisableEditor = True
    dxDBGrid1.Columns(5).DisableEditor = True
    dxDBGrid1.Columns(6).DisableEditor = True
    dxDBGrid1.Columns(7).DisableEditor = True
    dxDBGrid1.Columns(8).DisableEditor = True
    dxDBGrid1.Columns(0).Visible = False

End Sub

Private Sub txtcodigocen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(txtcodigocen.Text) <> "" Then
            cmdaceptar.SetFocus
        Else
            MsgBox "Ingresar un Centro de Costo", vbInformation + vbDefaultButton1, "Atencion"
            txtcodigocen.SetFocus
        End If
    End If
    
End Sub

Public Sub MEDIDAS()

    Me.Height = 3000
    dxDBGrid1.Visible = False
    SSPanel2.Visible = False
    
End Sub

Private Sub txtcodigocen_LostFocus()

    If Trim(txtcodigocen.Text) <> "" Then
        pcosto = txtcodigocen.Text
        If VALIDA_CCOSTO(txtcodigocen.Text) = True Then
            pnldescripcioncen.Caption = cnomcen
        Else
            MsgBox "Codigo de Centro de Costo no existe", vbInformation + vbDefaultButton1, "Atención"
            txtcodigocen.Text = "": txtcodigocen.SetFocus
        End If
    End If
    
End Sub

Private Sub txtcodigoal_LostFocus()

    If Trim(txtcodigoal.Text) <> "" Then
        If VALIDA_ALMACEN(txtcodigoal.Text) = True Then
            pnldescripcional.Caption = cnomalm
        Else
            MsgBox "Codigo de Almacen no existe", vbInformation + vbDefaultButton1, "Atencion"
            txtcodigoal.Text = "": txtcodigoal.SetFocus
        End If
    End If

End Sub

Private Function VALIDA_CCOSTO(pcosto As String)
Set RST = New ADODB.Recordset
Dim sw As Boolean

    sw = False
    If RST.State Then RST.Close
    RST.Open "Select * from centros where f3costo='" & Trim(pcosto) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RST.EOF Then
        cnomcen = Trim(RST!f3descrip & "")
        sw = True
    Else
        sw = False
    End If
    RST.Close
    VALIDA_CCOSTO = sw

End Function

Public Function VALIDA_ALMACEN(pcodialm As String)
Dim sw1 As Boolean

    Set RST1 = New ADODB.Recordset

    sw1 = False
    If RST1.State Then RST1.Close
    RST1.Open "Select * from ef2almacenes where f2codalm='" & Trim(pcodialm) & "'", cnn_dbbancos
    If Not RST1.EOF Then
        cnomalm = Trim(RST1!F2NOMALM & "")
        sw1 = True
    Else
        sw1 = False
    End If
    RST1.Close
    VALIDA_ALMACEN = sw1

End Function

Public Sub MEDIDAS_2()

    Me.Height = 7000
    dxDBGrid1.Visible = True
    SSPanel2.Visible = True
    
End Sub

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
            .DialogTitle = "Resumen por Centro de Costo"
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
