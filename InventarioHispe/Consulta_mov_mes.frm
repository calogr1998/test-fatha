VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Consulta_mov_mes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de movimientos x mes"
   ClientHeight    =   2820
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5265
   Icon            =   "Consulta_mov_mes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1092
      Left            =   84
      TabIndex        =   9
      Top             =   1236
      Width           =   5076
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   876
         TabIndex        =   14
         Top             =   600
         Width           =   576
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   876
         TabIndex        =   11
         Top             =   228
         Width           =   576
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   300
         Left            =   1488
         TabIndex        =   12
         Top             =   240
         Width           =   3468
         _Version        =   65536
         _ExtentX        =   6117
         _ExtentY        =   529
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   300
         Left            =   1500
         TabIndex        =   15
         Top             =   615
         Width           =   3465
         _Version        =   65536
         _ExtentX        =   6117
         _ExtentY        =   529
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
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto"
         Height          =   276
         Left            =   108
         TabIndex        =   13
         Top             =   612
         Width           =   912
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         Height          =   276
         Left            =   84
         TabIndex        =   10
         Top             =   240
         Width           =   912
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   84
      TabIndex        =   5
      Top             =   696
      Width           =   5076
      Begin VB.TextBox TxtCodProd 
         Height          =   288
         Left            =   888
         TabIndex        =   7
         Top             =   168
         Width           =   540
      End
      Begin Threed.SSPanel PnlProd 
         Height          =   300
         Left            =   1488
         TabIndex        =   8
         Top             =   180
         Width           =   3468
         _Version        =   65536
         _ExtentX        =   6117
         _ExtentY        =   529
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
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   252
         Left            =   72
         TabIndex        =   6
         Top             =   204
         Width           =   816
      End
   End
   Begin VB.Frame Frame1 
      Height          =   696
      Left            =   84
      TabIndex        =   0
      Top             =   -12
      Width           =   5076
      Begin VB.ComboBox Combo4 
         Height          =   288
         ItemData        =   "Consulta_mov_mes.frx":000C
         Left            =   4140
         List            =   "Consulta_mov_mes.frx":0019
         TabIndex        =   4
         Text            =   "2007"
         Top             =   288
         Width           =   852
      End
      Begin VB.ComboBox Combo3 
         Height          =   288
         ItemData        =   "Consulta_mov_mes.frx":002F
         Left            =   1596
         List            =   "Consulta_mov_mes.frx":003C
         TabIndex        =   3
         Text            =   "2007"
         Top             =   264
         Width           =   852
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         ItemData        =   "Consulta_mov_mes.frx":0052
         Left            =   2652
         List            =   "Consulta_mov_mes.frx":007A
         TabIndex        =   2
         Text            =   "Enero"
         Top             =   276
         Width           =   1332
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "Consulta_mov_mes.frx":00E2
         Left            =   120
         List            =   "Consulta_mov_mes.frx":010A
         TabIndex        =   1
         Text            =   "Enero"
         Top             =   264
         Width           =   1332
      End
   End
   Begin Threed.SSCommand cmdsalir 
      Height          =   360
      Left            =   3840
      TabIndex        =   16
      Top             =   2385
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdaceptar 
      Height          =   360
      Left            =   2460
      TabIndex        =   17
      Top             =   2400
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Aceptar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
End
Attribute VB_Name = "Consulta_mov_mes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
'SQL = "TRANSFORM Sum(Consulta1.F3CANPRO) AS SumaDeF3CANPRO" & _
'" SELECT Consulta1.F5CODPRO AS Codigo, Consulta1.Producto, Consulta1.F1NOMORI AS " & _
'" Origen, Sum(Consulta1.F3CANPRO) AS Total From (SELECT IIf(Month([IF4VALES].[F4FECVAL])=1,'Enero',IIf(" & _
'" Month([IF4VALES].[F4FECVAL])=2,'Febrero','Marzo')) AS MES, IF4VALES.F1CODORI, IF3VALES.F5CODPRO, " & _
'" IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, SF1ORIGENES.F1NOMORI, IF5PLA.F7CODMED" & _
'" FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF3VALES  " & _
'" INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND " & _
'" (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) as Consulta1 GROUP BY Consulta1.F5CODPRO, Consulta1.Producto, " & _
'" Consulta1.F1NOMORI ORDER BY Consulta1.Producto, Consulta1.F1NOMORI PIVOT Consulta1.MES;"

If rs.State = 1 Then rs.Close
If TxtCodProd.Text = "" Then wcodproducto = ""

rs.Open "TRANSFORM Sum(Consulta1.F3CANPRO) AS SumaDeF3CANPRO" & _
" SELECT Consulta1.F5CODPRO AS Codigo, [F5NOMPRO] & ' - ' & [F7CODMED] AS Producto, Consulta1.F1NOMORI AS Origen, Sum(Consulta1.F3CANPRO) AS Total " & _
" From (SELECT " & _
" IIf(Month([IF4VALES].[F4FECVAL])=1,'Enero'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=2,'Febrero'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=3,'Marzo'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=4,'Abril'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=5,'Mayo'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=6,'Junio'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=7,'Julio'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=8,'Agosto'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=9,'Setiembre'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=10,'Octubre'," & _
" IIf(Month([IF4VALES].[F4FECVAL])=11,'Noviembre'," & _
"'Diciembre'))))))))))) AS MES, " & _
" IF4VALES.F1CODORI, IF3VALES.F5CODPRO, " & _
" IF3VALES.F3CANPRO, IF5PLA.F5NOMPRO, SF1ORIGENES.F1NOMORI, IF5PLA.F7CODMED" & _
" FROM (IF4VALES INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI) INNER JOIN (IF3VALES  " & _
" INNER JOIN IF5PLA ON IF3VALES.F5CODPRO = IF5PLA.F5CODPRO) ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND " & _
" (IF4VALES.F2CODALM = IF3VALES.F2CODALM) Where SF1ORIGENES.CONSUMO = '*') as Consulta1 Where IF3VALES.F5CODPRO Like '" & IIf(wcodproducto = "", "%", wcodproducto) & "' GROUP BY Consulta1.F5CODPRO, [F5NOMPRO] & ' - ' & [F7CODMED], Consulta1.F1NOMORI " & _
" PIVOT Consulta1.MES", cnn_dbbancos, adOpenDynamic, adLockReadOnly
'MsgBox cnn_form

cnombase = "TEMPLUS.MDB"
cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
If cnn_form.State <> 1 Then cnn_form.Open cconex_form
sql = "Delete From TempConVeMes"
cnn_form.Execute sql
Dim datos As String, enero As Double, febrero As Double, marzo As Double, abril As Double, mayo As Double, junio As Double, julio As Double, agosto As Double
Dim setiembre As Double, octubre As Double, noviembre As Double, diciembre As Double
Do While rs.EOF = False
    datos = "'" & rs("Codigo") & "', '" & rs("Producto") & "', '" & rs("Origen") & "'"
    For I = 0 To rs.Fields.Count - 1
        If rs(I).Name = "Enero" Then
            enero = IIf(IsNull(rs("Enero")), 0, rs("Enero"))
        End If
        If rs(I).Name = "Febrero" Then
            febrero = IIf(IsNull(rs("Febrero")), 0, rs("Febrero"))
        End If
        If rs(I).Name = "Marzo" Then
            marzo = IIf(IsNull(rs("Marzo")), 0, rs("Marzo"))
        End If
        If rs(I).Name = "Abril" Then
            abril = IIf(IsNull(rs("Abril")), 0, rs("Abril"))
        End If
        If rs(I).Name = "Mayo" Then
            mayo = IIf(IsNull(rs("Mayo")), 0, rs("Mayo"))
        End If
        If rs(I).Name = "Junio" Then
            junio = IIf(IsNull(rs("Junio")), 0, rs("Junio"))
        End If
        If rs(I).Name = "Julio" Then
            julio = IIf(IsNull(rs("Julio")), 0, rs("Julio"))
        End If
        If rs(I).Name = "Agosto" Then
            agosto = IIf(IsNull(rs("Agosto")), 0, rs("Agosto"))
        End If
        If rs(I).Name = "Setiembre" Then
            setiembre = IIf(IsNull(rs("Setiembre")), 0, rs("Setiembre"))
        End If
        If rs(I).Name = "Octubre" Then
            octubre = IIf(IsNull(rs("Octubre")), 0, rs("Octubre"))
        End If
        If rs(I).Name = "Noviembre" Then
            noviembre = IIf(IsNull(rs("Noviembre")), 0, rs("Noviembre"))
        End If
        If rs(I).Name = "Diciembre" Then
            diciembre = IIf(IsNull(rs("Diciembre")), 0, rs("Diciembre"))
        End If
        
    Next I
    
    datos = datos & "," & enero & "," & febrero & ", " & marzo & ", " & abril & ", " & mayo & "," & junio & ", " & julio & ", " & agosto & ", " & setiembre & "," & octubre & ", " & noviembre & ", " & diciembre & ", " & Val(enero) + Val(febrero) + Val(marzo) + Val(abril) + Val(mayo) + Val(junio) + Val(julio) + Val(agosto) + Val(setiembre) + Val(octubre) + Val(noviembre) + Val(diciembre)
     
    sql = "Insert Into Tempconvemes Values(" & datos & ")"
    enero = 0
    febrero = 0
    marzo = 0
    abril = 0
    mayo = 0
    junio = 0
    julio = 0
    agosto = 0
    setiembre = 0
    octubre = 0
    noviembre = 0
    diciembre = 0
       cnn_form.Execute sql
    rs.MoveNext
Loop

acr_Con_venta_mes.datos.ConnectionString = cnn_form

acr_Con_venta_mes.datos.Source = "Select Codigo,Producto,ORIGEN,sum(ENERO) as Enero,sum(FEBRERO) as Febrero,Sum(MARZO) as Marzo,Sum(ABRIL) as Abril,sum(MAYO) as Mayo,sum(JUNIO) as Junio,Sum(JULIO) as Julio,Sum(AGOSTO) as Agosto,sum(SETIEMBRE) as Setiembre,sum(OCTUBRE) as Octubre,Sum(NOVIEMBRE) as Noviembre,Sum(DICIEMBRE) as Diciembre, sum(TOTAL) as Total From TEMPCONVEMES Group By  Codigo,Producto,ORIGEN"
acr_Con_venta_mes.Show 1
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
For I = 1 To 12
    Combo1.AddItem MonthName(Month(Date), False)
    Combo1.ItemData(Combo1.NewIndex) = I
    Combo2.AddItem MonthName(Month(Date), False)
    Combo2.ItemData(Combo2.NewIndex) = I
Next I
Combo1.ListIndex = 0
Combo2.ListIndex = Month(Date) - 1
End Sub
Private Sub TxtCodProd_DblClick()
Call TxtCodProd_KeyDown(vbKeyF2, 1)
End Sub

Private Sub TxtCodProd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    wcod_alm = TxtCodProd.Text
    wcodproducto = ""
    sw_ayuda_prod = True
    wtipoguia = "S"
    ayuda_productos.Show 1
    If wcodproducto = "" Then
        TxtCodProd.Text = ""
        PnlProd.Caption = "TODOS LOS PRODUCTOS"
    Else
        TxtCodProd.Text = wcodproducto
        PnlProd.Caption = wdesproducto
    End If
    
End If
End Sub

Private Sub TxtCodProd_LostFocus()
If TxtCodProd.Text = "" Then
    PnlProd.Caption = "TODOS LOS PRODUCTOS"
End If
End Sub
