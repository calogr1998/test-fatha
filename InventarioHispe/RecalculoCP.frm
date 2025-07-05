VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.Form RecalculoCP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costo Promedio"
   ClientHeight    =   1560
   ClientLeft      =   180
   ClientTop       =   1755
   ClientWidth     =   2955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   2955
   Begin VB.CommandButton Command1 
      Caption         =   "Recalcular Costo Pomedio"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _Version        =   65536
      _cx             =   3625
      _cy             =   450
      ForeColor       =   0
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinPos          =   0
      MaxPos          =   100
      Pos             =   0
      Step            =   10
      ShowText        =   -1  'True
      Orientation     =   0
      StartColor      =   16711680
      EndColor        =   16777215
      DrawBorderStyle =   1
      ShowTextStyle   =   1
      DrawBarStyle    =   3
      DrawBarBorderStyle=   2
   End
End
Attribute VB_Name = "RecalculoCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim MovVales As ADODB.Recordset
    Dim amovs1(0)  As a_grabacion
    Dim valor As Double
    Dim Cantidad As Integer
    Dim SQLm As String
    Dim CSQL1 As String
    Dim Csql2 As String
    
    dxProgressBar1.Pos = 0
    
    Me.MousePointer = vbHourglass
    
    sql = vbNullString
    sql = sql & "SELECT "
    sql = sql & "a.F2CODALM, "
    sql = sql & "a.F4NUMVAL, "
    sql = sql & "a.F5CODPRO, "
    sql = sql & "a.F4FECVAL, "
    sql = sql & "b.F4MONEDA "
    sql = sql & "FROM "
    sql = sql & "if3vales AS a "
    sql = sql & "INNER JOIN if4vales AS b "
    sql = sql & "ON (a.F4NUMVAL = b.F4NUMVAL) AND (a.F2CODALM = b.F2CODALM) "
    sql = sql & "WHERE "
    sql = sql & "(((Left([a].[F4NUMVAL],1))='S')) "
    sql = sql & "ORDER BY "
    sql = sql & "a.F2CODALM, a.F4FECVAL, a.F4NUMVAL"
    
    Set MovVales = New ADODB.Recordset
    
    MovVales.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        
    If Not MovVales.EOF Then
        MovVales.MoveFirst
        valor = 0
        Do While Not MovVales.EOF
            valor = valor + 1
            MovVales.MoveNext
        Loop
    End If
    
    dxProgressBar1.Pos = 0
    dxProgressBar1.Step = 100 / valor
        
    If Not MovVales.EOF Then
        MovVales.MoveFirst
        
        Do While Not MovVales.EOF
            valor = Costo_Calculado(MovVales.Fields("f5codpro"), MovVales.Fields("f4fecval"), MovVales.Fields("f4moneda"), MovVales.Fields("f4numval"))
            
            CSQL1 = "f5codpro = '" & MovVales.Fields("f5codpro") & "' and f4fecval = " & _
                    "cvdate('" & Format(MovVales.Fields("f4fecval"), "dd/mm/yyyy") & "') and f4numval = '" & MovVales.Fields("f4numval") & "';"
                
            amovs1(0).campo = "F3VALVTA": amovs1(0).valor = valor: amovs1(0).Tipo = "T"
            
            GRABA_REGISTRO_logistica amovs1(), "if3vales", "M", 0, cnn_dbbancos, CSQL1
            
            dxProgressBar1.DoStep
            
            MovVales.MoveNext
        Loop
    End If
    
    MovVales.Close
'    sql = "UPDATE IF6ALMA SET F6STOCKACT = 0 "
'    cnn_dbbancos.Execute sql
    'AlmacenaQuery_sql sql, cnn_dbbancos
    
    sql = vbNullString
    sql = sql & "select "
    sql = sql & "f2codalm, "
    sql = sql & "f5codpro, "
    sql = sql & "LEFT(F4NUMVAL,1) as tipo, "
    sql = sql & "sum(f3canpro) as cantidad "
    sql = sql & "from "
    sql = sql & "if3vales "
    sql = sql & "group by "
    sql = sql & "f5codpro, LEFT(F4NUMVAL,1),f2codalm"
    
    MovVales.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        
    If Not MovVales.EOF Then
        MovVales.MoveFirst
        
        Do While Not MovVales.EOF
            If MovVales.Fields("tipo") = "I" Then
                CSQL1 = "UPDATE IF5PLA SET F5STOCKACT = " & MovVales.Fields("cantidad") & " WHERE F5CODPRO = '" & MovVales.Fields("f5codpro") & "'"
                cnn_dbbancos.Execute CSQL1
                'AlmacenaQuery_sql CSQL1, cnn_dbbancos
                
                Csql2 = "UPDATE IF6ALMA SET F6STOCKACT = F6STOCKACT + " & MovVales.Fields("cantidad") & " WHERE F2CODALM = '" & MovVales.Fields("f2codalm") & "'AND F5CODPRO = '" & MovVales.Fields("f5codpro") & "'"
                cnn_dbbancos.Execute Csql2
                'AlmacenaQuery_sql Csql2, cnn_dbbancos
            Else
                CSQL1 = "UPDATE IF5PLA SET F5STOCKACT = F5STOCKACT - " & MovVales.Fields("cantidad") & " WHERE F5CODPRO = '" & MovVales.Fields("f5codpro") & "'"
                cnn_dbbancos.Execute CSQL1
                'AlmacenaQuery_sql CSQL1, cnn_dbbancos
                
                Csql2 = "UPDATE IF6ALMA SET F6STOCKACT = F6STOCKACT - " & MovVales.Fields("cantidad") & " WHERE F2CODALM = '" & MovVales.Fields("f2codalm") & "'AND F5CODPRO = '" & MovVales.Fields("f5codpro") & "'"
                cnn_dbbancos.Execute Csql2
                'AlmacenaQuery_sql Csql2, cnn_dbbancos
            End If
        
            MovVales.MoveNext
        Loop
    End If
    
    MovVales.Close
    
    Me.MousePointer = vbDefault
End Sub

Function Costo_Calculado(pcodigo As String, pfecha As Date, pmoneda As String, pnumvale As String) As Double
    Dim CosUni  As ADODB.Recordset
    Dim sql     As String
    
    sql = ""
    
    If pmoneda = "S" Then
        sql = sql & "SELECT "
        sql = sql & "IF3VALES.F5CODPRO, "
        sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, "
        sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
        sql = sql & "[VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
    Else
        sql = sql & "SELECT "
        sql = sql & "IF3VALES.F5CODPRO, "
        sql = sql & "Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I', IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, "
        sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I', IF3VALES.F3VALDOL*IF3VALES.F3CANPRO, (IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
        sql = sql & "[VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
    End If
    
    sql = sql & "FROM "
    sql = sql & "IF4VALES "
    sql = sql & "INNER JOIN IF3VALES ON (IF4VALES.F2CODALM = IF3VALES.F2CODALM) AND (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) "
    sql = sql & "Where "
    sql = sql & "IF3VALES.F4FECVAL <= CVDATE('" & Format(pfecha, "DD/MM/YYYY") & "') And "
    sql = sql & "IF3VALES.F5CODPRO = '" & pcodigo & "' and "
    sql = sql & "not(IF3VALES.F4FECVAL = CVDATE('" & Format(pfecha, "DD/MM/YYYY") & "') and IF3VALES.F4NUMVAL >= '" & pnumvale & "')"
    sql = sql & "GROUP BY "
    sql = sql & "IF3VALES.F5CODPRO"
    
    Set CosUni = New ADODB.Recordset
    
    CosUni.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    
    If Not CosUni.EOF Then
       Costo_Calculado = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
    End If
    
    CosUni.Close
End Function

Private Sub Form_Load()
    dxProgressBar1.Pos = 0
End Sub
