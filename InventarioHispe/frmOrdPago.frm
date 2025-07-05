VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrdPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorización de Pagos"
   ClientHeight    =   4965
   ClientLeft      =   5670
   ClientTop       =   1920
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4560
   Begin VB.Frame FraConfirmar 
      Height          =   1335
      Left            =   60
      TabIndex        =   24
      Top             =   6360
      Width           =   4455
      Begin VB.CommandButton CmdNO 
         Caption         =   "No"
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CmdSI 
         Caption         =   "Si"
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "¿Desea autorizar el pago correspondiente?"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.TextBox txtmonedas3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2100
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtmonedas1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2100
      TabIndex        =   16
      Top             =   720
      Width           =   375
   End
   Begin VB.Frame TxtFecMov 
      Height          =   2535
      Left            =   60
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox txtmonedas5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txtmonedas4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   1080
         Width           =   375
      End
      Begin VB.ComboBox Cmbmone 
         Height          =   315
         ItemData        =   "frmOrdPago.frx":0000
         Left            =   2040
         List            =   "frmOrdPago.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   660
         Width           =   1416
      End
      Begin VB.TextBox txtnuevaau 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtnuevosaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox txtobs 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker txt_fecha 
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   180682753
         CurrentDate     =   40611
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda "
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   19
         Left            =   180
         TabIndex        =   20
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Observación"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Saldo:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Autorización:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Width           =   1440
      End
   End
   Begin VB.TextBox txtsaldopendiente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2460
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtmonedas2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   1140
         Width           =   375
      End
      Begin VB.TextBox txtaufechas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtoc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Pendiente:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Autorizado a la Fecha:"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1140
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Compra:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1290
      End
   End
   Begin ActiveToolBars.SSActiveToolBars atbmenu 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   9
      Tools           =   "frmOrdPago.frx":001E
      ToolBars        =   "frmOrdPago.frx":71DF
   End
End
Attribute VB_Name = "frmOrdPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Af As New ADOFunctions
Dim Estado As String
Dim SaldoPendiente As Double
Dim MONTO As String
Dim Monedas As String
Dim amovs_cab(0 To 10)  As a_grabacion
Dim Est_Aut As Integer
Dim IDOP As String
Dim CodigoAut As Integer
Dim Codigo As String
Dim RsP As New ADODB.Recordset

Private Sub atbmenu_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim resp    As Integer
    
    Select Case Tool.ID
        Case "ID_Nuevo":

        Case "ID_Grabar":
                If sw_ordendepago = True Then
                    csql = "Select IDOP from IF4ORDEN_PAGO Where ORDEN = '" & txtoc.Text & "' And Estado = '0' "
                    Set RsP = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
                    If Not RsP.EOF Then
                        cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set Estado = '1', Est_Aut = " & Est_Aut & ", F4ESTADO=2  Where IDOP = '" & RsP.Fields("IDOP").Value & "' ")
                        lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Value = True
                        MsgBox "Autorizacion de Pago realizada"
                    Else
                        MsgBox "No puede grabar esta autorización de pago porque no existe una Autorización de Pago realizada para esta Orden de Compra"
                        
                    End If
                Else
                    
                    GrabarOrdPago
                    MsgBox "Autorizacion de Pago realizada"
                End If
                Unload Me
        Case "ID_Extornar"
            lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Value = False
            lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB1").Value = False
            cnn_dbbancos.Execute ("Delete From IF4ORDEN_PAGO Where ORDEN = '" & txtoc.Text & "'")
            txtaufechas.Text = "0.00"
            txtsaldopendiente.Text = "0.00"
            MsgBox "Autorización de Pago Extornado"
            Unload Me

        Case "ID_Imprimir":
        
        Case "ID_Eliminar"
        
        Case "ID_Aprobacion"
        
        Case "ID_Salir":
            If MsgBox("¿Desea salir sin grabar la autorización de pago?", 4 + 32, "ATENCIÓN") = vbYes Then
                Unload Me
            End If
    End Select

End Sub

Private Sub Cmbmone_Click()
    If Cmbmone.Text = "Soles" Then
            txtmonedas4.Text = "S/"
            txtmonedas5.Text = "S/"
            Monedas = "S"
        Else
            txtmonedas4.Text = "$"
            txtmonedas5.Text = "$"
            Monedas = "D"
    End If
End Sub

Private Sub Cmbmone_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 13 Then
        If Cmbmone.Text = "Soles" Then
            txtmonedas4.Text = "S/"
            txtmonedas5.Text = "S/"
            Monedas = "S"
        Else
            txtmonedas4.Text = "$"
            txtmonedas5.Text = "$"
            Monedas = "D"
        End If
        txtnuevaau.SetFocus
    End If
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdNO_Click()
        lista_oc.dxDBGrid1.Dataset.Edit
        lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Value = False
        lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB1").Value = False
        lista_oc.dxDBGrid1.Dataset.Post
        cnn_dbbancos.Execute ("Delete From IF4ORDEN_PAGO Where ORDEN = '" & txtoc.Text & "'")
        txtaufechas.Text = "0.00"
        txtsaldopendiente.Text = "0.00"
        'Unload Me
        FraConfirmar.Visible = False
End Sub

Private Sub CmdSI_Click()

        GrabarOrdPago
        MsgBox "Autorizacion de Pago realizada"
        FraConfirmar.Visible = False
        'Unload Me
End Sub

Private Sub Form_Activate()
If sw_e_ordenpago = True Then
    atbmenu.Tools.ITEM("ID_Extornar").Visible = False
    atbmenu.Tools.ITEM("ID_Eliminar").Visible = True
Else
    If sw_ordendepago = False Then
        txtnuevaau.SetFocus
        atbmenu.Tools.ITEM("ID_Eliminar").Visible = False
        atbmenu.Tools.ITEM("ID_Extornar").Visible = False
    Else
        atbmenu.Tools.ITEM("ID_Eliminar").Visible = False
        atbmenu.Tools.ITEM("ID_Extornar").Visible = True
        txtobs.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
Dim NuevaAu As Double
Dim Moneda As String
Dim aufechas As Double
txtmonedas4.Enabled = True
txtnuevaau.Enabled = True
ctipo = "A"
txtoc.Enabled = False
txttotal.Enabled = False
txtmonedas5.Enabled = False
txtnuevosaldo.Enabled = False
txtaufechas.Enabled = False
txtmonedas1.Enabled = False
txtmonedas2.Enabled = False
txtmonedas3.Enabled = False
txtsaldopendiente.Enabled = False
FraConfirmar.Visible = True
'Me.Width = 4650
'Me.Height = 7325
If sw_est_orden_pago = True Then
    Estado = "1"
    atbmenu.Tools.ITEM("ID_Grabar").Visible = True
    txtoc.Text = lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value
    
    MONTO = traerCampo("IF4ORDEN", "F4MONTO", "F4NUMORD", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
    txttotal.Text = Format(MONTO, "###,###,##0.00")
    
    Moneda = traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
    txtmonedas1.Text = IIf(Moneda = "D", "$", "S/")
    txtmonedas2.Text = IIf(Moneda = "D", "$", "S/")
    
    aufechas = Val(traerCampo("IF4ORDEN_PAGO", "Iif(Sum(IMPORTE) IS NULL, 0, Sum(Importe))", "ORDEN", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value, " And Estado = '1'") + 0)
    If aufechas > MONTO Then
        MsgBox "El pago a autorizar es mayor que el monto de la Orden de Compra, autorice un pago menor", vbInformation, "ATENCIÓN"
        Exit Sub
    Else
        txtaufechas.Text = Format(aufechas, "###,###,##0.00") 'rsp
    End If
    
    SaldoPendiente = Format(txttotal.Text - Format(txtaufechas.Text, "###,###,##0.00"), "###,###,##0.00")
    txtsaldopendiente.Text = Format(SaldoPendiente, "###,###,##0.00")
    txtnuevaau.Text = txtsaldopendiente.Text
    txtnuevosaldo.Text = "0.00"
    
    txt_fecha.Value = Format(Date, "dd/MM/yyyy")
    
    txtmonedas3.Text = IIf(Moneda = "D", "$", "S/")
    
    If txtmonedas1.Text = "S/" Then
        Cmbmone.Text = "Soles"
    Else
        Cmbmone.Text = "Dólares"
    End If
    GeneraAut
    GeneraCod
    
    sw_est_ordendepago = False
Else
    If sw_ordendepago = True Then
        txtmonedas4.Enabled = False
        txtnuevaau.Enabled = False
        Estado = "1"
        atbmenu.Tools.ITEM("ID_Grabar").Visible = True
        txtoc.Text = lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value
        
        MONTO = lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4MONTO").Value
        txttotal.Text = Format(MONTO, "###,###,##0.00")
        
        Moneda = traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
        txtmonedas1.Text = IIf(Moneda = "D", "$", "S/")
        txtmonedas2.Text = IIf(Moneda = "D", "$", "S/")
        
        'If rsp.State = 1 Then rsp.Close
        'csql = "Select Sum(Importe) as Importe from IF4ORDEN_PAGO where "
        aufechas = Val(traerCampo("IF4ORDEN_PAGO", "Iif(Sum(IMPORTE) IS NULL, 0, Sum(Importe))", "ORDEN", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value, " And Estado = '1'") + 0)
        If aufechas > MONTO Then
            MsgBox "El pago a autorizar es mayor que el monto de la Orden de Compra, autorice un pago menor", vbInformation, "ATENCIÓN"
            Exit Sub
        Else
            txtaufechas.Text = Format(aufechas, "###,###,##0.00") 'rsp
        End If
        NuevaAu = traerCampo("IF4ORDEN_PAGO", "Iif(Sum(IMPORTE) IS NULL, 0, Sum(Importe))", "ORDEN", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value, " And Estado = '0'")
        'NuevaAu = Format(txttotal.Text - Format(txtaufechas.Text, "###,###,##0.00"), "###,###,##0.00")
        SaldoPendiente = Format(txttotal.Text - Format(txtaufechas.Text, "###,###,##0.00"), "###,###,##0.00")
        txtsaldopendiente.Text = Format(SaldoPendiente, "###,###,##0.00")
        
        
        
        txt_fecha.Value = Format(Date, "dd/MM/yyyy")
        
        txtmonedas3.Text = IIf(Moneda = "D", "$", "S/")
    '    txtnuevaau.Text = SaldoPendiente
    '    txtnuevosaldo = Format(NuevaAu, "###,###,##0.00")
        txtnuevaau.Text = Format(NuevaAu, "###,###,##0.00")
        txtnuevosaldo.Text = Format(txtsaldopendiente.Text - txtnuevaau.Text, "###,###,##0.00") 'rsp
        'txtnuevosaldo = "0.00"
        If txtmonedas1.Text = "S/" Then
            Cmbmone.Text = "Soles"
        Else
            Cmbmone.Text = "Dólares"
        End If
    '    Me.Width = 4650
    '    'Me.Height = 2899
    '    Me.Height = 7325
        GeneraAut
        GeneraCod
    '    FraConfirmar.Visible = True
    '    If MsgBox("Desea Autorizar el Pago correspondiente", 4 + 32, "ATENCIÓN") = vbYes Then
    '        GrabarOrdPago
    '        MsgBox "Pago Autorizado satisfactoriamente"
    '        'Unload Me
    '    Else
    '        lista_oc.dxDBGrid1.Dataset.Edit
    '        lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB2").Value = False
    '        lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4VB1").Value = False
    '        lista_oc.dxDBGrid1.Dataset.Post
    '        cnn_dbbancos.Execute ("Delete From IF4ORDEN_PAGO Where ORDEN = '" & txtoc.Text & "'")
    '        txtaufechas.Text = "0.00"
    '        txtsaldopendiente.Text = "0.00"
    '        'Unload Me
    '    End If
        
    Else
        Estado = "0"
        atbmenu.Tools.ITEM("ID_Grabar").Visible = True
        txtoc.Text = lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value
        
        MONTO = traerCampo("IF4ORDEN", "F4MONTO", "F4NUMORD", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
        txttotal.Text = Format(MONTO, "###,###,##0.00")
        
        Moneda = traerCampo("IF4ORDEN", "F4TIPMON", "F4NUMORD", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value)
        txtmonedas1.Text = IIf(Moneda = "D", "$", "S/")
        txtmonedas2.Text = IIf(Moneda = "D", "$", "S/")
        
        'If rsp.State = 1 Then rsp.Close
        'csql = "Select Sum(Importe) as Importe from IF4ORDEN_PAGO where "
        aufechas = Val(traerCampo("IF4ORDEN_PAGO", "Iif(Sum(IMPORTE) IS NULL, 0, Sum(Importe))", "ORDEN", lista_oc.dxDBGrid1.Columns.ColumnByFieldName("F4NUMORD").Value) + 0)
        If aufechas > MONTO Then
            MsgBox "El pago a autorizar es mayor que el monto de la Orden de Compra, autorice un pago menor", vbInformation, "ATENCIÓN"
            Exit Sub
        Else
            txtaufechas.Text = Format(aufechas, "###,###,##0.00") 'rsp
        End If
        
        SaldoPendiente = Format(txttotal.Text - Format(txtaufechas.Text, "###,###,##0.00"), "###,###,##0.00")
        txtsaldopendiente.Text = Format(SaldoPendiente, "###,###,##0.00")
        txtnuevaau.Text = txtsaldopendiente.Text
        txtnuevosaldo.Text = "0.00"
        
        txt_fecha.Value = Format(Date, "dd/MM/yyyy")
        
        txtmonedas3.Text = IIf(Moneda = "D", "$", "S/")
        
        If txtmonedas1.Text = "S/" Then
            Cmbmone.Text = "Soles"
        Else
            Cmbmone.Text = "Dólares"
        End If
        
        GeneraCod
    End If
    'sw_ordendepago = False
End If

End Sub

Private Sub txtnuevaau_GotFocus()
If sw_ordendepago = False Then
    txtnuevaau.SelStart = 0: txtnuevaau.SelLength = Len(txtnuevaau.Text)
End If
End Sub

Private Sub txtnuevaau_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Val(txtnuevaau.Text) > SaldoPendiente Then
        MsgBox "El pago a autorizar es mayor que el monto de la Orden de Compra, autorice un pago menor", vbInformation, "ATENCIÓN"
        txtnuevaau.Text = ""
        txtnuevosaldo.Text = "0.00"
        Exit Sub

    Else
        txtnuevosaldo.Text = Format(txtsaldopendiente.Text - Format(txtnuevaau.Text, "###,###,##0.00"), "###,###,##0.00")
        txtobs.SetFocus
    End If
End If
   
End Sub

Private Sub txtnuevosaldo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    txtobs.SetFocus
End If
   
End Sub

Private Sub GrabarOrdPago()
amovs_cab(0).campo = "IDOP": amovs_cab(0).valor = IDOP: amovs_cab(0).Tipo = "T"
amovs_cab(1).campo = "ORDEN": amovs_cab(1).valor = txtoc.Text: amovs_cab(1).Tipo = "T"
amovs_cab(2).campo = "FECHA": amovs_cab(2).valor = txt_fecha.Value: amovs_cab(2).Tipo = "F"
amovs_cab(3).campo = "USUARIO": amovs_cab(3).valor = wusuario: amovs_cab(3).Tipo = "T"
amovs_cab(4).campo = "MONEDA": amovs_cab(4).valor = Monedas: amovs_cab(4).Tipo = "T"
amovs_cab(5).campo = "IMPORTE": amovs_cab(5).valor = Format(txtnuevaau.Text, "0.00"): amovs_cab(5).Tipo = "N"
amovs_cab(6).campo = "CORRELADOC": amovs_cab(6).valor = 0: amovs_cab(6).Tipo = "N"
amovs_cab(7).campo = "CORRELAANTICIPO": amovs_cab(7).valor = 0: amovs_cab(7).Tipo = "N"
amovs_cab(8).campo = "OBSERVACION": amovs_cab(8).valor = "" & txtobs.Text: amovs_cab(8).Tipo = "T"
amovs_cab(9).campo = "ESTADO": amovs_cab(9).valor = "1": amovs_cab(9).Tipo = "T"
If sw_ordendepago = False Then
    amovs_cab(10).campo = "EST_AUT": amovs_cab(10).valor = 0: amovs_cab(10).Tipo = "N"
Else
    amovs_cab(10).campo = "EST_AUT": amovs_cab(10).valor = "" & Est_Aut: amovs_cab(10).Tipo = "N"
End If

If ctipo = "A" Then     '--- Nuevo
         '------- GRABA CABECERA
         GRABA_REGISTRO_logistica amovs_cab(), "IF4ORDEN_PAGO", ctipo, 10, cnn_dbbancos, ""
Else    '--- Modificación
        '------- GRABA CABECERA
        GRABA_REGISTRO_logistica amovs_cab(), "IF4ORDEN_PAGO", ctipo, 10, cnn_dbbancos, ""
'        'AlmacenaQuery_sql csql, cnn_dbbancos
End If

If sw_ordendepago = False Then
    cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set F4ESTADO=4 where IDOP = '" & IDOP & "'")
Else
    cnn_dbbancos.Execute ("Update IF4ORDEN_PAGO Set F4ESTADO=4 where IDOP = '" & IDOP & "'")
End If

End Sub

Public Sub GeneraCod()
Dim rst As New ADODB.Recordset

sql = "select IDOP from IF4ORDEN_PAGO order by IDOP desc"
rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    If "" & rst(0).Value = "" Then
        Codigo = "000001"
    Else
        If Val(rst(0).Value) > 0 Then
            Codigo = Format(Val(rst(0).Value) + 1, "000000")
        Else
            Codigo = ""
        End If
    End If
    IDOP = Codigo
Else
    IDOP = "000001"
End If
rst.Close
End Sub

Public Sub GeneraAut()
Dim RsM As New ADODB.Recordset

sql = "select max(est_aut) from IF4ORDEN_PAGO WHERE ORDEN = '" & txtoc.Text & "' "
If RsM.State = adStateOpen Then RsM.Close
RsM.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not RsM.EOF Then
    If "" & RsM(0).Value = "" Then
        CodigoAut = 1
    Else
        If Val(RsM(0).Value) >= 0 Then
            CodigoAut = Val(RsM(0).Value) + 1
        Else
            CodigoAut = ""
        End If
    End If
    Est_Aut = CodigoAut
Else
    Est_Aut = 1
End If
RsM.Close
End Sub

