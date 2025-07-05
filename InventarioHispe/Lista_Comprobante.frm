VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGRID.DLL"
Begin VB.Form Lista_Comprobante 
   ClientHeight    =   6075
   ClientLeft      =   1710
   ClientTop       =   2385
   ClientWidth     =   8640
   LinkTopic       =   "Lista"
   ScaleHeight     =   6075
   ScaleWidth      =   8640
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   5865
      Left            =   90
      OleObjectBlob   =   "Lista_Comprobante.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   8430
   End
End
Attribute VB_Name = "Lista_Comprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bancos As ADODB.Connection

Private Sub Form_Load()
    
    Set bancos = New ADODB.Connection
    Me.Caption = "Lista de Movimientos"
    With bancos
        .Provider = "Microsoft.JET.OLEDB.4.0; Data Source=" & wrutabancos & "\db_bancos.mdb; Persist Security Info=False"
        .Open
    End With
    Conf_Grid
    Llenado_Grid

End Sub

Private Sub Conf_Grid()
    
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
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoNameCaseInsensitive)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        .Set (egoDynamicLoad)
        .Set (egoRowSelect)
        
    End With
  
End Sub

Public Sub Llenado_Grid()
Dim SQL     As String

    dxDBGrid1.Dataset.ADODataset.ConnectionString = bancos
    SQL = "Select NOMBRE,RUC,SERIE,NUM_DOCUMENTO,ANULADO from retendoc order by SERIE,NUM_DOCUMENTO DESC"
    dxDBGrid1.Dataset.ADODataset.CommandText = SQL
    dxDBGrid1.Dataset.Active = True
    dxDBGrid1.KeyField = "NUM_DOCUMENTO"

End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    gserie = "" & dxDBGrid1.Columns(0).Value
    gnummov = "" & dxDBGrid1.Columns(1).Value
    Me.Hide

End Sub
