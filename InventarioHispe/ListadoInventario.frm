VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Begin VB.Form ListadoInventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Listado de Inventario :::"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "ListadoInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2100
      Left            =   135
      OleObjectBlob   =   "ListadoInventario.frx":000C
      TabIndex        =   0
      Top             =   300
      Width           =   6885
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   120
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   8
      Tools           =   "ListadoInventario.frx":163C
      ToolBars        =   "ListadoInventario.frx":7B68
   End
End
Attribute VB_Name = "ListadoInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
