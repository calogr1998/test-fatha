VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form ReporteGuiasVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Reporte de Guías de Ventas :::"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "ReporteGuiasVenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   4455
      Begin Threed.SSCommand SSCommand1 
         Height          =   330
         Left            =   2985
         TabIndex        =   2
         Top             =   555
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&Procesar"
      End
      Begin VB.Label Label2 
         Caption         =   "por motivo venta."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   615
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reporte de Guías de remision, vales de salida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   4110
      End
   End
End
Attribute VB_Name = "ReporteGuiasVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
