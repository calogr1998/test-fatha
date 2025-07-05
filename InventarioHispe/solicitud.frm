VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{791923BA-56CB-4A36-9EA3-1B4ED74622AA}#1.0#0"; "csimxctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form solicitud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimiento interno"
   ClientHeight    =   9240
   ClientLeft      =   225
   ClientTop       =   1785
   ClientWidth     =   19335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "solicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   19335
   Begin InternetMailCtl.InternetMail InternetMail1 
      Left            =   360
      Top             =   6960
      _cx             =   741
      _cy             =   741
      Enabled         =   -1  'True
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   19185
      _ExtentX        =   33840
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "solicitud.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblusuario"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lblcreacion"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Mon"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Grid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "SSActiveToolBars2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "SSPanel1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "SSPanel3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin Threed.SSPanel SSPanel3 
         Height          =   570
         Left            =   240
         TabIndex        =   8
         Top             =   7080
         Width           =   10725
         _Version        =   65536
         _ExtentX        =   18918
         _ExtentY        =   1005
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
         Begin VB.TextBox txtlugar 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   6
            Top             =   120
            Width           =   8955
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Lugar de Entrega"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   9
            Top             =   180
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   16995
         _Version        =   65536
         _ExtentX        =   29977
         _ExtentY        =   4260
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
         Begin VB.ComboBox cmbDerivado 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "solicitud.frx":0028
            Left            =   1380
            List            =   "solicitud.frx":0032
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1560
            Visible         =   0   'False
            Width           =   1995
         End
         Begin MSComCtl2.DTPicker txtfecha 
            Height          =   315
            Left            =   12600
            TabIndex        =   35
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   138018817
            CurrentDate     =   40611
         End
         Begin VB.TextBox txtobservaciones 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1380
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1920
            Width           =   15375
         End
         Begin VB.TextBox txtnumxobra 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   15300
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   21
            Top             =   780
            Visible         =   0   'False
            Width           =   1480
         End
         Begin VB.TextBox txtnumxusuario 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   15300
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   20
            Top             =   1125
            Visible         =   0   'False
            Width           =   1480
         End
         Begin VB.TextBox lblsolicitud 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   15120
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Nº"
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtsolicitud 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   15120
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   18
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cboprioridad 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "solicitud.frx":00A5
            Left            =   1380
            List            =   "solicitud.frx":00B2
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   765
            Width           =   1995
         End
         Begin VB.TextBox pnlproveedor 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1140
            Width           =   10005
         End
         Begin VB.TextBox txtproveedor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5460
            MaxLength       =   11
            TabIndex        =   16
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox txtsolicitante 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   7
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox txtuupp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5460
            MaxLength       =   15
            TabIndex        =   2
            Top             =   765
            Width           =   1215
         End
         Begin VB.TextBox txtdesuupp 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   765
            Width           =   10005
         End
         Begin VB.TextBox txttc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   14340
            MaxLength       =   8
            TabIndex        =   14
            Text            =   "2.580"
            Top             =   240
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cmbestado 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "solicitud.frx":00CA
            Left            =   1380
            List            =   "solicitud.frx":00CC
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1170
            Width           =   1995
         End
         Begin VB.Frame frmmoneda 
            Caption         =   " Moneda "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   680
            Left            =   9840
            TabIndex        =   11
            Top             =   30
            Width           =   2205
            Begin Threed.SSOption optmoneda 
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   780
               _Version        =   65536
               _ExtentX        =   1376
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "&Soles"
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
               Value           =   -1  'True
            End
            Begin Threed.SSOption optmoneda 
               Height          =   240
               Index           =   1
               Left            =   1035
               TabIndex        =   13
               Top             =   240
               Width           =   825
               _Version        =   65536
               _ExtentX        =   1455
               _ExtentY        =   423
               _StockProps     =   78
               Caption         =   "&Dólares"
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
            End
         End
         Begin Threed.SSCheck chkcerrar 
            Height          =   360
            Left            =   15240
            TabIndex        =   22
            Top             =   1140
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "Cerrada"
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Derivado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   43
            Top             =   1560
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label lbldescripcion 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   34
            Top             =   1920
            Visible         =   0   'False
            Width           =   5955
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   180
            TabIndex        =   33
            Top             =   1965
            Width           =   1110
         End
         Begin VB.Label lblnumxobra 
            AutoSize        =   -1  'True
            Caption         =   "Nº por obra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   14400
            TabIndex        =   32
            Top             =   840
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblnumxusuario 
            AutoSize        =   -1  'True
            Caption         =   "Nº por usuario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   14160
            TabIndex        =   31
            Top             =   1200
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   12120
            TabIndex        =   30
            Top             =   285
            Width           =   450
         End
         Begin VB.Label lbltc 
            AutoSize        =   -1  'True
            Caption         =   "T/C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   14040
            TabIndex        =   29
            Top             =   300
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblprioridad 
            Caption         =   "Prioridad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   180
            TabIndex        =   28
            Top             =   810
            Width           =   855
         End
         Begin VB.Label lblcosto 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor sugerido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3840
            TabIndex        =   27
            Top             =   1200
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblsolicitantex 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   180
            TabIndex        =   26
            Top             =   405
            Width           =   735
         End
         Begin VB.Label lbluupp 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Costo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3840
            TabIndex        =   25
            Top             =   840
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   24
            Top             =   1260
            Width           =   495
         End
         Begin VB.Label LblSolicitante 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2340
            TabIndex        =   23
            Top             =   360
            Width           =   7155
         End
      End
      Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars2 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131082
         ToolBarsCount   =   1
         ToolsCount      =   9
         Tools           =   "solicitud.frx":00CE
         ToolBars        =   "solicitud.frx":7293
      End
      Begin DXDBGRIDLibCtl.dxDBGrid Grid 
         Height          =   4005
         Left            =   0
         OleObjectBlob   =   "solicitud.frx":737B
         TabIndex        =   5
         Top             =   3000
         Width           =   18855
      End
      Begin CONTROLSLibCtl.dxColorBtn Mon 
         Height          =   405
         Left            =   11160
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Soles"
         Top             =   2040
         Width           =   405
         _Version        =   65536
         _cx             =   706
         _cy             =   706
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FillColor       =   65535
         ForeColor       =   12648447
         Caption         =   "MN"
         Enabled         =   -1  'True
         CaptionStringCount=   1
         GroupIndex      =   -1
         Stuck           =   -1  'True
         PictureLayout   =   1
         Pushed          =   0   'False
      End
      Begin VB.Label Lblcreacion 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   13200
         TabIndex        =   42
         Top             =   7680
         Width           =   3885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de creación: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   11520
         TabIndex        =   41
         Top             =   7680
         Width           =   1440
      End
      Begin VB.Label lblusuario 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1920
         TabIndex        =   40
         Top             =   7680
         Width           =   8925
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usuario: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   39
         Top             =   7680
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   11520
         TabIndex        =   38
         Top             =   7200
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10200
         TabIndex        =   37
         Top             =   2160
         Width           =   690
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   582
      _ExtentY        =   582
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      ActiveColors    =   -1  'True
      Tools           =   "solicitud.frx":FCCC
      ToolBars        =   "solicitud.frx":17B2D
   End
End
Attribute VB_Name = "solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Af As New ADOFunctions
Dim RsC As New ADODB.Recordset
Dim rst             As New ADODB.Recordset
Dim solicitud       As String * 12
Dim items           As Byte
Dim seleccion       As Boolean
Dim cnombase        As String
Dim cnomtabla       As String
Dim CadSql          As String
Dim cconex_form     As String
Dim cnn_form        As New ADODB.Connection
Dim sw_nuevo_item   As Boolean
Dim sw_activate     As Boolean
Dim sw_cabecera     As Boolean
Dim sw_detalle      As Boolean
Dim Cantidad        As Double
Dim sql             As String

Dim Values()            As Variant
'Dim amovs_cab(0 To 27)  As a_grabacion
Dim amovs_cab(0 To 26)  As a_grabacion
Dim amovs_det(0 To 27)  As a_grabacion
Dim ctipo               As String * 1
Dim cvalores            As String
Dim RSDETALLE           As New ADODB.Recordset
Dim sw_ayuda_prod       As Boolean
Dim rsmedidas           As New ADODB.Recordset
Dim sw_ayuda            As Boolean
Dim wcerrado            As String * 1
Dim wexcel              As Byte
Dim rsdescuento         As ADODB.Recordset
Dim V1 As Integer
'Private cImgInfo As cImageInfo


Dim sigv As Double, cigv As Double, Num As Double
Private Sub elimina(pnumero As String)
'On Error GoTo ERROR_ELIMINA
ReDim amovs(0 To 0) As a_grabacion
Dim cmes            As String * 2
Dim sw_elimina      As Boolean

    If Len(Trim("" & pnumero)) = 0 Then
        MsgBox "La Solicitud no ha Sido Grabada", 16, "Sistema de Logística"
        Exit Sub
    End If
    
    If MsgBox("¿Está Seguro(a) de Eliminar La Solicitud Nº " & txtsolicitud.Text & "?", vbYesNo + vbDefaultButton2 + vbQuestion, "Atención") = vbYes Then
        If rsif4orden.State = adStateOpen Then rsif4orden.Close
        rsif4orden.Open "SELECT F4NUMORD FROM IF4ORDEN WHERE F4CODSOLICITUD='" & pnumero & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif4orden.EOF Then
            If MsgBox("La solicitud está asociada la Orden de Compra " & rsif4orden!F4NUMORD & ", ¿está seguro(a) de eliminar la solicitud?", vbYesNo, "Atención") = vbYes Then
                sw_elimina = True
            Else
                sw_elimina = False
            End If
        Else
            sw_elimina = True
        End If
        
        If sw_elimina = True Then
            If ctipoadm_bd = "M" Then
                sql = ("DELETE FROM TB_DETSOLICITUD WHERE COD_SOLICITUD = '" & pnumero & "'")
                cnn_dbbancos.Execute sql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                Actualiza_Log sql, cnn_dbbancos.ConnectionString
                sql = ("DELETE FROM TB_CABSOLICITUD WHERE COD_SOLICITUD = '" & pnumero & "'")
                cnn_dbbancos.Execute sql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                Actualiza_Log sql, cnn_dbbancos.ConnectionString
            Else
            
                sql = ("DELETE * FROM TB_DETSOLICITUD WHERE COD_SOLICITUD = '" & pnumero & "'")
                cnn_dbbancos.Execute sql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                Actualiza_Log sql, cnn_dbbancos.ConnectionString
                sql = ("DELETE * FROM TB_CABSOLICITUD WHERE COD_SOLICITUD = '" & pnumero & "'")
                cnn_dbbancos.Execute sql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                Actualiza_Log sql, cnn_dbbancos.ConnectionString
            End If
            DELETEREC_LOG cnomtabla, cnn_form
            
            nuevo
            AdicionaItem
        End If
    End If
    
    Exit Sub
    
ERROR_ELIMINA:
    MsgBox "Ha ocurrido el sgte. error " & Err.Description, vbCritical, "Atencion"
    Resume Next
    
End Sub

Sub pasaVariable()
        cigv = Grid.Columns.ColumnByFieldName("ds_conIgv").Value & ""
        sigv = Grid.Columns.ColumnByFieldName("ds_sinIgv").Value & ""
End Sub

Private Sub cboprioridad_Change()

    If Len(Trim(cboprioridad.Text)) > 0 And sw_cabecera = False Then
        sw_cabecera = True
    End If
    
End Sub

Private Sub cboprioridad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub


Private Sub chkcerrar_Click(Value As Integer)
If wcerrado = "S" Then
    chkcerrar = True
Else
    chkcerrar = False
End If
End Sub

Private Sub cmbestado_Click()

    If Len(Trim(txttc.Text)) > 0 And sw_cabecera = False Then
        sw_cabecera = True
    End If

End Sub

Private Sub cmbestado_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
        
    End If

End Sub

Private Sub Grid_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case UCase("ds_conIgv")
    Text = Format(Text, "###,###,##0.0000")
Case UCase("DS_SINIGV")
    Text = Format(Text, "###,###,##0.0000")
Case UCase("precio")
    Text = Format(Text, "###,###,##0.00")
End Select
End Sub
Private Sub Grid_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)

Select Case UCase(Grid.Columns.FocusedColumn.FieldName)
Case UCase("afecto"):
                sw_detalle = True
                Grid.Dataset.Edit
                sigv = 0
                cigv = 0
                Num = 0
                pasaVariable
                If Grid.Columns.ColumnByFieldName("afecto").Value = False Then
                        Grid.Columns.ColumnByFieldName("afecto").Value = True
                        Num = Val(sigv * wIgv) / 100
                        Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Val(sigv + Num)
                    ' Grid.Columns.ColumnByFieldName("afecto").Value = True
                     Else
                        Grid.Columns.ColumnByFieldName("afecto").Value = False
                         Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Grid.Columns.ColumnByFieldName("ds_sinIgv").Value
                End If
                Grid.Columns.ColumnByFieldName("PRECIO").Value = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value * Grid.Columns.ColumnByFieldName("ds_conIgv").Value)
                Grid.Dataset.Post
    Grid.Columns.FocusedIndex = 7

End Select
End Sub

Private Sub Grid_OnCustomDrawFooter(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Select Case UCase(Column.FieldName)
Case UCase("precio")
    Text = Format(Text, "###,###,##0.00")
Case UCase("DS_CANTIDAD")
    Text = Format(Text, "###,###,##0.00")
    If Val(Text) > 0 Then
        Color = vbBlue
    Else
        Color = vbRed
    End If
    FontColor = vbWhite
End Select
End Sub

Private Sub Grid_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    Dim afec As Integer
    Dim x As Integer
    Dim z As Integer
    
    Select Case UCase(Grid.Columns.FocusedColumn.FieldName)

        Case "F5CODFAB", "COD_PRODUCTO":
        'If Grid.Columns.FocusedColumn.FieldName = "F5CODFAB" Or Grid.Columns.FocusedColumn.FieldName = "cod_producto" Then
            wcodproducto = ""
            wcod_alm = ""
            sw_ayuda_prod = True
            wllamada = 1
            If Len(Trim(txtuupp.Text)) = 0 Then
                    MsgBox "Debe consignar un Centro de Costo de manera obligatoria", vbExclamation, "Sistema de Logística"
                    Exit Sub
            End If
            If Grid.Columns.ColumnByFieldName("PARTIDA").Value = "" Then
                If txtuupp.Text = "013" Then
                    MsgBox "La partida de presupuesto es obligatoria para esta obra", vbExclamation, "Sistema de Logística"
                    Exit Sub
                Else
                    wcodpartida = ""
                    wcodpresupuesto = ""
                End If
            Else
                wcodpartida = Grid.Columns.ColumnByFieldName("PARTIDA").Value
                wcodpresupuesto = txtuupp.Text
            End If
            ayuda_productos_partida.Show 1
            With ayuda_productos_partida.dxDBGrid1
                    .Dataset.Filtered = True
                    .Dataset.Filter = "F4PERINT = -1"
                    .Dataset.First
                    x = 0
                        Do While Not .Dataset.EOF
                            z = .Dataset.RecordCount
                            If z = 0 Then Exit Sub
                            x = x + 1
                                If Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = "" Then
                                    Grid.Dataset.Edit
                                Else
                                    Grid.Dataset.Append
                                End If
                                Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = .Columns.ColumnByFieldName("f5codpro").Value
                                Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value = .Columns.ColumnByFieldName("f5nompro").Value
                                Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = .Columns.ColumnByFieldName("F7SIGMED").Value 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida & "", "T", cnn_dbbancos)
                                Grid.Columns.ColumnByFieldName("DS_desUNIMED").Value = .Columns.ColumnByFieldName("F7SIGMED").Value 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida & "", "T", cnn_dbbancos)
                                Grid.Columns.ColumnByFieldName("STOK").Value = IIf(IsNull(.Columns.ColumnByFieldName("f6stockact").Value), 0, .Columns.ColumnByFieldName("f6stockact").Value)
                                Grid.Columns.ColumnByFieldName("PRECIO").Value = IIf(IsNull(.Columns.ColumnByFieldName("f5vtanet").Value), 0, .Columns.ColumnByFieldName("f5vtanet").Value)
                                Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value = .Columns.ColumnByFieldName("f5fob").Value
                                Grid.Columns.ColumnByFieldName("CS_FENTREGA").Value = Format(txtfecha.Value, "dd/mm/yyyy")
                                Grid.Columns.ColumnByFieldName("F5CODFAB").Value = .Columns.ColumnByFieldName("f5codfab").Value
                                Grid.Columns.ColumnByFieldName("F5CODMARCA").Value = "" 'wcodmarca
                                Grid.Columns.ColumnByFieldName("F5MARCA").Value = "" 'wmarca
                                Grid.Columns.ColumnByFieldName("ds_sinIgv").Value = IIf(IsNull(.Columns.ColumnByFieldName("f5vtanet").Value), 0, .Columns.ColumnByFieldName("f5vtanet").Value)
                                Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Format(Grid.Columns.ColumnByFieldName("ds_sinIgv").Value * 1.18, "0.00")
                                Grid.Columns.ColumnByFieldName("F5CODCOSTO").Value = txtuupp.Text
                                Grid.Columns.ColumnByFieldName("F5DESCOSTO").Value = txtdesuupp.Text
                                Grid.Columns.ColumnByFieldName("ruc_proveedor").Value = txtproveedor.Text
                                Grid.Columns.ColumnByFieldName("NOMPROV").Value = pnlproveedor.Text
                             '
                             If .Columns.ColumnByFieldName("f5afecto").Value = "*" Then
                                 afec = 1
                             Else
                                 afec = 0
                             End If
                             Grid.Columns.ColumnByFieldName("afecto").Value = afec
                             Grid.Dataset.Post
                            .Dataset.Next
                        Loop
                        If x = 0 And Len(Trim(wcodproducto)) > 0 Then
                            Grid.Dataset.Edit
                            Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = wcodproducto
                            Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value = wdesproducto
                            Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = wmedida 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida & "", "T", cnn_dbbancos)
                            Grid.Columns.ColumnByFieldName("DS_desUNIMED").Value = ObtenerCampo("ef7medidas", "f7sigmed", "f7sigmed", wmedida & "", "T", cnn_dbbancos) 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida & "", "T", cnn_dbbancos)
                            Grid.Columns.ColumnByFieldName("STOK").Value = wstockact
                            Grid.Columns.ColumnByFieldName("PRECIO").Value = wprecos
                            Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value = 1
                            Grid.Columns.ColumnByFieldName("CS_FENTREGA").Value = Format(txtfecha.Value, "dd/mm/yyyy")
                            Grid.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
                            Grid.Columns.ColumnByFieldName("F5CODMARCA").Value = wcodmarca
                            Grid.Columns.ColumnByFieldName("F5MARCA").Value = "" 'wmarca
                            Grid.Columns.ColumnByFieldName("ds_sinIgv").Value = wprecos
                            Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Format(wprecos * 1.18, "0.00")
                            Grid.Columns.ColumnByFieldName("F5CODCOSTO").Value = txtuupp.Text
                            Grid.Columns.ColumnByFieldName("F5DESCOSTO").Value = txtdesuupp.Text
                            Grid.Columns.ColumnByFieldName("ruc_proveedor").Value = txtproveedor.Text
                            Grid.Columns.ColumnByFieldName("NOMPROV").Value = pnlproveedor.Text
                             '
                            If wafecto = "*" Then
                                afec = 1
                            Else
                                afec = 0
                            End If
                            Grid.Columns.ColumnByFieldName("afecto").Value = afec
                            sw_detalle = True
                            Grid.Dataset.Post
                        End If
                    Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Index
                    Unload ayuda_productos_partida
                    
            End With
        Case "F5DESCOSTO":
                wcodcosto = "": wdescosto = "": wunicosto = "":
                Ayuda_Centros.Show 1
                If Len(Trim(wcodcosto)) > 0 Then
                    Grid.Dataset.Edit
                    Grid.Columns.ColumnByFieldName("F5CODCOSTO").Value = wcodcosto
                    Grid.Columns.ColumnByFieldName("F5DESCOSTO").Value = wdescosto
                    sw_detalle = True
                    Grid.Dataset.Post
                End If
        Case "NOMPROV":
                wrucprov = "": wnomprov = ""
                ayuda_proveedores_log.Show 1
                If Len(Trim(wrucprov)) > 0 Then
                    Grid.Dataset.Edit
                    Grid.Columns.ColumnByFieldName("ruc_proveedor").Value = wrucprov
                    Grid.Columns.ColumnByFieldName("NOMPROV").Value = wnomprov
                    Grid.Dataset.Post
                    sw_detalle = True
                End If
                sw_ayuda = True
                
    End Select

    If Grid.Columns.FocusedColumn.ObjectName = "COLUMNELIMINAR" Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
            sw_nuevo_item = True
            If Grid.Count = 1 Then
                Grid.Dataset.Delete
                AdicionaItem
                sw_detalle = False
                'atbmenu.Tools("IDGrabar").Enabled = False
            Else
                Grid.Dataset.Delete
            End If
            
            sw_nuevo_item = False
        End If
    ElseIf Grid.Columns.FocusedColumn.ObjectName = "Partida" Then
        wnumordentrab = ""
        wcod_alm = ""
        sw_ayuda_prod = True
        wllamada = 1
        ayuda_orden_trab.Show 1
        Unload ayuda_orden_trab
        If Len(Trim(wnumordentrab)) > 0 Then
            Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("PARTIDA").Value = wnumordentrab
            Grid.Columns.ColumnByFieldName("DES_PARTIDA").Value = wobservacion
            Grid.Dataset.Post
            'Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Index
        End If

    End If
End Sub


Private Sub Grid_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Dim cunidad     As String
    
    If sw_nuevo_item = False Then
        
       ' Select Case Grid.Columns.FocusedColumn.FieldName
     Select Case (Grid.Columns.FocusedColumn.FieldName)

            Case "F5CODFAB":
                If Len(Trim(Grid.Columns.ColumnByFieldName("F5CODFAB").Value)) > 0 Then
                    sw_detalle = True
                    wcodfab1 = Grid.Columns.ColumnByFieldName("F5CODFAB").Value & ""
                    If wllamada = 1 Then
                        sql = "select f5nompro,f5stockact, F5FOB,f7codmed,f5codfab,f5marca,f5codpro from if5pla where f5codfab='" & wcodfab1 & "' and f5marca='" & wcodmarca & "'"
                        wllamada = 0
                    Else
                        sql = "select f5nompro,f5stockact, F5FOB,f7codmed,f5codfab,f5marca,f5codpro from if5pla where f5codfab='" & wcodfab1 & "'"
                    End If
                    If rst.State = adStateOpen Then rst.Close
                    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                    If Not (rst.EOF) Then
                        Grid.Dataset.Edit
                        Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value = rst!F5NOMPRO & ""
                        Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = rst!f5codpro & ""
                        cunidad = ""
                        If rsmedidas.State = adStateOpen Then rsmedidas.Close
                        rsmedidas.Open "SELECT F7SIGMED FROM EF7MEDIDAS WHERE EF7MEDIDAS.F7CODMED='" & rst!f7codmed & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsmedidas.EOF Then
                            cunidad = rsmedidas.Fields("F7SIGMED") & ""
                        End If
                        rsmedidas.Close
                        Set rsmedidas = Nothing
                        
                        Grid.Columns.ColumnByFieldName("F5CODMARCA").Value = "" 'Trim(rst.Fields("F5MARCA") & "")
                        If rsmarcas.State = adStateOpen Then rsmarcas.Close
                        rsmarcas.Open "SELECT F2DESMAR FROM EF2MARCAS WHERE F2CODMAR='" & rst.Fields("F5MARCA") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                        If Not rsmarcas.EOF Then
                            Grid.Columns.ColumnByFieldName("f5marca").Value = "" 'rsmarcas.Fields("F2DESMAR")
                        End If
                        rsmarcas.Close
                        Set rsmarcas = Nothing
                                                                        
                        Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = cunidad
                        Grid.Columns.ColumnByFieldName("STOK").Value = IIf(IsNull(rst!f5stockact), Null, rst!f5stockact)
                        Grid.Columns.ColumnByFieldName("PRECIO").Value = IIf(IsNull(rst!F5FOB), Null, rst!F5FOB)
                        Grid.Columns.ColumnByFieldName("CS_FENTREGA").Value = Format(txtfecha.Value, "dd/mm/yyyy")
                    Else
                        Grid.Dataset.Edit
                        Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = ""
                        Grid.Columns.ColumnByFieldName("F5CODFAB").Value = ""
                        Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value = ""
                        Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = ""
                        Grid.Columns.ColumnByFieldName("STOK").Value = Null
                        Grid.Columns.ColumnByFieldName("PRECIO").Value = Null
                    End If
                    rst.Close
                    Grid.Columns.FocusedIndex = 2
                Else
                    Grid.Columns.FocusedIndex = 0
                End If
'            Case "afecto":
'                sw_detalle = True
'                Grid.Dataset.Edit
'                pasaVariable
'                If Grid.Columns.ColumnByFieldName("afecto").Value = False Then
'                        Grid.Columns.ColumnByFieldName("ds_sinIgv").Value = Grid.Columns.ColumnByFieldName("ds_conIgv").Value
'                     Else
'                        Num = Val(sigv * wIgv) / 100
'                        Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Val(sigv + Num)
'                End If
'                Grid.Dataset.Post
'                Grid.Columns.FocusedIndex = 7
            Case "ds_conIgv":
                sw_detalle = True
                Grid.Dataset.Edit
                pasaVariable
                If Grid.Columns.ColumnByFieldName("afecto").Value = True Then
                 Num = Val(cigv * wIgv) / 100
                 'Grid.Columns.ColumnByFieldName("ds_sinIgv").Value = Val(cigv - Num)
                 Grid.Columns.ColumnByFieldName("ds_sinIgv").Value = Format(Val(cigv / Val(1# + wIgv / 100)), "0.0000")
                 
                'Dim sigv As Double, cigv As Double
                Else
                 Grid.Columns.ColumnByFieldName("ds_sinIgv").Value = Grid.Columns.ColumnByFieldName("ds_conIgv").Value
                End If
                    Grid.Columns.ColumnByFieldName("PRECIO").Value = Format(Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value * Grid.Columns.ColumnByFieldName("ds_conIgv").Value) * (1 - Val(Grid.Columns.ColumnByFieldName("DESCUENTO").Value) / 100), "0.00")
                    Grid.Dataset.Post
                    Grid.Columns.FocusedIndex = 8
            
            Case "ds_sinIgv":
                sw_detalle = True
                Grid.Dataset.Edit
                pasaVariable
                If Grid.Columns.ColumnByFieldName("afecto").Value = True Then
                    Num = Val(sigv * wIgv) / 100
                    Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Val(sigv + Num)
                Else
                    Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Grid.Columns.ColumnByFieldName("ds_sinIgv").Value
                End If
                    Grid.Columns.ColumnByFieldName("PRECIO").Value = Format(Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value * Grid.Columns.ColumnByFieldName("ds_conIgv").Value) * (1 - Val(Grid.Columns.ColumnByFieldName("DESCUENTO").Value) / 100), "0.00")
                    Grid.Dataset.Post
                    Grid.Columns.FocusedIndex = 8
            Case "DS_CANTIDAD":
                sw_detalle = True
                Grid.Dataset.Edit
                pasaVariable
                If Grid.Columns.ColumnByFieldName("afecto").Value = True Then
                    Num = Val(sigv * wIgv) / 100
                    Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Val(sigv + Num)
                 Else
                    Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Grid.Columns.ColumnByFieldName("ds_sinIgv").Value
                End If
                Grid.Columns.ColumnByFieldName("PRECIO").Value = Format(Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value * Grid.Columns.ColumnByFieldName("ds_conIgv").Value) * (1 - Val(Grid.Columns.ColumnByFieldName("DESCUENTO").Value) / 100), "0.00")
                Grid.Dataset.Post
                Grid.Columns.FocusedIndex = 7
            Case "DESCUENTO":
                sw_detalle = True
                Grid.Dataset.Edit
                pasaVariable
                If Grid.Columns.ColumnByFieldName("afecto").Value = True Then
                    Num = Val(sigv * wIgv) / 100
                    Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Val(sigv + Num)
                 Else
                    Grid.Columns.ColumnByFieldName("ds_conIgv").Value = Grid.Columns.ColumnByFieldName("ds_sinIgv").Value
                End If
                Grid.Columns.ColumnByFieldName("PRECIO").Value = Format(Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value * Grid.Columns.ColumnByFieldName("ds_conIgv").Value) * (1 - Val(Grid.Columns.ColumnByFieldName("DESCUENTO").Value) / 100), "0.00")
                Grid.Dataset.Post
                Grid.Columns.FocusedIndex = 0
            Case "NOMPROV":
                sw_detalle = True
                Grid.Dataset.Edit
                If Len(Grid.Columns.ColumnByFieldName("NOMPROV").Value) = 0 Then
                    Grid.Columns.ColumnByFieldName("ruc_proveedor").Value = ""
                End If
                Grid.Dataset.Post
            Case Else
                If wcerrado <> "S" Then
                    sw_detalle = True
                    Grid.Dataset.Edit
'                    wprecio = Grid.Columns.ColumnByFieldName("PRECIO").Value
'                    wtotdscto = wprecio * Grid.Columns.ColumnByFieldName("PORCDSCTO").Value / 100
'                    If Not IsNumeric(wtotdscto) Then
'                        wtotdscto = 0
'                    End If
'                    Grid.Columns.ColumnByFieldName("TOTDSCTO").Value = Format(wtotdscto, "#,###,###0.00")
'
'                    Grid.Columns.ColumnByFieldName("TOTAL").Value = Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value * Grid.Columns.ColumnByFieldName("PRECIO").Value - wtotdscto
'                    Grid.Dataset.Refresh
'                    Grid.Columns.FocusedIndex = 5
                End If
        End Select
    End If
    sw_detalle = True
End Sub
Private Sub grid_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)

    If KeyCode = 115 Or KeyCode = 46 Then
        If MsgBox("¿Desea Eliminar el ítem seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención") = vbYes Then
            sw_nuevo_item = True
            Grid.Dataset.Delete
            sw_nuevo_item = False
        End If
    End If
    
    If KeyCode = 113 Then
        If Grid.Columns.FocusedIndex = 0 Then
            wcodproducto = ""
            wcod_alm = ""
            sw_ayuda_prod = True
            Me.MousePointer = vbHourglass
            ayuda_productos.Show 1
            Me.MousePointer = vbDefault
            If Len(Trim(wcodproducto)) > 0 Then
                Grid.Dataset.Edit
                Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = wcodproducto
                Grid.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
                Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = wmedida
                Grid.Columns.ColumnByFieldName("DS_desUNIMED").Value = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida & "", "T", cnn_dbbancos)
                Grid.Columns.ColumnByFieldName("F5CODMARCA").Value = wcodmarca
                Grid.Columns.ColumnByFieldName("F5MARCA").Value = "" 'wmarca
                If rst.State = adStateOpen Then rst.Close
                sql = "select f5nompro,f5stockact, f5precos,f7codmed from if5pla where f5codpro='" & wcodproducto & "'"
                rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                If Not (rst.EOF) Then
                    Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value = rst!F5NOMPRO & ""
                    Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = rst!f7codmed & ""
                    Grid.Columns.ColumnByFieldName("STOK").Value = IIf(IsNull(rst!f5stockact), Null, rst!f5stockact)
                    Grid.Columns.ColumnByFieldName("PRECIO").Value = IIf(IsNull(rst!F5PRECOS), Null, rst!F5PRECOS)
                    Grid.Columns.ColumnByFieldName("CS_FENTREGA").Value = Format(txtfecha.Value, "dd/mm/yyyy")
                End If
                rst.Close
                Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Index - 2
            End If
        End If
    End If
    
End Sub
Private Sub grid_OnKeyPress(Key As Integer)
If Grid.Columns.FocusedColumn.FieldName = "F5CODFAB" Then
    Key = valida(1, Key, , True)
End If
End Sub
Private Sub grid_OnKeyUp(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 113 Then
    wcodproducto = ""
        wcod_alm = ""
        sw_ayuda_prod = True
        wllamada = 1
        ayuda_productos.Show 1
        If Len(Trim(wcodproducto)) > 0 Then
            Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value = wcodproducto
            Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value = wdesproducto
            Grid.Columns.ColumnByFieldName("DS_UNIDMED").Value = wmedida
            Grid.Columns.ColumnByFieldName("DS_desUNIMED").Value = ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", wmedida & "", "T", cnn_dbbancos)
            Grid.Columns.ColumnByFieldName("STOK").Value = wstockact
            Grid.Columns.ColumnByFieldName("PRECIO").Value = wprecos
            Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value = 0
            Grid.Columns.ColumnByFieldName("CS_FENTREGA").Value = Format(txtfecha.Value, "dd/mm/yyyy")
            Grid.Columns.ColumnByFieldName("F5CODFAB").Value = wcodfab
            Grid.Columns.ColumnByFieldName("F5CODMARCA").Value = wcodmarca
            Grid.Columns.ColumnByFieldName("F5MARCA").Value = ""
            Grid.Dataset.Post
            Grid.Columns.FocusedIndex = Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Index - 1
        End If
End If
End Sub

Private Sub grid_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Single, ByVal Y As Single)

    If Grid.Columns.FocusedIndex = 1 Then
        If Len(Trim("" & Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value)) > 0 Then
            lbldescripcion.Visible = True
            lbldescripcion.Caption = Trim("" & Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value)
        Else
            lbldescripcion.Caption = ""
            lbldescripcion.Visible = False
        End If
    Else
        lbldescripcion.Caption = ""
        lbldescripcion.Visible = False
    End If
        
End Sub

Private Sub Form_Activate()

    If sw_activate = False Then
        sw_activate = True
        If txtfecha.Enabled Then
            'txtfecha.SetFocus
        End If
    End If
    
    Me.left = (Screen.Width - Me.Width) / 2
    Me.top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Load()

'On Error GoTo CapturaError
    Me.MousePointer = vbHourglass
    wllamada = 0
    
    ActualizaProductos
    
    wexcel = 0
    wcerrado = "N"
    
    If Trim(wtiposalida) = "*" Then  '----- LA EMPRESA ES CONSTRUCTORA (AIC)
        lblnumxobra.Visible = True
        txtnumxobra.Visible = True
        lblnumxusuario.Visible = True
        txtnumxusuario.Visible = True
    Else
        lblnumxobra.Visible = False
        txtnumxobra.Visible = False
        lblnumxusuario.Visible = False
        txtnumxusuario.Visible = False
    End If
    
    If wf1show_ccosto = "N" Then
        lblcosto.Visible = False
        lblcliente.Visible = False
        txtcliente.Visible = False
    Else
        lblcosto.Visible = True
    End If
    
'    Me.Height = 7710
'    Me.left = 1600
'    Me.top = 1150
'    Me.Width = 12890
        
    If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
    cnn_dbbancos.Open StrConexDbBancos
    sw_ayuda_prod = False
    sw_nuevo_item = False
    sw_activate = False
    cmbestado.Enabled = False
    Call inicio
    Call CargarEstado
    Call CargarPrioridad
    
    
    sw_nuevo_item = False
    cnombase = "templus.mdb" '"TMP_BANCOS.mdb"
    cnomtabla = "solicitud"
    
    Set rsdescuento = New ADODB.Recordset
    cconex_form = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutatemp & "\" & cnombase & ";Persist Security Info=False"
    If cnn_form.State = adStateOpen Then cnn_form.Close
    cnn_form.Open cconex_form
    Call CONFIGURA_GRID
    If wTipoReq = 1 Then
        Me.Caption = "Requerimiento de Compra"
    Else
        Me.Caption = "Requerimiento de Servicio"
    End If
    
    
    sw_detalle = False
    If sw_nuevo_documento = True Then   'Nueva Solicitud
        'DELETEREC_LOG cnomtabla, cnn_form
        'grid.Dataset.Refresh
        nuevo
        'V1 = 0
        'Do While grid.Dataset.RecordCount = 0
             sw_nuevo_documento = False
            AdicionaItemNuevo
            sw_nuevo_documento = True
        '    V1 = V1 + 1
        'Loop
        sw_cabecera = False
    Else    'Modifica la Solicitud
        sw_cabecera = True
        nuevo
        SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
        SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
        SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Enabled = True
        BUSCA_SOLICITUD
        sw_cabecera = False
        sw_nuevo_documento = False
    End If
    SSActiveToolBars1.Tools.ITEM("ID_Anular").Visible = False
    Me.MousePointer = vbDefault
    Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    Resume Next
    Exit Sub
End Sub

Public Sub inicio()
    
    wcodsolicitante = "": wnomsolicitante = ""
    wcodproducto = "": wdesproducto = ""
    seleccion = False: items = 0
    txtfecha.Value = Format(Date, "dd/mm/yyyy")
    txttc.Text = "0.000"
    CodFirmaSolicitud(1) = "": CodFirmaSolicitud(2) = ""
    CodFirmaAprobacion(1) = "": CodFirmaAprobacion(2) = ""
    
    If rscambios.State = adStateOpen Then rscambios.Close
    If ctipoadm_bd = "M" Then
        sql = "SELECT * FROM CAMBIOS WHERE FECHA= '" & txtfecha.Value & "'"
    Else
        sql = "SELECT * FROM CAMBIOS WHERE FECHA= CVDate( '" & txtfecha.Value & "' )"
    End If
    
    rscambios.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rscambios.EOF Then
        txttc.Text = Format(Val("" & rscambios.Fields("cambio")), "0.000")
    End If
    rscambios.Close
    
End Sub


Private Sub Form_Unload(Cancel As Integer)



    sw_nuevo_item = True
    'grid.Dataset.Post
    Grid.Dataset.Close
    cnn_form.Close
    
    ELIMINA_BD_N wrutatemp, cnombase
    
    
    
    If sw_ayuda_prod = True Then
        Unload ayuda_productos
    End If
    
    lista_solicitudes.dxDBGrid1.Dataset.Active = False
    lista_solicitudes.dxDBGrid1.Dataset.Refresh
    lista_solicitudes.dxDBGrid1.Dataset.Active = True
    
End Sub



Private Sub Mon_Click()
If Mon.Caption = "US" Then
    Mon.Caption = "MN"
    Mon.ForeColor = &H80FFFF
    Mon.FillColor = &HFFFF&
    Mon.Pushed = 0
    'PnlSigMon(0).Caption = "MN"
    wMon = "S"
    'PnlOficial.Visible = False
    'PnlBasImp(0).BackColor = &HC0FFFF
    'PnlMonIna(0).BackColor = &HC0FFFF
    'txtigv(0).BackColor = &HC0FFFF
    'TxtOtrImp(0).BackColor = &HC0FFFF
    'PnlTotal(0).BackColor = &HC0FFFF
Else
    Mon.Caption = "US"
    Mon.ForeColor = &HC0FFC0
    Mon.FillColor = &H8000&
    Mon.Pushed = 0
    'PnlSigMon(0).Caption = "US"
    wMon = "D"
    'PnlBasImp(0).BackColor = &HC0FFC0
    'PnlMonIna(0).BackColor = &HC0FFC0
    'txtigv(0).BackColor = &HC0FFC0
    'TxtOtrImp(0).BackColor = &HC0FFC0
    'PnlTotal(0).BackColor = &HC0FFC0
End If
End Sub

Private Sub optmoneda_Click(Index As Integer, Value As Integer)
sw_cabecera = True
End Sub

Private Sub optmoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txttc.SetFocus
    Else
        If UCase$(Chr$(KeyAscii)) = "S" Then
            optmoneda(0).Value = True
            txttc.SetFocus
        ElseIf UCase$(Chr$(KeyAscii)) = "D" Then
            optmoneda(1).Value = True
            txttc.SetFocus
        End If
    End If

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.Id
        Case "ID_Nuevo":
            Me.MousePointer = vbHourglass
            sw_nuevo_documento = False
            sw_detalle = False
            sw_cabecera = False
            nuevo
            AdicionaItemNuevo
            sw_nuevo_documento = True
            Me.MousePointer = vbDefault
            activar True
            txtfecha.Enabled = True
            txtfecha.SetFocus
            SSActiveToolBars1.Tools.ITEM("ID_Grabar").Enabled = True
        Case "ID_Grabar":
            If Trim$(txtsolicitante.Text) = Empty Then
                MsgBox "Debe Ingresar Solicitante", vbExclamation, "Sistema de Logística"
                txtsolicitante.Enabled = True
                txtsolicitante.SetFocus
                Exit Sub
            End If

            If txtdesuupp.Text = "" Then
                MsgBox "Debe Seleccionar un Centro de Costo", vbInformation, "Sistema de Logística"
                txtuupp.SetFocus
                Exit Sub
            End If

            'If optmoneda(0).Value = False And optmoneda(1).Value = False And optmoneda(2).Value = False Then
            '    MsgBox "Debe Seleccionar Moneda de la Operación", vbExclamation, "Sistema de Logística"
            '    optmoneda(0).SetFocus
            '    Exit Sub
            'End If
            'If Val(txttc.Text) = 0 Or (Not (IsNumeric(txttc.Text))) Then
            '    MsgBox "Debe Ingresar Tipo de Cambio", vbExclamation, "Sistema de Logística"
            '    txttc.SetFocus
            '    Exit Sub
            'End If

            Me.MousePointer = vbHourglass
            If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
                sw_nuevo_item = True
                Grid.Dataset.Post
                sw_detalle = True
            End If
            If sw_cabecera = True Or sw_detalle = True Then
                If Val(ObtenerCampo("TB_CABSOLICITUD", "cs_estado", "cod_solicitud", txtsolicitud.Text, "T", cnn_dbbancos)) > 1 Then
                    MsgBox "La solicitud ya ha sido aprobada. No se puede modificar", vbInformation, "Sistema de Logística"
                Else
                    grabar
                End If
            Else
                MsgBox "No ha realizado ninguna modificación, no requiere grabar", vbInformation, "Sistema de Logística"
            End If
            Me.MousePointer = vbDefault
        Case "ID_Eliminar":
            Me.MousePointer = vbHourglass
            elimina txtsolicitud.Text
            Me.MousePointer = vbDefault
        Case "ID_Imprimir":
            imprimir
        Case "ID_Comparativo":
            ConsProd_Proveedores.Show 1
        Case "ID_Email":
            'jopcion = 1
            Email
        Case "ID_Consulta":
            If Val(ObtenerCampo("TB_CABSOLICITUD", "cs_estado", "cod_solicitud", txtsolicitud.Text, "T", cnn_dbbancos)) > 1 Then
                MsgBox "La solicitud ya ha sido aprobada anteriormente", vbInformation, "Sistema de Logística"
            Else
                enviarcorreoParaAprobacion

            End If
            'cons_solicitudes.Show 1
'        Case "ID_Lista":
'            If Grid.Dataset.State = dsEdit Then
'                'grid.Dataset.Post
'                sw_nuevo_item = True
'            End If
'
'            If sw_detalle = True Then
'                resp = MsgBox("La Solicitud no ha sido grabada. ¿Desea grabarla?", vbQuestion + vbYesNo, "Sistema de Logística")
'                If resp = vbYes Then
'                    If Trim$(txtsolicitante.Text) = Empty Then
'                        MsgBox "Debe Ingresar Solicitante", vbExclamation, "Sistema de Logística"
'                        txtsolicitante.SetFocus
'                        Exit Sub
'                    End If
'
'                'If pnlproveedor.Text = "" Then
'                '    MsgBox "Debe Seleccionar un Proveedor", vbInformation, "Sistema de Logística"
'                '    txtproveedor.SetFocus
'                '    Exit Sub
'                'End If
'
'                If optmoneda(0).Value = False And optmoneda(1).Value = False Then
'                    MsgBox "Debe Seleccionar Moneda de la Operación", vbExclamation, "Sistema de Logística"
'                    optmoneda(0).SetFocus
'                    Exit Sub
'                End If
'                'If Val(txttc.Text) = 0 Or (Not (IsNumeric(txttc.Text))) Then
'                '    MsgBox "Debe Ingresar Tipo de Cambio", vbExclamation, "Sistema de Logística"
'                '    txttc.SetFocus
'                '    Exit Sub
'                'End If
'                    grabar
'                    If wexito Then
'                        sw_detalle = False
'                        sw_cabecera = False
'
'                        SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
'                        SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
'                        Unload Me
'                    End If
'                Else
'                    sw_detalle = False
'                    sw_cabecera = False
'
'                    SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
'                    SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
'                    Unload Me
'                End If
'            Else
'                sw_detalle = False
'                sw_cabecera = False
'
'                SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
'                SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
'                Unload Me
'            End If
'        Case "ID_Salir":
'            Unload lista_solicitudes
'            Unload Me
        Case "ID_Lista"
            Unload Me
        Case Else
            MsgBox "Opciones Inhabilitadas Temporalmente.", vbInformation + vbOKOnly, App.ProductName
    End Select

End Sub

Private Sub SSActiveToolBars2_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
'    MsgBox "Opciones Inhabilitadas Temporalmente.", vbInformation + vbOKOnly, App.ProductName
    Select Case Tool.Id
        Case "ID_SolicitadoPor":
            acceso.Show vbModal
            If Len(Trim$(CodFirmaSolicitud(1))) > 0 Then
                SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = True
                SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = True
            End If
            'CodFirmaAprobacion(1) = CodFirmaSolicitud(1)
            txtsolicitante.Text = Trim$(CodFirmaSolicitud(1))
            sw_cabecera = True
            If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
                Grid.Dataset.Post
                sw_detalle = True
            End If
            If Len(txtsolicitante.Text & "") > 0 Then grabar
            
            sw_nuevo_documento = False
            'SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
            'SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
            'sw_cabecera = False
            'sw_detalle = False
            'imprimir
            'EMAIL
'        Case "ID_AprobadoPor":
'            aprobacion.Show vbModal
'            If Len(Trim$(CodFirmaAprobacion(1))) > 0 Then
'                SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = False
'                SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = False
'                SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Enabled = False
'                SSActiveToolBars1.Tools.ITEM("ID_Grabar").Enabled = False
'                If Grid.Dataset.State = dsEdit Or Grid.Dataset.State = dsInsert Then
'                    Grid.Dataset.Post
'                    sw_detalle = True
'                End If
'                cmbestado.ListIndex = 1
'                grabar
'                sw_nuevo_documento = False
'                SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
'                SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
'                sw_cabecera = False
'                sw_detalle = False
'            End If
    End Select

End Sub

Private Sub txtfecha_GotFocus()

    'txtfecha.FocusSelect = True
    
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{TAB}"
    End If

End Sub

Private Sub txtfecha_LostFocus()

    If IsDate(txtfecha.Value) = True Then
        If ctipoadm_bd = "M" Then
            rscambios.Open "SELECT * FROM CAMBIOS WHERE FECHA= '" & txtfecha.Value & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        Else
            rscambios.Open "SELECT * FROM CAMBIOS WHERE FECHA= CVDate( '" & txtfecha.Value & "' )", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        End If
        If Not rscambios.EOF Then
            txttc.Text = Format(Val("" & rscambios.Fields("cambio")), "0.000")
        End If
        rscambios.Close
    Else
        MsgBox "Fecha incorrecta. Verifique.", 48, "Atención"
        txtfecha.SetFocus
    End If

End Sub

Private Sub txtlugar_Change()
    
    If Len(Trim(txtlugar.Text)) > 0 And sw_cabecera = False Then
        sw_cabecera = True
    End If
    
End Sub

Private Sub txtlugar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtobservaciones.SetFocus
    End If

End Sub

Private Sub txtobservaciones_Change()
    
    If Len(Trim(txtobservaciones.Text)) > 0 And sw_cabecera = False Then
        sw_cabecera = True
    End If
    
End Sub

Private Sub txtobservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
End If
End Sub

Private Sub txtproveedor_Change()
pnlproveedor.Text = ""
sw_cabecera = True
End Sub

Private Sub txtproveedor_DblClick()

    txtproveedor_KeyDown 113, 0

End Sub

Private Sub txtproveedor_GotFocus()

    txtproveedor.SelStart = 0: txtproveedor.SelLength = Len(txtproveedor.Text)
    
End Sub

Private Sub txtproveedor_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        sw_ayuda = True
        'Screen.MousePointer = vbHourglass
        'hlp_proveedores.Show 1
        ayuda_proveedores_log.Show 1
        sw_ayuda = False
        If Len(Trim(wrucprov)) > 0 Then
            txtproveedor.Text = wrucprov
            pnlproveedor.Text = wnomprov
            
            txtproveedor_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtproveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cboprioridad.SetFocus
    Else
        KeyAscii = valida(3, KeyAscii)
    End If
    
End Sub

Private Sub txtproveedor_LostFocus()

    If sw_ayuda = False Then
        If Len(Trim(txtproveedor.Text)) > 0 Then
            If rst.State = adStateOpen Then rst.Close
            rst.Open "SELECT F2NOMPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & Trim(txtproveedor.Text) & "'", cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not rst.EOF Then
                pnlproveedor.Text = "" & rst.Fields("F2NOMPROV")
            Else
                MsgBox "El Proveedor no Existe. Verifique.", vbInformation, "Atención"
                pnlproveedor.Text = ""
                txtproveedor.Text = ""
                txtproveedor.SetFocus
            End If
            rst.Close
            Set rst = Nothing
        End If
    End If

End Sub

Private Sub txtsolicitante_Change()
On Error Resume Next
Dim rst As New ADODB.Recordset
    'If Not inicio Then swGrabacion = True
    If Len(Trim(txtsolicitante.Text)) > 0 Then
        If rst.State = adStateOpen Then rst.Close
        rst.Open "SELECT f2nomuser FROM ef2users WHERE ucase(f2coduser)=ucase('" & Trim(txtsolicitante.Text) & "')", cnn_dbbancos, 3, 1
        If Not rst.EOF Then
            LblSolicitante.Caption = UCase("" & rst.Fields("f2nomuser"))
        Else
            LblSolicitante.Caption = "NO EXISTE"
            'MsgBox "Código de solicitante no existe. Verifique.", vbInformation, "Atención"
            txtsolicitante.SetFocus
        End If
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub txtsolicitante_DblClick()
Call txtsolicitante_KeyDown(113, 0)
End Sub

Private Sub txtsolicitante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    End If
End Sub

Private Sub txtsolicitante_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodusuario = ""
        sw_ayuda = True
        ayuda_usuarios.Show 1
'        ayu_f_p_c.Show 1
        sw_ayuda = False
        If Len(wcodusuario) > 0 Then
            txtsolicitante.Text = wcodusuario
        End If
    End If
End Sub

Private Sub txtsolicitante_LostFocus()
txtsolicitante_Change
End Sub

Private Sub txtsolicitante_GotFocus()
    
    txtsolicitante.SelStart = 0: txtsolicitante.SelLength = Len(txtsolicitante.Text)

End Sub

Public Function NuevaSolicitud() As String
Dim Af As New ADOFunctions
Dim rst As New ADODB.Recordset
Dim numobra As String * 1
Dim numero_sol As String
    'Nuevo Procedimiento
    If Len(Trim(wnumpedido)) > 0 Then
        sql = "select TOP 1 cod_solicitud from tb_cabsolicitud where left(cod_solicitud,1) = '" & wnumpedido & "' order by val(right(cod_solicitud,6)) desc"
        Set rst = Af.OpenSQLForwardOnly(sql, cconex_dbbancos)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            solicitud = wnumpedido & Format$(Val(right(rst.Fields("cod_solicitud"), 11)) + 1, "00000000000")
        Else
            solicitud = wnumpedido & Format$(1, "00000000000")
        End If
        If wTipoReq = 1 Then
            NuevaSolicitud = OC & Format$(solicitud, "0000000000")
        Else
            NuevaSolicitud = OS & Format$(solicitud, "0000000000")
        End If
    Else
        If Len(Trim(wnumord)) > 0 Then
            numobra = wnumord
        Else
            numobra = "0"
        End If
        If wTipoReq = 1 Then
            sql = "select TOP 1 cod_solicitud from tb_cabsolicitud where LEFT(cod_solicitud,2) = 'RC' and mid(cod_solicitud,3,1) = '" & numobra & "' order by val(RIGHT(cod_solicitud,10)) desc"
            numero_sol = "RC"
        Else
            sql = "select TOP 1 cod_solicitud from tb_cabsolicitud where LEFT(cod_solicitud,2) = 'RS' and mid(cod_solicitud,3,1) = '" & numobra & "' order by val(RIGHT(cod_solicitud,10)) desc"
            numero_sol = "RS"
        End If
        Set rst = Af.OpenSQLForwardOnly(sql, cconex_dbbancos)
        Actualiza_Log sql, cconex_dbbancos
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            solicitud = Val(right(rst.Fields("cod_solicitud"), 9)) + 1
            NuevaSolicitud = numero_sol & numobra & Format$(solicitud, "000000000")
        Else
            solicitud = 1
            NuevaSolicitud = numero_sol & numobra & Format$(solicitud, "000000000")
        End If
        
    End If
    rst.Close
    Set rst = Nothing
    
End Function

Private Sub txttc_Change()
    
    If Len(Trim(txttc.Text)) > 0 And sw_cabecera = False Then
        sw_cabecera = True
    End If
    
End Sub

Private Sub txttc_GotFocus()
    
    txttc.SelStart = 0: txttc.SelLength = Len(txttc.Text)

End Sub

Private Sub txttc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        'If txtuupp.Visible = True Then
            txtuupp.SetFocus
        'Else
        '    grid.SetFocus
        'End If
    Else
        KeyAscii = valida(4, KeyAscii, txttc.Text)
    End If
    
End Sub

Private Sub txttc_LostFocus()
    
    txttc.Text = Format(txttc.Text, "0.000")
    
End Sub

Public Sub GRABAR_SOLICITUD()
On Error GoTo HNDERR
Dim aux             As ADODB.Recordset
Dim Fila            As Long
Dim nnumxusuario    As String
Dim nnumxobra       As String
Dim presug          As Double
Dim SUBTOTAL        As Double
Dim nitems          As Integer
Dim nfil            As Integer
Dim cnumvale        As String
Dim z As Integer

    If sw_nuevo_documento = True Then
        ctipo = "A"
        txtsolicitud.Text = NuevaSolicitud
        nnumxobra = Format(Obtiene_numxobra(Trim$(left(txtuupp.Text, 3))), "000000000")
        txtnumxobra.Text = Format(nnumxobra, "000000000")
        nnumxusuario = Format(Obtiene_numxusuario(Trim$(wcodsolicitante)), "000000000")
        txtnumxusuario.Text = Format(nnumxusuario, "000000000")
    Else
        ctipo = "M"
    End If
    
    cnumvale = txtsolicitud.Text
    
    If wTipoReq = 1 Then
        amovs_cab(16).campo = "CS_DOCUMENTO": amovs_cab(16).valor = "OC": amovs_cab(16).Tipo = "T"
        TOC = "OC"
    Else
        amovs_cab(16).campo = "CS_DOCUMENTO": amovs_cab(16).valor = "OS": amovs_cab(16).Tipo = "T"
        TOC = "OS"
    End If
    amovs_cab(0).campo = "COD_SOLICITUD": amovs_cab(0).valor = txtsolicitud.Text: amovs_cab(0).Tipo = "T"
    amovs_cab(1).campo = "NUMXUSUARIO": amovs_cab(1).valor = txtnumxusuario.Text: amovs_cab(1).Tipo = "T"
    amovs_cab(2).campo = "NUMXOBRA": amovs_cab(2).valor = txtnumxobra.Text: amovs_cab(2).Tipo = "T"
    amovs_cab(3).campo = "EMPRESA": amovs_cab(3).valor = wempresa: amovs_cab(3).Tipo = "T"
    amovs_cab(4).campo = "F4TIPCAM": amovs_cab(4).valor = txttc.Text: amovs_cab(4).Tipo = "N"
    amovs_cab(5).campo = "UUPP": amovs_cab(5).valor = "": amovs_cab(5).Tipo = "T"
    amovs_cab(6).campo = "CS_FECHA": amovs_cab(6).valor = Format(txtfecha.Value, "DD/MM/YYYY"): amovs_cab(6).Tipo = "F"
    amovs_cab(7).campo = "CS_CODCOSTO": amovs_cab(7).valor = txtuupp.Text: amovs_cab(7).Tipo = "T"
    amovs_cab(8).campo = "CS_DESCOSTO": amovs_cab(8).valor = txtdesuupp.Text: amovs_cab(8).Tipo = "T"
    amovs_cab(9).campo = "CS_CODSOLICITANTE": amovs_cab(9).valor = Trim(txtsolicitante.Text): amovs_cab(9).Tipo = "T"
    amovs_cab(10).campo = "CS_CLI": amovs_cab(10).valor = "": amovs_cab(10).Tipo = "T"
    amovs_cab(11).campo = "CS_TOTAL": amovs_cab(11).valor = Grid.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue: amovs_cab(11).Tipo = "N"
    amovs_cab(12).campo = "CS_PRIORIDAD": amovs_cab(12).valor = right(cboprioridad.Text, 1): amovs_cab(12).Tipo = "T"
    
    
    If optmoneda(0).Value Then
        xmoneda = "S"
    ElseIf optmoneda(1).Value Then
        xmoneda = "D"
    Else
        xmoneda = "E"
    End If
    amovs_cab(13).campo = "CS_MONEDA": amovs_cab(13).valor = xmoneda: amovs_cab(13).Tipo = "T"
    amovs_cab(14).campo = "CS_LUGENTR": amovs_cab(14).valor = Trim(txtlugar.Text): amovs_cab(14).Tipo = "T"
    amovs_cab(15).campo = "CS_OBSERVACIONES": amovs_cab(15).valor = Trim(txtobservaciones.Text): amovs_cab(15).Tipo = "T"
    amovs_cab(17).campo = "CS_USUARIO": amovs_cab(17).valor = wusuario: amovs_cab(17).Tipo = "T"
    'Si hay firma de solicitud
    If Len(Trim(CodFirmaSolicitud(1))) > 0 Then
        amovs_cab(18).campo = "CS_CODPREPARADOX": amovs_cab(18).valor = Trim(CodFirmaSolicitud(1)): amovs_cab(18).Tipo = "T"
        SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = True
        SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = True
    Else
        amovs_cab(18).campo = "CS_CODPREPARADOX": amovs_cab(18).valor = "": amovs_cab(18).Tipo = "T"
    End If
    'Si hay firma de aprobacion
    If Len(Trim(CodFirmaAprobacion(1))) > 0 Then
        amovs_cab(19).campo = "CS_APROBADOX": amovs_cab(19).valor = Trim(CodFirmaAprobacion(1)): amovs_cab(19).Tipo = "T"
        amovs_cab(23).campo = "VBJEFECC": amovs_cab(23).valor = 1: amovs_cab(23).Tipo = "N"
        SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = False
        SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = False
        SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Enabled = False
        SSActiveToolBars1.Tools.ITEM("ID_Grabar").Enabled = False
    Else
        amovs_cab(19).campo = "CS_APROBADOX": amovs_cab(19).valor = "": amovs_cab(19).Tipo = "T"
        amovs_cab(23).campo = "VBJEFECC": amovs_cab(23).valor = 0: amovs_cab(23).Tipo = "N"
    End If
    amovs_cab(20).campo = "CS_PROVEEDOR": amovs_cab(20).valor = txtproveedor.Text: amovs_cab(20).Tipo = "T"
    amovs_cab(21).campo = "CS_ESTADO": amovs_cab(21).valor = right(cmbestado.Text, 1): amovs_cab(21).Tipo = "T"
    amovs_cab(22).campo = "ANULADO": amovs_cab(22).valor = "N": amovs_cab(22).Tipo = "T"
    If ctipo = "A" Then
        amovs_cab(24).campo = "Fecha_graba": amovs_cab(24).valor = Format(Now(), "DD/MM/YYYY H:mm"): amovs_cab(24).Tipo = "F"
        amovs_cab(25).campo = "Fecha_modifica": amovs_cab(25).valor = Empty: amovs_cab(25).Tipo = "F"
        amovs_cab(26).campo = "cs_motivos": amovs_cab(26).valor = Empty: amovs_cab(26).Tipo = "T"
    Else
        amovs_cab(24).campo = "Fecha_graba": amovs_cab(24).valor = Lblcreacion.Caption: amovs_cab(24).Tipo = "F"
        amovs_cab(25).campo = "Fecha_modifica": amovs_cab(25).valor = Format(Now(), "DD/MM/YYYY H:mm"): amovs_cab(25).Tipo = "F"
        amovs_cab(26).campo = "cs_motivos": amovs_cab(26).valor = wusuario: amovs_cab(26).Tipo = "T"
    End If
    'Add
    'amovs_cab(27).campo = "Derivado": amovs_cab(27).valor = right(cmbDerivado(1).Text, 8): amovs_cab(27).Tipo = "T"
    
    'DETALLE
    amovs_det(0).campo = "COD_SOLICITUD": amovs_det(0).valor = "": amovs_det(0).Tipo = "T"
    amovs_det(1).campo = "ITEM": amovs_det(1).valor = "": amovs_det(1).Tipo = "N"
    amovs_det(2).campo = "DS_CANTIDAD": amovs_det(2).valor = "": amovs_det(2).Tipo = "N"
    amovs_det(3).campo = "CANDIS": amovs_det(3).valor = "": amovs_det(3).Tipo = "N"
    amovs_det(4).campo = "COD_PRODUCTO": amovs_det(4).valor = "": amovs_det(4).Tipo = "T"
    amovs_det(5).campo = "DS_UNIDMED": amovs_det(5).valor = "": amovs_det(5).Tipo = "T"
    amovs_det(6).campo = "DS_DESCRIPCION": amovs_det(6).valor = "": amovs_det(6).Tipo = "T"
    amovs_det(7).campo = "PRECIO": amovs_det(7).valor = "": amovs_det(7).Tipo = "N"
    amovs_det(8).campo = "PRESUG": amovs_det(8).valor = "": amovs_det(8).Tipo = "N"
    amovs_det(9).campo = "STOK": amovs_det(9).valor = "": amovs_det(9).Tipo = "N"
    amovs_det(10).campo = "CS_FENTREGA": amovs_det(10).valor = "": amovs_det(10).Tipo = "F"
    amovs_det(11).campo = "CS_SUBTOT": amovs_det(11).valor = "": amovs_det(11).Tipo = "N"
    amovs_det(12).campo = "CS_AFECTO": amovs_det(12).valor = "": amovs_det(12).Tipo = "T"
    amovs_det(13).campo = "F5MARCA": amovs_det(13).valor = "": amovs_det(13).Tipo = "T"
    amovs_det(14).campo = "F5CODFAB": amovs_det(14).valor = "": amovs_det(14).Tipo = "T"
    amovs_det(15).campo = "F5CODMARCA": amovs_det(15).valor = "": amovs_det(15).Tipo = "T"
    amovs_det(16).campo = "F5CODCOSTO": amovs_det(16).valor = "": amovs_det(16).Tipo = "T"
    
    amovs_det(17).campo = "f5CONigv": amovs_det(17).valor = "": amovs_det(17).Tipo = "N"
    amovs_det(18).campo = "f5SINigv": amovs_det(18).valor = "": amovs_det(18).Tipo = "N"
    amovs_det(19).campo = "F5AFECTO": amovs_det(19).valor = "": amovs_det(19).Tipo = "N"
    amovs_det(20).campo = "ruc_proveedor": amovs_det(20).valor = "": amovs_det(20).Tipo = "T"
    amovs_det(21).campo = "centro": amovs_det(21).valor = "": amovs_det(21).Tipo = "T"
    amovs_det(22).campo = "proveedor": amovs_det(22).valor = "": amovs_det(22).Tipo = "T"
    amovs_det(23).campo = "f3pordct": amovs_det(23).valor = "": amovs_det(23).Tipo = "N"
    amovs_det(24).campo = "cs_documento": amovs_det(24).valor = "": amovs_det(24).Tipo = "T"
    amovs_det(25).campo = "observa": amovs_det(25).valor = "": amovs_det(25).Tipo = "T"
    amovs_det(26).campo = "PARTIDA": amovs_det(26).valor = "": amovs_det(26).Tipo = "T"
    amovs_det(27).campo = "DES_PARTIDA": amovs_det(27).valor = "": amovs_det(27).Tipo = "T"
    
    '------------------- CALCULA NUMERO DE FILAS
    nitems = 0
    nitems = Val("" & Grid.Dataset.RecordCount)
    Grid.Dataset.Last
    
    If Len(Trim(Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value)) = 0 Then
        nitems = nitems - 1
    End If
    '---------------------------------------------
    ReDim Values(27, nitems)
    
    If Grid.Dataset.RecordCount > 0 Then
        nfil = 0
        For nfil = 0 To nitems - 1
            Grid.Dataset.RecNo = nfil + 1
            If Len(Trim(Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value & "")) > 0 Then
                Values(0, nfil) = txtsolicitud.Text
                Values(1, nfil) = Val(nfil + 1 & "")
                Values(2, nfil) = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                If ctipo = "A" Then
                    Values(3, nfil) = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                    Cantidad = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                Else
                    If IsNull(Grid.Columns.ColumnByFieldName("CANDIS").Value) Then
                        Values(3, nfil) = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                        Cantidad = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                    Else
                        If Val(Grid.Columns.ColumnByFieldName("CANDIS").Value & "") = Val(Grid.Columns.ColumnByFieldName("CANT_ANT").Value & "") Then
                            Values(3, nfil) = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                            Cantidad = Val(Grid.Columns.ColumnByFieldName("DS_CANTIDAD").Value & "")
                        Else
                            Values(3, nfil) = Val(Grid.Columns.ColumnByFieldName("CANDIS").Value & "")
                            Cantidad = Val(Grid.Columns.ColumnByFieldName("CANDIS").Value & "")
                        End If
                    End If
                End If
                Values(4, nfil) = Grid.Columns.ColumnByFieldName("COD_PRODUCTO").Value & ""
                'Values(5, nfil) = Grid.Columns.ColumnByFieldName("ds_unidmed").Value & ""
                Values(5, nfil) = Grid.Columns.ColumnByFieldName("DS_desUNIMED").Value & ""
                Values(6, nfil) = Grid.Columns.ColumnByFieldName("ds_descripcion").Value & ""
                Values(7, nfil) = Val(Grid.Columns.ColumnByFieldName("precio").Value & "")
                Values(8, nfil) = Val(Grid.Columns.ColumnByFieldName("presug").Value & "")
                Values(9, nfil) = Val(Grid.Columns.ColumnByFieldName("stok").Value & "")
                Values(10, nfil) = Format(Grid.Columns.ColumnByFieldName("cs_fentrega").Value & "", "DD/MM/YYYY")
                presug = Val(Grid.Columns.ColumnByFieldName("presug").Value & "")
                SUBTOTAL = Cantidad * presug
                Values(11, nfil) = SUBTOTAL
                Values(12, nfil) = "*"
                Values(13, nfil) = "" 'Grid.Columns.ColumnByFieldName("F5MARCA").Value & ""
                Values(14, nfil) = Grid.Columns.ColumnByFieldName("F5CODFAB").Value & ""
                Values(15, nfil) = Grid.Columns.ColumnByFieldName("F5CODMARCA").Value & ""
                Values(16, nfil) = txtuupp.Text
                
                Values(17, nfil) = Grid.Columns.ColumnByFieldName("ds_conIgv").Value & ""
                Values(18, nfil) = Grid.Columns.ColumnByFieldName("ds_sinIgv").Value & ""
                If Grid.Columns.ColumnByFieldName("afecto").Value = True Then
                    Values(19, nfil) = 1
                    Else
                    Values(19, nfil) = 0
                End If
                Values(20, nfil) = Grid.Columns.ColumnByFieldName("ruc_proveedor").Value & ""
                Values(21, nfil) = Grid.Columns.ColumnByFieldName("f5descosto").Value & ""
                Values(22, nfil) = Grid.Columns.ColumnByFieldName("NOMPROV").Value & ""
                Values(23, nfil) = Grid.Columns.ColumnByFieldName("DESCUENTO").Value & ""
                Values(24, nfil) = TOC & ""
                Values(25, nfil) = Grid.Columns.ColumnByFieldName("observa").Value & ""
                Values(26, nfil) = Grid.Columns.ColumnByFieldName("PARTIDA").Value & ""
                Values(27, nfil) = Grid.Columns.ColumnByFieldName("DES_PARTIDA").Value & ""
                'nfil = nfil + 1
            End If
        Next
    End If
    
    cvalores = "1111111111111111111111111111"
    
    '-------------------------------------------------------
    If ctipo = "A" Then     '--- Nuevo
        '------- GRABA CABECERA
        Dim x As Integer
        Do While x = 0
        sql = "select cod_solicitud from tb_cabsolicitud where cod_solicitud = '" & txtsolicitud.Text & "'"
        Set rst = Af.OpenSQLForwardOnly(sql, cconex_dbbancos)
            If rst.RecordCount > 0 Then
                solicitud = left(txtsolicitud.Text, 3) & Format$(Val(right(txtsolicitud.Text, 9)) + 1, "0000000000")
                amovs_cab(0).valor = solicitud
                txtsolicitud.Text = solicitud
                
                For z = 1 To nfil
                    Values(0, nfil) = txtsolicitud.Text
                Next
            Else
                x = 1
            End If
        Loop
        'GRABA_REGISTRO_logistica amovs_cab(), "TB_CABSOLICITUD", ctipo, 27, cnn_dbbancos, ""
        GRABA_REGISTRO_logistica amovs_cab(), "TB_CABSOLICITUD", ctipo, 26, cnn_dbbancos, ""
        
        If sw_GRABA_REGISTRO_logistica = True Then
            '------- GRABA DETALLE
            GRABA_REGISTRO_logistica_DET amovs_det(), "TB_DETSOLICITUD", ctipo, 27, cnn_dbbancos, "", Values(), nfil - 1, cvalores, "", ""
        End If
        'enviarcorreoAmail
    Else    '--- Modificación
        '------- GRABA CABECERA
        'GRABA_REGISTRO_logistica amovs_cab(), "TB_CABSOLICITUD", ctipo, 27, cnn_dbbancos, "cod_solicitud = '" & cnumvale & "'"
        GRABA_REGISTRO_logistica amovs_cab(), "TB_CABSOLICITUD", ctipo, 26, cnn_dbbancos, "cod_solicitud = '" & cnumvale & "'"
        'enviarcorreoAmail
        '-------------------------------------------------------
        '------- GRABA DETALLE
        If ctipoadm_bd = "M" Then
        
            sql = ("DELETE FROM TB_DETSOLICITUD WHERE cod_solicitud = '" & cnumvale & "'")
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        Else
            
            sql = ("DELETE * FROM TB_DETSOLICITUD WHERE cod_solicitud = '" & cnumvale & "'")
            cnn_dbbancos.Execute sql
            'AlmacenaQuery_sql sql, cnn_dbbancos
            Actualiza_Log sql, cnn_dbbancos.ConnectionString
        End If
        GRABA_REGISTRO_logistica_DET amovs_det(), "TB_DETSOLICITUD", "A", 27, cnn_dbbancos, "cod_solicitud = '" & cnumvale & "'", Values(), nfil - 1, cvalores, "", ""
    End If
    '-------------------------------------------------------
    Exit Sub

HNDERR:
    MsgBox Err.Description, vbCritical, "Atención"
    Resume Next
    
End Sub
''Sub enviarcorreoAmail()
''        sql = "select F2NOMUSER,MAIL,PASSWORD,F2enviaMail from EF2USERS where F2CODUSER ='" & wusuario & "'"
''        If Rs.State = adStateOpen Then Rs.Close
''        Rs.Open sql, cnn_dbbancos, 3, 1
''        If Rs.RecordCount > 0 Then
''        nombre = Rs.Fields(0) & ""
''        mail = Rs.Fields(1) & ""
''        pass = Rs.Fields(2) & ""
''        perms = Rs.Fields("F2enviaMail") & ""
''        End If
''        If perms = "*" Then
''            If mail <> "" Then
''                destino = CargaDestinatarios & ", " & mail
''            Else
''                destino = CargaDestinatarios
''            End If
''
''            'enviacorreo Me.InternetMail1, "mail.neodeter.com", "nombre ", "Sistemas_2@neodeter.com ", "neodeter017", "jc_gilardi@yahoo.com,controlplus.peru@gmail.com,jk_20063@hotmail.com", "Requerimiento Nª: " & txtsolicitud.Text, "Requerimiento Nª: " & txtsolicitud.Text 'prueba
''            asunto = wnomcia & " - " & wanno & ", Requerimiento Nº: " & txtsolicitud.Text & ", Prioridad:" & cboprioridad.Text
''            cuerpo = wnomcia & " - " & wanno & ", Requerimiento Nª: " & txtsolicitud.Text & ", Observacion :" & txtobservaciones.Text & ", Prioridad:" & cboprioridad.Text
''
''            MsgBox ("Su requerimiento está siendo enviado."), vbInformation, "CONTROLPLUS!"
''            Me.MousePointer = vbhourglass
''            enviacorreoGmail Me.InternetMail1, "pop.gmail.com", CStr(nombre), "controlplus.peru@gmail.com ", "infoplus12345", CStr(destino), CStr(asunto), CStr(cuerpo)
''            Me.MousePointer = vbdefault
'''            If Not enviacorreoGmail(Me.InternetMail1, "pop.gmail.com", CStr(nombre), "controlplus.peru@gmail.com ", "infoplus12345", CStr(destino), CStr(asunto), CStr(cuerpo)) Then
'''              If MsgBox("No se ha podido envía por mail. ¿Desea reintentar?", vbYesNo, "CONTROLPLUS!") = vbYes Then
'''                enviacorreoGmailMe.InternetMail1 , "pop.gmail.com", CStr(nombre), "controlplus.peru@gmail.com ", "infoplus12345", CStr(destino), CStr(asunto), CStr(cuerpo)
'''              End If
'''            End If
''        End If
''End Sub

Sub enviarcorreoParaAprobacion()

cuerpo = cargacuerpo(txtsolicitud.Text)

            Dim codUser As String
            nombre = "CONTROLPLUS"
            codUser = ObtenerCampo("CENTROS", "F3RESPONSABLE", "F3COSTO", txtuupp.Text, "T", cnn_dbbancos)
            destino = ObtenerCampo("EF2USERS", "MAIL", "F2CODUSER", codUser, "T", cnn_dbbancos)
            If destino <> "" Then
                destino = destino & ";psalas@betania.com.pe;pflores@bkasociados.com.pe;responder.britania@gmail.com;betania_bk@outlook.com.pe"
                asunto = "OK - " & txtsolicitud.Text & "-" & wempresa
                Me.MousePointer = vbHourglass
    '            enviacorreoGmail Me.InternetMail1, "pop.gmail.com", CStr(nombre), "responder.britania@gmail.com ", "infoplus1234", CStr(destino), CStr(asunto), CStr(cuerpo)
    '            enviacorreohotmail Me.InternetMail1, "pop.live.com", CStr(nombre), "betania_bk@outlook.com.pe", "infoplus1234", CStr(destino), CStr(asunto), CStr(cuerpo)
                'InternetMail1.ComposeMessage(
                'enviacorreoPOP Me.InternetMail1, "mail.betania.com.pe", CStr(nombre), "aprobaciones@betania.com.pe", "4pr0.b3t4", CStr(destino), CStr(asunto), CStr(cuerpo)
                
                EnviarCorreoOcx.Show 1
                If Sw_Act = True Then
                    sql = "update tb_cabsolicitud set cs_estado='0' where cod_solicitud='" & txtsolicitud.Text & "'"
                End If
                cnn_dbbancos.Execute sql
                'AlmacenaQuery_sql sql, cnn_dbbancos
                Actualiza_Log sql, cnn_dbbancos.ConnectionString
                Me.MousePointer = vbDefault
            End If
'            If Not enviacorreoGmail(Me.InternetMail1, "pop.gmail.com", CStr(nombre), "controlplus.peru@gmail.com ", "infoplus12345", CStr(destino), CStr(asunto), CStr(cuerpo)) Then
'              If MsgBox("No se ha podido envía por mail. ¿Desea reintentar?", vbYesNo, "CONTROLPLUS!") = vbYes Then
'                enviacorreoGmailMe.InternetMail1 , "pop.gmail.com", CStr(nombre), "controlplus.peru@gmail.com ", "infoplus12345", CStr(destino), CStr(asunto), CStr(cuerpo)
'              End If
'            End If
 '       End If
End Sub
Function cargacuerpo(ByVal n_solicitud As String) As String
Dim cu, Cab As String
Dim cod As String, dis As String, Cant As String, midi  As String
Dim Line As String
Dim CABD As String
             
sql = "select item,cod_producto,ds_descripcion,ds_unidmed,ds_cantidad,stok,precio,presug,cs_fentrega,candis,f5marca,f5codfab,F5CODMARCA from tb_detsolicitud where cod_solicitud='" & n_solicitud & "'"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1
                
     d = 1
    Do While Not Rs.EOF
        cod = Trim(Rs.Fields("cod_producto") & "")
        dis = Trim(Rs.Fields("ds_descripcion") & "")
        Cant = Rs.Fields("ds_cantidad") & ""
        midi = Trim(ObtenerCampo("ef7medidas", "F7NOMMED", "f7codmed", Rs!ds_unidmed & "", "T", cnn_dbbancos))
        cu = cod & Space(10) _
        & Trim(dis) & Space(50 - Len(Trim(left(dis, 49)))) _
        & Cant & Space(30 - Len(Cant)) _
        & midi & Space(30 - Len(midi)) & vbCrLf
    If d = 1 Then
        Cab = "CODIGO" & Space(14) _
        & "DISCRIPCION" & Space(50 - Len("DISCRIPCION")) _
        & "CANTIDAD" & Space(30 - Len("CANTICAD")) _
        & "U.MEDIDA" & Space(30 - Len("U.MIDIDA")) & vbCrLf
        Line = ""
        For x = 1 To 109
            Line = Line & "="
        Next
        CABD = "Solicitante       : " & LblSolicitante.Caption & Space(30 - Len(LblSolicitante.Caption)) & "Fecha        : " & txtfecha.Value & vbCrLf
        CABD = CABD & "Prioridad         : " & left(cboprioridad.Text, 10) & Space(30 - Len(left(cboprioridad.Text, 10))) & "Nº solicitud : " & txtsolicitud.Text & vbCrLf
        CABD = CABD & "Centro de costos  : " & txtdesuupp.Text & Space(30 - Len(txtdesuupp.Text)) & "Observación  : " & txtobservaciones.Text & vbCrLf
        cargacuerpo = CABD & Line & vbCrLf & Cab & Line & vbCrLf & cu
    Else
        cargacuerpo = cargacuerpo & cu
    End If
        d = d + 1
        Rs.MoveNext
    Loop
End Function

Private Function CargaDestinatarios() As String
Dim F As Integer
CargaDestinatarios = ""
sql = "SELECT EF2TAREAS.F2CODTAREA, EF2USERS.F2NOMUSER, EF2USERS.MAIL FROM EF2TAREAS INNER JOIN (EF2TAREAUSERS INNER JOIN EF2USERS ON EF2TAREAUSERS.F2CODUSER = EF2USERS.F2CODUSER) ON EF2TAREAS.F2CODTAREA = EF2TAREAUSERS.F2CODTAREA WHERE (((EF2TAREAS.F2CODTAREA)='0005'))"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1

If Rs.RecordCount > 0 Then
F = 1
    Do While Not Rs.EOF
        If F = 1 Then
            If Not IsNull(Rs!Mail) Then
                CargaDestinatarios = Rs!Mail
            End If
        Else
            If Not IsNull(Rs!Mail) Then
                CargaDestinatarios = CargaDestinatarios & ", " & Rs!Mail
            End If
        End If
        F = F + 1
        Rs.MoveNext
    Loop
    If right(Trim(CargaDestinatarios), 1) = "," Then
    CargaDestinatarios = left(Trim(CargaDestinatarios), Len(Trim(CargaDestinatarios)) - 1)
    End If
    If Mid(Trim(CargaDestinatarios), 1, 1) = "," Then
    CargaDestinatarios = Mid(Trim(CargaDestinatarios), 2, Len(Trim(CargaDestinatarios)))
    End If
    CargaDestinatarios = CargaDestinatarios
Else
    CargaDestinatarios = ""
End If
End Function

Public Function CalculaTotal()

    CalculaTotal = Grid.Columns(9).SummaryFooterValue
    
End Function

Public Sub BUSCA_SOLICITUD()
Dim cnn             As ADODB.Connection
Dim aux             As ADODB.Recordset
Dim sw_nuevo_temp   As Boolean
Dim Precio          As Double
Dim presug          As Double
Dim TOTAL           As Double
Dim Tipo            As String

    Set cnn = New ADODB.Connection
    Set aux = New ADODB.Recordset
    Tipo = lista_solicitudes.dxDBGrid1.Columns(0).Value
    solicitud = lista_solicitudes.dxDBGrid1.Columns(1).Value
    txtsolicitud.Text = solicitud

    'Recupera datos de tb_cabsolicitud - Archivo de cabecera
'    sql = "select * from tb_cabsolicitud where " _
'    & "cod_solicitud='" & Format(solicitud, "000000000000") & "' and cs_documento='" & Trim(TIPO) & "'"   '
    
    Rem SK ADD:
    sql = vbNullString
    sql = sql & "SELECT "
    sql = sql & "* "
    sql = sql & "FROM "
    sql = sql & "TB_CABSOLICITUD "
    sql = sql & "WHERE "
    sql = sql & "COD_SOLICITUD = '" & Trim(solicitud) & "' AND "
    sql = sql & "CS_DOCUMENTO = '" & Trim(Tipo) & "'"
    
    If sw_nuevo_documento = False Then
        DELETEREC_LOG cnomtabla, cnn_form
        AdicionaItem
        sw_nuevo_documento = True
    End If
    
    Grid.Dataset.ADODataset.ConnectionString = cnn_form
    Grid.Dataset.Active = True
    
    Grid.Dataset.Close
    Grid.Dataset.Open
    
    Grid.OptionEnabled = False
    Grid.Dataset.DisableControls
        
    If RsC.State = adStateOpen Then RsC.Close
    RsC.Open sql, cnn_dbbancos, 3, 1
    If RsC.RecordCount > 0 Then
        With RsC
            RsC.MoveFirst
            txtfecha.Value = .Fields("cs_fecha")
            txtuupp.Text = .Fields("cs_codcosto") & ""
            txtdesuupp.Text = .Fields("cs_descosto") & ""
            
            SeleccionaEnComboRight .Fields("cs_prioridad") & "", cboprioridad
            If left$(.Fields("cs_moneda"), 1) = "S" Then
                optmoneda(0).Value = True
            ElseIf left$(.Fields("cs_moneda"), 1) = "D" Then
                optmoneda(1).Value = True
            Else
                'optmoneda(2).Value = True
            End If
            txttc.Text = Format$(.Fields("f4tipcam"), "0.000")
            txtlugar.Text = .Fields("cs_lugentr") & ""
            txtobservaciones.Text = .Fields("cs_observaciones") & ""
            txtnumxobra.Text = .Fields("numxobra") & ""
            txtnumxusuario.Text = .Fields("numxusuario") & ""
            cmbestado.ListIndex = Val("" & .Fields("cs_estado").Value) - 1
            'Obtiene Nombre del solicitante
            If cnn.State = adStateOpen Then cnn.Close
            'cnn.Open "provider=microsoft.jet.oledb.4.0;data source=" & wrutabancos & "\DB_BANCOS.mdb"
            'sql = "select f2nomuser From ef2users where f2coduser='" & wusuario & "'"
            'aux.Open sql, cnn, adOpenDynamic, adLockOptimistic
            txtsolicitante.Text = .Fields("cs_codsolicitante")
            Lblcreacion.Caption = .Fields("Fecha_graba") & ""
            'cmbDerivado(1).Text = .Fields("Derivado") & ""
               
            If Not (IsNull(.Fields("cs_codpreparadox"))) Then
                If Len(Trim$(.Fields("cs_codpreparadox"))) > 0 Then
                    CodFirmaSolicitud(1) = Trim$(.Fields("cs_codpreparadox"))
                    SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = True
                    SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = True
                Else
                    CodFirmaSolicitud(1) = ""
                    SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = True
                    SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = False
                End If
            End If
                
            If Not (IsNull(.Fields("cs_orden"))) Then
                If Len(Trim$(.Fields("cs_orden"))) > 0 Then
                    'CodFirmaAprobacion(1) = Trim$(.Fields("cs_aprobadox"))
                    SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = False
                    SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = False
                    SSActiveToolBars1.Tools.ITEM("ID_Grabar").Enabled = False
                    SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Enabled = False
                Else
                    CodFirmaAprobacion(1) = ""
                End If
            End If
            
    '        aux.Close
            
            txtproveedor.Text = Trim(.Fields("CS_PROVEEDOR") & "")
            pnlproveedor.Text = ""
            If Len(Trim(txtproveedor.Text)) > 0 Then
                If aux.State = adStateOpen Then aux.Close
                aux.Open "SELECT F2NOMPROV FROM EF2PROVEEDORES WHERE F2NEWRUC='" & Trim(txtproveedor.Text) & "'", cnn_dbbancos, adOpenStatic, adLockOptimistic
                If Not aux.EOF Then
                    pnlproveedor.Text = "" & aux.Fields("F2NOMPROV")
                End If
                aux.Close
                Set aux = Nothing
            End If
            
           
    
            chkcerrar.Caption = "Solicitud Cerrada"
            If "" & .Fields("F4CERRADO") = "S" Then
                wcerrado = "S"
                chkcerrar.Value = True
                activar False
            Else
                chkcerrar.Value = False
                wcerrado = "N"
                activar True
            End If
            RsC.Close
            
            'Recupera datos de tb_detsolicitud - Archivo de detalle
            If ctipoadm_bd = "M" Then
                sql = "select item,cod_producto,ds_descripcion,ds_unidmed,ds_cantidad,stok,precio,f3pordct, " _
                & "presug,cs_fentrega,candis,f5marca,f5codfab,F5CODMARCA,f5CONigv,f5SINigv,F5AFECTO from tb_detsolicitud where cod_solicitud='" & solicitud & "' order " _
                & "by item"
            Else
                sql = "select item,cod_producto,ds_descripcion,ds_unidmed,ds_cantidad,stok,precio,f3pordct, " _
                & "presug,cs_fentrega,candis,f5marca,f5codfab,F5CODMARCA,f5CONigv,f5SINigv,F5AFECTO,ruc_proveedor,F5CODCOSTO,centro, proveedor,observa,PARTIDA,DES_PARTIDA from tb_detsolicitud " _
                & "where cod_solicitud='" & solicitud & "' AND cs_documento='" & Tipo & "' order by val(item)"
            End If
            sw_nuevo_temp = False
            
            If rst.State = adStateOpen Then rst.Close
            rst.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not (rst.EOF) Then
                sw_nuevo_item = True
                rst.MoveFirst
                Do While Not (rst.EOF)
                    If sw_nuevo_temp = False Then
                        If sw_nuevo_documento = True Then
                            Grid.Dataset.Edit
                        Else
                            Grid.Dataset.Append
                        End If
                        sw_nuevo_temp = True
                    Else
                        Grid.Dataset.Append
                    End If
                    Grid.Dataset.FieldValues("ITEM") = rst!ITEM
                    Grid.Dataset.FieldValues("COD_PRODUCTO") = rst!COD_PRODUCTO
                    Grid.Dataset.FieldValues("F5CODMARCA") = "" & rst!F5CODMARCA
                    Grid.Dataset.FieldValues("F5MARCA") = "" & rst!F5MARCA
                    Grid.Dataset.FieldValues("F5CODFAB") = "" & rst!f5codfab
                    Grid.Dataset.FieldValues("DS_DESCRIPCION") = rst!ds_descripcion
                    Grid.Dataset.FieldValues("DS_UNIDMED") = rst!ds_unidmed
                    Grid.Dataset.FieldValues("DS_DESUNIMED") = rst!ds_unidmed 'ModUtilitario.ObtenerCampoV2(cnn_dbbancos, "F7SIGMED", "EF7MEDIDAS", "F7CODMED", Trim(rst!ds_unidmed & ""), "T") 'ObtenerCampo("ef7medidas", "f7sigmed", "f7codmed", rst!ds_unidmed & "", "T", cnn_dbbancos)
                    Cantidad = IIf(IsNull(rst!ds_cantidad), Null, Format(rst!ds_cantidad, "###,##0.00"))
                    Grid.Dataset.FieldValues("DS_CANTIDAD") = rst!ds_cantidad
                    Grid.Dataset.FieldValues("STOK") = IIf(IsNull(rst!STOK), Null, rst!STOK)
                    Precio = IIf(IsNull(Val("" & rst!Precio)), Null, Format(Val("" & rst!Precio), "###,##0.00"))
                    Grid.Dataset.FieldValues("PRECIO") = Precio
                    presug = IIf(IsNull(Val("" & rst!presug)), Null, Format(Val("" & rst!presug), "###,##0.00"))
                    Grid.Dataset.FieldValues("PRESUG") = presug
                    TOTAL = IIf(IsNull(Cantidad), 0, Cantidad) * IIf(IsNull(Precio), 0, Precio)
                    Grid.Dataset.FieldValues("TOTAL") = TOTAL
                    Grid.Dataset.FieldValues("CS_FENTREGA") = Format(rst!cs_fentrega, "dd/mm/yyyy")
                    Grid.Dataset.FieldValues("CANDIS") = Val("" & rst!candis)
                    Grid.Dataset.FieldValues("CANT_ANT") = Val("" & rst!ds_cantidad)
                    
                    Grid.Dataset.FieldValues("ds_conIgv") = Val("" & rst!f5CONigv)
                    Grid.Dataset.FieldValues("ds_sinIgv") = Val("" & rst!f5SINigv)
                    Grid.Dataset.FieldValues("afecto") = rst!F5AFECTO
                    Grid.Dataset.FieldValues("F5CODCOSTO") = rst!F5CODCOSTO
                    Grid.Dataset.FieldValues("f5descosto") = rst!centro
                    Grid.Dataset.FieldValues("ruc_proveedor") = rst!ruc_proveedor
                    Grid.Dataset.FieldValues("NOMPROV") = rst!proveedor
                    Grid.Dataset.FieldValues("DESCUENTO") = rst!F3PORDCT
                    Grid.Dataset.FieldValues("observa") = rst!OBSERVA
                    Grid.Dataset.FieldValues("PARTIDA") = rst!Partida
                    Grid.Dataset.FieldValues("DES_PARTIDA") = rst!DES_PARTIDA
                    rst.MoveNext
                Loop
                Grid.Dataset.Post
                sw_nuevo_item = False
            End If
            rst.Close
        End With
    End If
    Grid.Dataset.EnableControls
    Grid.Dataset.Open
    Grid.OptionEnabled = True
End Sub
Public Sub CargarUsuario()
Dim rst As New ADODB.Recordset
    sql = "select f2coduser,f2nomuser From ef2users where ucase(f2coduser)=ucase('" & wusuario & "')"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, 3, 1
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        txtsolicitante.Text = Trim$(rst.Fields("f2coduser") & "")
        wnomsolicitante = UCase(Trim$(rst.Fields("f2nomuser") & ""))
        wcodsolicitante = UCase(Trim$(rst.Fields("f2coduser") & ""))
        lblusuario.Caption = wnomsolicitante
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Email()
Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String


If MsgBox("Los Datos serán enviados via E-mail para su aprobación", vbOKCancel + vbQuestion, "InfoPlus") = vbOK Then
GeneraPDF

'sTo = wemailsol
'sCC = wemailccsol
'sBCC = ""
'sSubject = wasuntosol & " Solicitud de Compra : " & txtsolicitud.Text
'sBody = "Observaciones : " & Trim(txtobservaciones.Text)
'
'ret = Shell("Start.exe " _
'    & "mailto:" & """" & sTo & """" _
'    & "?Subject=" & """" & sSubject & """" _
'    & "&cc=" & """" & sCC & """" _
'    & "&bcc=" & """" & sBCC & """" _
'    & "&Body=" & """" & sBody & """" _
'    & "&File=" & """" & "c:\autoexec.bat" & """" _
'    , 0)

'ret = Shell("Start.exe " _
    & "mailto:" & """" & sTo & """" _
    & "?Subject=" & """" & sSubject & """" _
    & "&cc=" & """" & sCC & """" _
    & "&bcc=" & """" & sBCC & """" _
    & "&Body=" & """" & sBody & """" _
    & "&File=" & """" & "c:\autoexec.bat" & """" _
    , vbHide)
End If
End Sub

Private Sub grabar()

'If rst.State = adStateOpen Then rst.Close
'sql = "select * from solicitud"
'rst.Open sql, cnn_form, adOpenStatic, adLockOptimistic
'wexito = False
'If Not rst.EOF Then
'    wcont1 = 0
'    Do While Not rst.EOF
'        If (Len("" & rst("DS_DESCRIPCION")) > 0) Then
'            wcont1 = wcont1 + 1
'            wexito = True
'        End If
'        rst.MoveNext
'    Loop
'End If
'rst.Close

wcont1 = 0
For I = 1 To Grid.Dataset.RecordCount
    Grid.Dataset.RecNo = I
    If (Len("" & Grid.Columns.ColumnByFieldName("DS_DESCRIPCION").Value) > 0) Then
        wcont1 = wcont1 + 1
        wexito = True
    End If
Next


If Not wexito Then
    MsgBox "Debe Seleccionar y/o Ingresar lo(s) Producto(s) de la Solicitud", vbInformation, "Sistema de Logística"
    Exit Sub
End If

Me.MousePointer = vbHourglass

GRABAR_SOLICITUD

sw_nuevo_documento = False
sw_cabecera = False
sw_detalle = False

SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = True
SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = True
SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Enabled = True
            
Me.MousePointer = vbDefault
MsgBox "Solicitud actualizada", vbInformation, "Sistema de Logistica"
'jopcion = 2
'EMAIL
        
End Sub

Private Sub GeneraPDF()
On Error Resume Next
Dim oEXL As ActiveReportsExcelExport.ARExportExcel
Dim oPDF As ActiveReportsPDFExport.ARExportPDF
wexcel = 0
Me.MousePointer = vbHourglass
With acr_solicitud
    .datos.ConnectionString = cnn_dbbancos
    sql = "select b.item,b.ds_cantidad,b.ds_unidmed,b.ds_descripcion,b.stok, " _
    & "b.presug,a.cs_descosto,a.cs_fecha,a.cs_lugentr,b.F5CODFAB, " _
    & "a.numxusuario,a.numxobra,b.cs_fentrega,a.cs_codcosto, format(ds_cantidad,'#,###,###0.00')*format(precio,'#,###,###0.00') as total " _
    & "from tb_cabsolicitud as a, tb_detsolicitud as b where " _
    & "a.cod_solicitud=b.cod_solicitud and a.cod_solicitud='" & txtsolicitud.Text & "' order by val(b.item)"
    .datos.Source = sql
    
    If Trim(wtiposalida) = "*" Then  '----- LA EMPRESA ES CONSTRUCTORA (AIC)
        .lblTitulo.Caption = "ORDEN DE REQUERIMIENTO Nº "
    Else
        .lblTitulo.Caption = "ORDEN DE REQUERIMIENTO Nº "
    End If
    
    .lblempresa.Caption = wnomcia
    .fldfecha.Text = txtfecha.Value
    .solicitante.Text = ObtenerCampo("ef2users", "f2nomuser", "f2coduser", txtsolicitante.Text, "T", cnn_dbbancos)
    '.solicitudx.Text = txtsolicitud.Text
    .observaciones.Text = txtobservaciones.Text
    '.flduupp.Text = txtuupp.Text
    '.flddesuupp.Text = txtdesuupp.Text
    .Caption = "Solicitud de Suministros"
    .left = Screen.Width * 2
    .top = Screen.Height * 2
    .WindowState = 0
    sw_nuevo_doc = True
    .Show 1
    
    
End With

Unload acr_solicitud
sql = "SELECT * FROM EF2USERS_DER WHERE CODIGO='017'"
If Rs.State = 1 Then Rs.Close
Rs.Open sql, cnn_dbbancos, 3, 1
wDestinatarios = ""
wSubject = "Solicitud de Suministros - " & txtsolicitud.Text & " - " & txtsolicitante.Text
I = 1
Do While Not Rs.EOF
    If I = 1 Then
        wDestinatarios = ObtenerCampo("ef2users", "mail", "f2coduser", Rs!F2CODUSER & "", "T", cnn_dbbancos)
    Else
        wDestinatarios = wDestinatarios & "; " & ObtenerCampo("ef2users", "mail", "f2coduser", Rs!F2CODUSER & "", "T", cnn_dbbancos)
    End If
    I = I + 1
    Rs.MoveNext
Loop
wDestinatarioOculto = "sistema.solicitudes@gmail.com"
Load MailSend
'    .left = Screen.Width * 2
'    .top = Screen.Height * 2
'    .Show 1
    

Unload MailSend
Me.MousePointer = vbDefault
MsgBox "E-Mail Enviado"
End Sub


Private Sub imprimir()
Dim oEXL As ActiveReportsExcelExport.ARExportExcel
    
Me.MousePointer = vbHourglass
With acr_solicitud
    .datos.ConnectionString = cnn_dbbancos
    sql = "select b.item,b.ds_cantidad,b.ds_unidmed,b.ds_descripcion,b.stok, " _
    & "b.presug,a.cs_descosto,a.cs_fecha,a.cs_lugentr,b.F5CODFAB, " _
    & "a.numxusuario,a.numxobra,b.cs_fentrega,a.cs_codcosto, format(ds_cantidad,'#,###,###0.00')*format(precio,'#,###,###0.00') as total,b.observa " _
    & "from tb_cabsolicitud as a, tb_detsolicitud as b where " _
    & "a.cod_solicitud=b.cod_solicitud and a.cod_solicitud='" & txtsolicitud.Text & "' order by val(b.item)"
    
    sql = "SELECT b.f5CONigv,b.f5SINigv, b.item, b.ds_cantidad, b.ds_unidmed, b.ds_descripcion, b.stok, b.presug, a.cs_descosto, a.cs_fecha,b.cod_producto, "
    sql = sql & "a.cs_LugEntr, b.f5codfab, a.NUMXUSUARIO, a.NUMXOBRA, b.cs_fentrega, a.cs_codcosto, "
    sql = sql & "Format(ds_cantidad,'#,##0.00')*Format(f5SINigv,'#,##0.00') AS totalsin, "
    sql = sql & "Format(ds_cantidad,'#,##0.00')*Format(f5CONigv,'#,##0.00') AS totalcon , b.F5CODCOSTO, b.ruc_proveedor,b.observa "
    sql = sql & "FROM tb_cabsolicitud AS a LEFT JOIN (tb_detsolicitud AS b LEFT JOIN EF7MEDIDAS AS m ON b.ds_unidmed = m.F7CODMED) "
    sql = sql & "ON a.cod_solicitud = b.cod_solicitud "
    sql = sql & "where a.cod_solicitud='" & right(txtsolicitud.Text, 12) & "'"
    sql = sql & "ORDER BY Val(b.item) "
    ' [ds_cantidad]*[f5SINigv] AS totalsin, [ds_cantidad]*[f5CONigv] AS totalcon
    
    .datos.Source = sql
    
    If Trim(wtiposalida) = "*" Then  '----- LA EMPRESA ES CONSTRUCTORA (AIC)
        .lblTitulo.Caption = "ORDEN DE REQUERIMIENTO Nº " & txtsolicitud.Text
    Else
        .lblTitulo.Caption = "ORDEN DE REQUERIMIENTO Nº " & txtsolicitud.Text
    End If
    .fldMoneda.Text = IIf(optmoneda(0).Value = True, "NUEVOS SOLES", "DOLARES")
    .Lblsigno.Caption = IIf(optmoneda(0).Value = True, "S/", "US$")
    '.lblempresa.Caption = wnomcia
    '.LblDireccion.Caption = wdireccion
    '.LblDistrito.Caption = wDistrito
    '.LblTelefono.Caption = "Teléfono: " & wtelefono
    '.LblRuc.Caption = "R.U.C. " & wrucempresa
    
'    If FileExist(App.Path & "\" & wrucempresa & ".jpg") = True Then
        '.'lblempresa.Visible = False
        '.ImageLogo.Visible = True
        '.ImageLogo.Picture = LoadPicture(App.Path & "\" & wrucempresa & ".jpg")
        'Set cImgInfo = New cImageInfo
'        With cImgInfo
'            .ReadImageInfo (App.Path & "\" & wrucempresa & ".jpg")
'
'            acr_solicitud.ImageLogo.Height = 850
'            acr_solicitud.ImageLogo.Width = 850 * .Width / .Height
'        End With
'        .ImageLogo.Top = 0
'        .ImageLogo.Left = 0
'    Else
'        '.lblempresa.Visible = True
'        .ImageLogo.Visible = False
'        '.Caption = wnomcia  '"OCCIDENTAL BUSINESS CORPORATION S.A.C."
'
'    End If
        
    .fldfecha.Text = "Fecha: " & txtfecha.Value
    .solicitante.Text = ObtenerCampo("ef2users", "f2nomuser", "f2coduser", txtsolicitante.Text, "T", cnn_dbbancos)
   ' .solicitudx.Text = txtsolicitud.Text
    .observaciones.Text = txtobservaciones.Text
'    .FldCodCentro.Text = Left(txtuupp.Text, 3)
'    .FldNomCentro.Text = ObtenerCampo("centros", "f3descrip", "f3costo", .FldCodCentro.Text, "T", cnn_dbbancos)
'    .FldCodSubCentro.Text = txtuupp.Text
'    .FldNomSubCentro.Text = ObtenerCampo("centros", "f3descrip", "f3costo", .FldCodSubCentro.Text, "T", cnn_dbbancos)
    If wexcel = 1 Then
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = "c:\mis documentos\sol" & Format(txtsolicitud.Text, "000000000000") & ".xls"
        oEXL.Export acr_solicitud.Pages
        .Run
        wexcel = 0
    Else
        .Caption = "Solicitud de Suministros"
        .Show vbModal
    End If
End With
Me.MousePointer = vbDefault
End Sub

Private Function Obtiene_numxusuario(pUsuario As String)
Dim nnumxusuario    As Double
Dim sql             As String

    sql = "SELECT COUNT(COD_SOLICITUD) AS NSOL FROM TB_CABSOLICITUD WHERE UCASE(CS_CODSOLICITANTE)=UCASE('" & pUsuario & "')"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos, 3, 1
    If Not (rst.EOF) Then
        nnumxusuario = Val("" & rst.Fields("NSOL")) + 1
    End If
    rst.Close
    Obtiene_numxusuario = nnumxusuario

End Function

Private Function Obtiene_numxobra(pobra As String)
Dim nnumxobra   As Double
Dim sql         As String

    sql = "SELECT COUNT(COD_SOLICITUD) AS NSOL FROM TB_CABSOLICITUD WHERE CS_CODCOSTO='" & pobra & "'"
    If rst.State = adStateOpen Then rst.Close
    rst.Open sql, cnn_dbbancos
    If Not (rst.EOF) Then
        nnumxobra = Val("" & rst.Fields("NSOL")) + 1
    End If
    rst.Close
    Obtiene_numxobra = nnumxobra

End Function

Private Sub nuevo()
    
    Call inicio
    txtsolicitante.Text = ""
    LblSolicitante.Caption = ""
    lblusuario.Caption = ""
    txtsolicitud.Text = ""
    txttc.Text = "0.000"
    txtnumxobra.Text = ""
    txtnumxusuario.Text = ""
    txtlugar.Text = ""
    'txtorden.Text = ""
    txtobservaciones.Text = ""
    txtuupp.Text = "": txtdesuupp.Text = ""
    txtproveedor.Text = "": pnlproveedor.Text = ""
    
    Call CargarUsuario
    cboprioridad.ListIndex = 0
    cmbestado.ListIndex = 0
    
    SSActiveToolBars1.Tools.ITEM("ID_Imprimir").Enabled = False
    SSActiveToolBars1.Tools.ITEM("ID_Email").Enabled = False
    SSActiveToolBars1.Tools.ITEM("ID_Eliminar").Enabled = False
    
    SSActiveToolBars2.Tools.ITEM("ID_SolicitadoPor").Enabled = True
    SSActiveToolBars2.Tools.ITEM("ID_AprobadoPor").Enabled = True
                     
End Sub

Private Sub CONFIGURA_GRID()
With Grid.Options

        .Set (egoEditing)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoCanDelete)
        .Set (egoCanAppend)
        .Set (egoCanInsert)
        .Set (egoImmediateEditor)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
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
         .Set (egoExpandOnDblClick)
         .Set (egoShowGrid)
         .Set (egoShowButtons)
         .Set (egoNameCaseInsensitive)
         .Set (egoShowHeader)
         .Set (egoShowPreviewGrid)
         .Set (egoShowBorder)
         .Set (egoDynamicLoad)

    End With
        
        
End Sub

Private Sub Grid_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
        
    If sw_nuevo_item = False Then
        If Action = daInsert Then
            Grid.Dataset.Edit
            Grid.Columns.ColumnByFieldName("ITEM").Value = Grid.Dataset.RecordCount + 1
            Grid.Columns.FocusedIndex = 0
        End If
    End If
           
End Sub

Private Sub Grid_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)

    If sw_nuevo_item = False Then
        If Action = daInsert Then
            If Grid.Dataset.RecordCount > 0 Then
                If Len(Trim(Grid.Columns.ColumnByFieldName("cod_producto").Value & "")) = 0 Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            Grid.Dataset.Delete
        End If
    End If

End Sub

Private Sub AdicionaItem()

Dim sw_nuevo_temp   As Boolean
Dim I               As Integer
    
    On Error Resume Next
    
    Grid.Dataset.Active = False
    Grid.Dataset.Close
    Grid.Dataset.ADODataset.ConnectionString = cnn_form
    Grid.Dataset.ADODataset.CommandText = "select * from solicitud"
    Grid.Dataset.Active = True
    Grid.Dataset.Open
    
'    If sw_nuevo_documento = False Then
'
'        If Grid.Dataset.RecordCount > 0 Then
'            For I = Grid.Dataset.RecordCount To 1 Step -1
'                Grid.Dataset.RecNo = I
'                Grid.Dataset.Delete
'            Next
'        End If
'        Grid.Dataset.Refresh
'    End If
    
    With Grid.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
        For I = 1 To 1
            If sw_nuevo_temp = False Then
                If sw_nuevo_documento = True Then
                    .Edit
                Else
                    .Append
                End If
                sw_nuevo_temp = True
            Else
                .Append
            End If
            .FieldValues("ITEM") = I
            .FieldValues("COD_PRODUCTO") = ""
            .FieldValues("DS_UNIDMED") = ""
            .FieldValues("DS_DESCRIPCION") = ""
            .FieldValues("DS_CANTIDAD") = Null
            .FieldValues("STOK") = Null
            .FieldValues("PRECIO") = Null
            .FieldValues("PRESUG") = Null
            .FieldValues("CS_FENTREGA") = Format$(Date, "dd/mm/yyyy")
            .FieldValues("F5CODFAB") = ""
            .FieldValues("F5MARCA") = ""
            .FieldValues("F5CODMARCA") = ""
    
        Next
        .Post
        
        '
        sw_nuevo_item = False
    End With
    Grid.Dataset.Close
    Grid.Dataset.Open
    Grid.Dataset.Refresh
    

End Sub

Private Sub AdicionaItemNuevo()

Dim sw_nuevo_temp   As Boolean
Dim I               As Integer
    
    
    
    Grid.Dataset.Active = False
    Grid.Dataset.Close
    DELETEREC_LOG cnomtabla, cnn_form
    Grid.Dataset.ADODataset.ConnectionString = cnn_form
    Grid.Dataset.ADODataset.CommandText = "SELECT * FROM SOLICITUD"
    Grid.Dataset.Active = True
    Grid.Dataset.Open
    
    With Grid.Dataset
        sw_nuevo_temp = False
        sw_nuevo_item = True
        .Append
        .FieldValues("ITEM") = 1
        .FieldValues("COD_PRODUCTO") = ""
        .FieldValues("DS_UNIDMED") = ""
        .FieldValues("DS_DESCRIPCION") = ""
        .FieldValues("DS_CANTIDAD") = Null
        .FieldValues("STOK") = Null
        .FieldValues("PRECIO") = Null
        .FieldValues("PRESUG") = Null
        .FieldValues("CS_FENTREGA") = Format$(Date, "dd/mm/yyyy")
        .FieldValues("F5CODFAB") = ""
        .FieldValues("F5MARCA") = ""
        .FieldValues("F5CODMARCA") = ""
        .FieldValues("ds_sinIgv") = 0
        .FieldValues("ds_conIgv") = 0
        .FieldValues("afecto") = 0
        .Post
        sw_nuevo_item = False
    End With
    Grid.Dataset.Close
    Grid.Dataset.Open
    Grid.Dataset.Refresh
    

End Sub


Private Sub txtuupp_DblClick()

    txtuupp_KeyDown 113, 0
    
End Sub

Private Sub txtuupp_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodcosto = "": wdescosto = ""
        sw_ayuda = True
        Ayuda_Centros.Show 1
        sw_ayuda = False
        If Len(Trim(wcodcosto)) > 0 Then
            If Grid.Dataset.RecordCount > 0 And Len(Trim(Grid.Columns(4).Value)) > 0 Then
                MsgBox "No puede cambiar la obra si ya eligió recursos", vbInformation + vbDefaultButton1, "Atención"
                Exit Sub
            Else
                txtuupp.Text = wcodcosto
                txtdesuupp.Text = wdescosto
                txtlugar.Text = ObtenerCampo("CENTROS", "F3DIRECCION", "F3COSTO", wcodcosto, "T", cnn_dbbancos)
                txtuupp_KeyPress 13
            End If
            If Len(Trim(txtuupp.Text)) > 0 And sw_cabecera = False Then
        sw_cabecera = True
    End If
        End If
    End If
            
End Sub

Private Sub txtuupp_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        ModUtilitario.pulsarTecla vbKeyTab 'SendKeys "{tab}"
    Else
        'KeyAscii = valida(3, KeyAscii)
    End If
    
End Sub

Private Sub txtuupp_LostFocus()

    If Len(Trim(txtuupp.Text)) > 0 Then
        If VALIDA_CC(txtuupp.Text) = True Then
            txtdesuupp.Text = wdescosto
            txtlugar.Text = ObtenerCampo("CENTROS", "F3DIRECCION", "F3COSTO", wcodcosto, "T", cnn_dbbancos)
        Else
            MsgBox "Centro de costo no existe", vbInformation + vbDefaultButton1, "Atención"
            txtuupp.Text = "" ': txtuupp.SetFocus
        End If
    End If

End Sub

Private Sub CargarEstado()
    cmbestado.Clear
    
    cmbestado.AddItem "Registrando" & Space(100) & "1"
    cmbestado.AddItem "Aprobado" & Space(100) & "2"
    cmbestado.AddItem "Atendido" & Space(100) & "3"
    cmbestado.AddItem "Cerrado" & Space(100) & "4"
    cmbestado.AddItem "Anulado" & Space(100) & "5"
    
    cmbestado.ListIndex = 0

End Sub

Private Sub CargarPrioridad()
   cboprioridad.Clear
     
   cboprioridad.AddItem "Normal" & Space(100) & "1"
   cboprioridad.AddItem "Baja" & Space(100) & "0"
   cboprioridad.AddItem "Alta" & Space(100) & "2"
    
   cboprioridad.ListIndex = 0

End Sub


Public Sub activar(Estado As Boolean)
txtfecha.Enabled = Not Estado
txtproveedor.Locked = Not Estado
cboprioridad.Locked = Not Estado
cmbestado.Locked = Not Estado
txtlugar.Locked = Not Estado
txtobservaciones.Locked = Not Estado
frmmoneda.Enabled = Estado
txttc.Locked = Not Estado

Grid.Columns.ColumnByFieldName("COD_PRODUCTO").DisableEditor = Not Estado
Grid.Columns.ColumnByFieldName("COD_PRODUCTO").DisableEditor = Not Estado
Grid.Columns.ColumnByFieldName("DS_CANTIDAD").DisableEditor = Not Estado
Grid.Columns.ColumnByFieldName("CS_FENTREGA").DisableEditor = Not Estado
End Sub

Public Sub ActualizaProductos()
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset

'Actualiza Maestro de Productos
If wactprod = "S" And wuseractprod = "S" And Len(Trim(wuserempresa)) > 0 Then
    xempresa = wuserempresa
   
    wfecha = Date
    If rst.State = adStateOpen Then rst.Close
    'SQL = "select * from if5pla where f5fecing<=CVDate( '" & wfecha & "') and f5tipo='P' and (f5descontinuado<>'S' or f5descontinuado IS NULL)"
    sql = "select * from if5pla where (f5transferido<>'S' or f5transferido is null) and f5tipo='P' and (f5descontinuado<>'S' or f5descontinuado IS NULL)"
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        Do While Not rst.EOF
            f5codfab = "" & rst("f5codfab")
            F5MARCA = "" & rst("f5marca")
            
            f5codpro = "" & rst("F5CODPRO")
            F5PREVTA = rst("F5PREVTA")
            F5VALVTA = Val("" & rst("F5VALVTA"))
            F5IGVVTA = Val("" & rst("F5IGVVTA"))
            F5TIPO = "" & rst("F5TIPO")
            f7codmed = "" & rst("F7CODMED")
            F5AFECTO = "" & rst("F5AFECTO")
            F5TIPPRO = "" & rst("F5TIPPRO")
            F5TIPESTADO = "" & rst("F5TIPESTADO")
            F5FOB = rst("F5FOB")
            F5FLETE = Val("" & rst("F5FLETE"))
            F5FACTOR = Val("" & rst("F5FACTOR"))
            F5MANIPULEO = Val("" & rst("F5MANIPULEO"))
            F5INSTALACION = Val("" & rst("F5INSTALACION"))
            F5OTROSCOSTOS = Val("" & rst("F5OTROSCOSTOS"))
            F5MONEDA = "" & rst("F5MONEDA")
            F5FACTOR_MIN = rst("F5FACTOR_MIN")
            F5ORIGEN = "" & rst("F5ORIGEN")
            F5CTACON = "" & rst("F5CTACON")
            F5MODELO = "" & rst("F5MODELO")
            F5PORDES = rst("F5PORDES")
            F5PRECOS = Val("" & rst("F5PRECOS"))
            F5PREMAY = Val("" & rst("F5PREMAY"))
            F5ESTVAL = "" & rst("F5ESTVAL")
            F5DESCUE = rst("F5DESCUE")
            f5partara = "" & rst("F5PARTARA")
            F5UBICA2 = "" & rst("F5UBICA2")
            F5CENTRO = "" & rst("F5CENTRO")
            F5INSUMO = "" & rst("F5INSUMO")
            F5CANTIDAD = rst("F5CANTIDAD")
            F5FECMOD = "" & rst("F5FECMOD")
            F5USERMOD = "" & rst("F5USERMOD")
            F2COD_ALM = "" & rst("F2COD_ALM")
            F5MONEDAORI = "" & rst("F5MONEDAORI")
            F5COSTOEURO = rst("F5COSTOEURO")
            F5TCEURO = rst("F5TCEURO")
            F5TIPOCOSTO = "" & rst("F5TIPOCOSTO")
            f5descontinuado = "" & rst("F5DESCONTINUADO")
            F5PRECIOMODIF = "" & rst("F5PRECIOMODIF")
            linea = "" & rst("linea")
            F5NOMPRO = "" & rst("F5NOMPRO")
            F5NOMPRO2 = "" & rst("F5NOMPRO2")
            F5TEXTO = "" & rst("F5TEXTO")
            F5TEXTO_ING = "" & rst("F5TEXTO_ING")
            F5FECING = "" & rst("F5FECING")
            'F5FECMOD = "" & rst("f5fecmod")
            If IsDate(F5FECMOD) Then
                cad1 = ",F5FECMOD"
                cad2 = ",cvdate('" & rst("F5FECMOD") & "') "
                cad3 = ",F5FECMOD=cvdate('" & rst("F5FECMOD") & "') "
            Else
                cad1 = ""
                cad2 = ""
                cad3 = ""
            End If
            
            rst("F5TRANSFERIDO") = "S"
            rst.Update
            
            If rst2.State = adStateOpen Then rst2.Close
            sql = "select f5codpro,f5fecmod from if5pla where f5codfab='" & f5codfab & "' and f5marca='" & F5MARCA & "'"
            rst2.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
            If rst2.EOF Then        'Producto no Existe
                Cab = "insert into if5pla (F5CODFAB,F5MARCA,F5CODPRO,F5PREVTA,F5VALVTA,F5IGVVTA,F5TIPO,F7CODMED, " _
                & "F5AFECTO,F5TIPPRO,F5TIPESTADO,F5FOB,F5FLETE,F5FACTOR,F5MANIPULEO,F5INSTALACION,F5OTROSCOSTOS,F5MONEDA, " _
                & "F5FACTOR_MIN,F5ORIGEN,F5CTACON,F5MODELO,F5PORDES,F5PRECOS,F5PREMAY,F5ESTVAL,F5DESCUE,F5PARTARA, " _
                & "F5UBICA2,F5CENTRO,F5INSUMO,F5CANTIDAD,F5USERMOD,F2COD_ALM,F5MONEDAORI,F5COSTOEURO,F5TCEURO, " _
                & "F5TIPOCOSTO,F5DESCONTINUADO,linea,F5NOMPRO,F5NOMPRO2,F5TEXTO,F5TEXTO_ING,F5FECING" & cad1 & ")"
                Det = "values ('" & f5codfab & "','" & F5MARCA & "','" & f5codpro & "'," & _
                F5PREVTA & "," & F5VALVTA & "," & F5IGVVTA & ",'" & F5TIPO & "','" & f7codmed & "','" & F5AFECTO & _
                "','" & F5TIPPRO & "','" & F5TIPESTADO & "'," & F5FOB & "," & F5FLETE & "," & F5FACTOR & "," & _
                F5MANIPULEO & "," & F5INSTALACION & "," & F5OTROSCOSTOS & ",'" & F5MONEDA & "'," & F5FACTOR_MIN & _
                ",'" & F5ORIGEN & "','" & F5CTACON & "','" & F5MODELO & "'," & F5PORDES & "," & F5PRECOS & "," & _
                F5PREMAY & ",'" & F5ESTVAL & "'," & F5DESCUE & ",'" & f5partara & "','" & F5UBICA2 & "','" & _
                F5CENTRO & "','" & F5INSUMO & "'," & F5CANTIDAD & ",'" & F5USERMOD & "','" & _
                F2COD_ALM & "','" & F5MONEDAORI & "'," & F5COSTOEURO & "," & F5TCEURO & ",'" & F5TIPOCOSTO & _
                "','" & f5descontinuado & "','" & linea & "','" & F5NOMPRO & "','" & F5NOMPRO2 & "','" & F5TEXTO & "','" & F5TEXTO_ING & "','" & F5FECING & "'" & cad2 & ")"
                
                cnn_dbbancos.Execute Cab & " " & Det
                'AlmacenaQuery_sql Cab & " " & Det, cnn_dbbancos
            Else        'Producto ya Existe
                'If IsDate(rst("f5fecmod")) Then
                    'If rst("f5fecmod") > rst2("f5fecmod") Then
                        sql = "update if5pla set F5CODFAB='" & f5codfab & "',F5MARCA='" & F5MARCA & "',F5CODPRO='" & f5codpro & "',F5PREVTA=" & F5PREVTA & ",F5VALVTA=" & _
                        F5VALVTA & ",F5IGVVTA=" & F5IGVVTA & ",F5TIPO='" & F5TIPO & "',F7CODMED='" & f7codmed & "',F5AFECTO='" & _
                        F5AFECTO & "',F5TIPPRO='" & F5TIPPRO & "',F5TIPESTADO='" & F5TIPESTADO & "',F5FOB=" & F5FOB & ",F5FLETE=" & _
                        F5FLETE & ",F5FACTOR=" & F5FACTOR & ",F5MANIPULEO=" & F5MANIPULEO & ",F5INSTALACION=" & _
                        F5INSTALACION & ",F5OTROSCOSTOS=" & F5OTROSCOSTOS & ",F5MONEDA='" & F5MONEDA & "',F5FACTOR_MIN=" & _
                        F5FACTOR_MIN & ",F5ORIGEN='" & F5ORIGEN & "',F5CTACON='" & F5CTACON & "',F5MODELO='" & _
                        F5MODELO & "',F5PORDES=" & F5PORDES & ",F5PRECOS=" & F5PRECOS & ",F5PREMAY=" & F5PREMAY & ",F5ESTVAL='" & _
                        F5ESTVAL & "',F5DESCUE=" & F5DESCUE & ",F5PARTARA='" & f5partara & "',F5UBICA2='" & _
                        F5UBICA2 & "',F5CENTRO='" & F5CENTRO & "',F5INSUMO='" & F5INSUMO & "',F5CANTIDAD=" & _
                        F5CANTIDAD & ",F5USERMOD='" & F5USERMOD & "',F2COD_ALM='" & _
                        F2COD_ALM & "',F5MONEDAORI='" & F5MONEDAORI & "',F5COSTOEURO=" & F5COSTOEURO & ",F5TCEURO=" & _
                        F5TCEURO & ",F5TIPOCOSTO='" & F5TIPOCOSTO & "',linea='" & linea & "'" & cad3 _
                        & "where f5codfab='" & f5codfab & "' and f5marca='" & F5MARCA & "'"
                        
                        cnn_dbbancos.Execute sql
                        'AlmacenaQuery_sql sql, cnn_dbbancos
                        Actualiza_Log sql, cnn_dbbancos.ConnectionString
                    'End If
                'End If
            End If
            rst.MoveNext
        Loop
    End If
    If rst2.State = adStateOpen Then rst2.Close
    If rst.State = adStateOpen Then rst.Close
    If cnn.State = adStateOpen Then cnn.Close
End If
End Sub

Public Sub Excel()
wexcel = 1
imprimir
End Sub
