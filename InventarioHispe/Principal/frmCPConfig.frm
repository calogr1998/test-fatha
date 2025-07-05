VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCPConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Modulo"
   ClientHeight    =   6735
   ClientLeft      =   4470
   ClientTop       =   2385
   ClientWidth     =   7335
   Icon            =   "frmCPConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7335
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10186
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Vales"
      TabPicture(0)   =   "frmCPConfig.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Orden de Compra"
      TabPicture(1)   =   "frmCPConfig.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "General"
      TabPicture(2)   =   "frmCPConfig.frx":05C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   " Configuraciones de Modulo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox chkAlertaStockMin 
            Caption         =   "Mostrar Alerta de Stock Minimo al cargar el Modulo."
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Configuracion de Orden de Compra/Servicio "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox chkOrdenCompraAbierto 
            Caption         =   "Orden de Compra/Servicio abierto en otra Ventana Principal."
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Configuracion de Vales "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox chkValeSalidaAbierto 
            Caption         =   "Vale de Salida abierto en otra Ventana Principal."
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1680
            Width           =   4935
         End
         Begin VB.CheckBox chkValeIngresoAbierto 
            Caption         =   "Vale de Ingreso abierto en otra Ventana Principal."
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   4935
         End
         Begin MSComCtl2.DTPicker dtpSalidaFecha 
            Height          =   300
            Left            =   4680
            TabIndex        =   4
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   97910785
            CurrentDate     =   41982
         End
         Begin MSComCtl2.DTPicker dtpIngresoFecha 
            Height          =   300
            Left            =   4680
            TabIndex        =   6
            Top             =   300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   97910785
            CurrentDate     =   41982
         End
         Begin VB.CheckBox chkIngresoUsarFechaDefault 
            Caption         =   "Usar Fecha por predeterminada para Vale de Ingreso."
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   4935
         End
         Begin VB.CheckBox chkSalidaUsarFechaDefault 
            Caption         =   "Usar Fecha por predeterminada para Vale de Salida."
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   4935
         End
      End
   End
   Begin MSComDlg.CommonDialog cmdlgConfiguracion 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCPConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Funcción que abre el cuadro de dialogo y retorna la ruta
'******************************************************************
Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
    On Local Error GoTo errFunction
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
    Exit Function
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
End Function
    
Private Sub Check2_Click()

End Sub

Private Sub chkIngresoUsarFechaDefault_Click()
    On Error Resume Next
    
    dtpIngresoFecha.Enabled = CBool(chkIngresoUsarFechaDefault.value)
    dtpIngresoFecha.SetFocus
End Sub

Private Sub chkSalidaUsarFechaDefault_Click()
    On Error Resume Next
    
    dtpSalidaFecha.Enabled = CBool(chkSalidaUsarFechaDefault.value)
    dtpSalidaFecha.SetFocus
End Sub

Private Sub cmdaceptar_Click()
    Dim strFichero As String
    
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeIngresoUsarFechaPredeterminada", IIf(CBool(chkIngresoUsarFechaDefault.value), 1, 0)
        ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeIngresoFechaPredeterminada", IIf(CBool(chkIngresoUsarFechaDefault.value), Trim(dtpIngresoFecha.value & ""), 0)
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeSalidaUsarFechaPredeterminada", IIf(CBool(chkSalidaUsarFechaDefault.value), 1, 0)
        ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeSalidaFechaPredeterminada", IIf(CBool(chkSalidaUsarFechaDefault.value), Trim(dtpSalidaFecha.value & ""), 0)
        
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeIngresoAbierto", IIf(CBool(chkValeIngresoAbierto.value), 1, 0)
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "ValeSalidaAbierto", IIf(CBool(chkValeSalidaAbierto.value), 1, 0)
    
    
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "OrdenCompraAbierta", IIf(CBool(chkOrdenCompraAbierto.value), 1, 0)
    
    ModUtilitario.sWrtIni strFichero, "ConfigCP", "MostrarAlertaStockMinimo", IIf(CBool(chkAlertaStockMin.value), 1, 0)
    
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    
    Unload Me
End Sub

Private Sub inicializarControles()
    Dim strFichero As String
    
    strFichero = wrutatemp & strNombreFicheroConfigCPusuario
    
    If Dir(Trim(strFichero), vbArchive) <> vbNullString Then
        chkIngresoUsarFechaDefault.value = IIf(ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoUsarFechaPredeterminada", "l") = "1", vbChecked, vbUnchecked)
            dtpIngresoFecha.value = IIf(CBool(chkIngresoUsarFechaDefault.value), ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoFechaPredeterminada", "l"), Date)
            dtpIngresoFecha.Enabled = CBool(chkIngresoUsarFechaDefault.value)
            
        chkSalidaUsarFechaDefault.value = IIf(ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaUsarFechaPredeterminada", "l") = "1", vbChecked, vbUnchecked)
            dtpSalidaFecha.value = IIf(CBool(chkSalidaUsarFechaDefault.value), ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaFechaPredeterminada", "l"), Date)
            dtpSalidaFecha.Enabled = CBool(chkSalidaUsarFechaDefault.value)
            
        chkValeIngresoAbierto.value = IIf(ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeIngresoAbierto", "l") = "1", vbChecked, vbUnchecked)
        chkValeSalidaAbierto.value = IIf(ModUtilitario.sGetINI(strFichero, "ConfigCP", "ValeSalidaAbierto", "l") = "1", vbChecked, vbUnchecked)
        
        chkOrdenCompraAbierto.value = IIf(ModUtilitario.sGetINI(strFichero, "ConfigCP", "OrdenCompraAbierta", "l") = "1", vbChecked, vbUnchecked)
        
        chkAlertaStockMin.value = IIf(ModUtilitario.sGetINI(strFichero, "ConfigCP", "MostrarAlertaStockMinimo", "l") = "1", vbChecked, vbUnchecked)
    Else
        MsgBox "Imposible cargar configuración, archivo .INI no existe.", vbInformation + vbOKOnly, App.ProductName
        
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    inicializarControles
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

