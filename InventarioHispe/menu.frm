VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.MDIForm menu 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Logística"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   1395
   ClientWidth     =   2880
   LinkTopic       =   "MDIForm1"
   Tag             =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrTitle 
      Interval        =   100
      Left            =   600
      Top             =   0
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   2
      ToolsCount      =   102
      PersonalizedMenus=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "menu.frx":0000
      ToolBars        =   "menu.frx":3EE84
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
    
    Select Case pGroup.Index
        Case 0:
            Select Case ItemLinkIndex
                Case 0: lista_solicitudes.Show 1
                Case 1:
                    wtipo_orden = "1"
                    lista_oc.Show 1
                Case 2:
                    wtipo_orden = "0"
                    lista_oc.Show 1
                Case 3: frmregiscom.Show 1
                Case 4: comprobante_retencion.Show 1
                Case 5: Importaciones.Show 1
            End Select
        Case 1:
            Select Case ItemLinkIndex
                Case 0:
                    wtipoguia = "I"
                    lista_vales.Show 1
                Case 1:
                    wtipoguia = "S"
                    lista_vales.Show 1
                Case 2:
                    FrmTransferencias.Show 1
                Case 3:
                    FrmTransforma.Show 1
                Case 4:
                    ProcesaVale.Show 1
                Case 5:
                    'frmformula.Show 1
                    ListaFormulas.Show 1
            End Select
        Case 2:
            Select Case ItemLinkIndex
                Case 0: cons_solicitudes.Show 1
                Case 1: cons_ordenes_compra.Show 1
                Case 2: kardex.Show 1
                Case 3: ocompra_pendientes.Show 1
                Case 4: Consulta_PreciosdeProductos.Show 1
                Case 5: inventario_valorizado.Show 1
                Case 6: RInventario.Show 1
                
            End Select
        Case 3:
            Select Case ItemLinkIndex
                Case 0: lista_prod.Show 1
                Case 1: Lista_proveedores.Show 1
            End Select
    End Select
    
End Sub

Private Sub MDIForm_Load()
Dim cnn_dbingre     As New ADODB.Connection
Dim rsingre         As New ADODB.Recordset
        
    ctipoadm_bd = "" '---- "M" -> mysql
    
    If wingreso = False Then
        cnn_dbingre.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\BANCOWIN\DB_INGRE.MDB" & ";Persist Security Info=False"
        rsingre.Open "SELECT * FROM TB_INGRESO", cnn_dbingre, adOpenDynamic, adLockOptimistic
        If Not rsingre.EOF Then
            rsingre.MoveFirst
            wempresa = Trim(rsingre.Fields("EMPRESA") & "")
            wusuario = Trim(rsingre.Fields("USUARIO") & "")
            
            cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RUTAS.MDB" & ";Persist Security Info=False"
            If cnn_control.State = adStateOpen Then cnn_control.Close
            cnn_control.Open cconex_control
            rscontrol.Open "SELECT * FROM SRUTAS WHERE EMPRESA ='" & wempresa & "'", cnn_control, adOpenDynamic, adLockOptimistic
            If Not rscontrol.EOF Then
                wrutaconta = Trim(rscontrol.Fields("CONTABILIDAD") & "")
                wrutabancos = Trim(rscontrol.Fields("BANCOS") & "")
                wrutatemp = Trim(rscontrol.Fields("TEMPORALES") & "")
                wcontacnt = Trim(rscontrol.Fields("CONTACNT") & "")
            End If
            rscontrol.Close
            cnn_control.Close
            
        End If
        rsingre.Close
        cnn_dbingre.Close
    End If
    StrConexDbBancos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_bancos.MDB;Persist Security Info=False"
    cconex_dbbancos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\DB_BANCOS.MDB" & ";Persist Security Info=False"
    If cnn_dbbancos.State = adStateOpen Then cnn_dbbancos.Close
    cnn_dbbancos.Open cconex_dbbancos
    
    '----------------------------------------------------
    cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CNT_BANC.MDB" & ";Persist Security Info=False"
    cnn_control.Open cconex_control
    
    If rscontrol.State = adStateOpen Then rscontrol.Close
    rscontrol.Open "SELECT * FROM BF1CNT WHERE F1DIR ='" & wempresa & "'", cnn_control, adOpenDynamic, adLockOptimistic
    If Not rscontrol.EOF Then
        wf1renovacion = Trim(rscontrol.Fields("F1TIPO_RENOV_XPAGAR") & "")
        wf1agente = Trim(rscontrol.Fields("F1AGENTERET") & "")
    End If
    rscontrol.Close
    cnn_control.Close
    '----------------------------------------------------
    
    If cnn_control.State = adStateOpen Then cnn_control.Close
    cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CONTROL.MDB" & ";Persist Security Info=False"
    cnn_control.Open cconex_control
    
    If rscontrol.State = adStateOpen Then rscontrol.Close
    rscontrol.Open "SELECT * FROM SF1PARAIN WHERE F1CODEMP ='" & wempresa & "'", cnn_control
    If Not rscontrol.EOF Then
        wconc_compra = rscontrol.Fields("F1CONC_COMPRA") & ""
        wconc_salxtransf = rscontrol.Fields("F1CONC_SALXTRANSF") & ""
        wconc_salida = rscontrol.Fields("F1CONC_SALIDA") & ""
        wconc_ing_obra = rscontrol.Fields("F1CONC_ING_OBRA") & ""
        wtiposalida = Trim(rscontrol.Fields("F1TIPOSALIDA") & "")
        wmoneda_productos = Trim(rscontrol.Fields("F1MONEDA_PRODUCTOS") & "")
    End If
    rscontrol.Close
    
    If rscontrol.State = adStateOpen Then rscontrol.Close
    rscontrol.Open "SELECT * FROM SF1PARAM WHERE F1CODEMP ='" & wempresa & "'", cnn_control
    If Not rscontrol.EOF Then
        wf1visualiza_precio_hlp = rscontrol.Fields("F1COLVALVTA") & ""
    End If
    rscontrol.Close
    
    cconex_ctrcom = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CTRCOM.MDB" & ";Persist Security Info=False"
    cnn_ctrcom.Open cconex_ctrcom
    
    If rsparam_com.State = adStateOpen Then rsparam_com.Close
    rsparam_com.Open "SELECT * FROM PARAM_COM WHERE F1CODEMP ='" & wempresa & "'", cnn_ctrcom
    If Not rsparam_com.EOF Then
        wnomcia = Trim(rsparam_com.Fields("F1NOMEMP") & "")
        wrucempresa = Trim(rsparam_com.Fields("F1RUCEMP") & "")
        wemailsol = Trim(rsparam_com.Fields("F1EMAIL_SOLICITUD") & "")
        wemailccsol = Trim(rsparam_com.Fields("F1EMAIL_CCSOL") & "")
        wasuntosol = Trim(rsparam_com.Fields("F1ASUNTO_SOL") & "")
        wtextosol = Trim(rsparam_com.Fields("F1TEXTO_SOL") & "")
        wemailoc = Trim(rsparam_com.Fields("F1EMAIL_OC") & "")
        wemailccoc = Trim(rsparam_com.Fields("F1EMAIL_CCOC") & "")
        wasuntooc = Trim(rsparam_com.Fields("F1ASUNTO_OC") & "")
        wtextooc = Trim(rsparam_com.Fields("F1TEXTO_OC") & "")
        
        mes = right(rsparam_com.Fields("f1proame") & "", 2)
        wocompra = Trim(rsparam_com.Fields("f1ocompra") & "")  '---- MODIFICAR POR wf1trasoc
        wf1tipdoc_asoc = rsparam_com.Fields("F1TIPDOC_ASOC") & ""
        'wf1inggasto = rsparam_com.Fields("f1inggasto") & ""
        wbancos = Trim(rsparam_com.Fields("f1bancos") & "")
        'wf1traslado = rsparam_com.Fields("f1traslado") & ""
        gctapag = "" & rsparam_com.Fields("F1CTAPAG")
        wf1formatov = rsparam_com.Fields("f1formatov") & ""
        wf1numera = rsparam_com.Fields("f1numera") & ""
        wf1viscod = rsparam_com.Fields("f1viscod") & ""
        wf1trasoc = rsparam_com.Fields("F1TRASOC") & ""
        
        wIgv = Val("" & rsparam_com.Fields("F1IGV"))
        
        gfonavi = Val("" & rsparam_com.Fields("F1FONAVI"))
        gretenc = Val("" & rsparam_com.Fields("F1RETENC"))
        'wingobra = Trim(rsparam_com.Fields("f1ingobra") & "")
        worigen = Trim(rsparam_com.Fields("f1origen") & "")
        wctaigv = Trim(rsparam_com.Fields("f1ctaigv") & "")
        wctaotros = Trim(rsparam_com.Fields("f1ctaotros") & "")
        wredsuma = Trim(rsparam_com.Fields("f1redsuma") & "")
        wredresta = Trim(rsparam_com.Fields("f1redresta") & "")
        wctaret = Trim(rsparam_com.Fields("f1ctaret") & "")
        wctafon = Trim(rsparam_com.Fields("f1ctafonavi") & "")
        wf1formato = rsparam_com.Fields("f1formatorc") & ""
        wanno = left(rsparam_com.Fields("f1proame") & "", 4)
        wdcto = Trim(rsparam_com.Fields("f1dcto") & "")
        'wf1cnting = rsparam_com.Fields("f1cnting") & ""
        wf1uupp = Trim(rsparam_com.Fields("F1UUPP") & "")
        wf1show_ccosto = Trim(rsparam_com.Fields("F1SHOW_CCOSTO") & "")
        wf1direc1 = Trim(rsparam_com.Fields("F1DIREMP") & "")
        wf1direc2 = Trim(rsparam_com.Fields("F1DIREMP2") & "")
        wf1visualiza_dctos = Trim(rsparam_com.Fields("F1VISUALIZA_DCTOS") & "")
        wf1visualiza_advalorem = Trim(rsparam_com.Fields("F1VISUALIZA_ADVALOREM") & "")
        wf1visualiza_import_venta = Trim(rsparam_com.Fields("F1IMPORT_VENTA") & "")
        wvisualiza_cod = Trim(rsparam_com.Fields("F1VISUALIZA_COD") & "")
    End If
    rsparam_com.Close
        
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    cnn_control.Close
    
    cnn_ctrcom.Close
    
    cnn_dbbancos.Close
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    
    Select Case Tool.Id
        '----- LOGISTICA
'        Case "ID_SolicituddeMateriales": lista_solicitudes.Show 1
        Case "ID_OrdendeCompra":
            wtipo_orden = "1"
            frmListaOrden.Show
'        Case "ID_OC_Internacional":
'            wtipo_orden = "0"
'            lista_oc.Show 1
'        Case "ID_RegistrodeCompras": frmregiscom.Show 1
'        Case "ID_ComprobantedeRetenciòn": comprobante_retencion.Show 1
'        Case "ID_ImportacionesReg": Importaciones.Show 1
'
'        Case "ID_Salir": Unload Me
'
'        '----- INVENTARIOS
'        Case "ID_ValedeIngreso": wtipoguia = "I": lista_vales.Show 1
'        Case "ID_ValedeSalida": wtipoguia = "S": lista_vales.Show 1
'        Case "ID_Transferencias": FrmTransferencias.Show 1
'
'        '----- CONSULTAS
'        Case "ID_Cons_SolicituddeMateriales": cons_solicitudes.Show 1
'        Case "ID_Cons_OrdenesdeCompra": cons_ordenes_compra.Show 1
'        Case "ID_Kardex": kardex.Show 1
'        Case "ID_OrdenesdeCompraPendientes": ocompra_pendientes.Show 1
'        Case "ID_Productos": lista_prod.Show 1
'        Case "ID_ListadePrecios": Consulta_PreciosdeProductos.Show 1
'
'        Case "ID_Aplicacion": aplica.Show 1
'        Case "ID_FormadePago": forma_pago_contado.Show 1
'        Case "ID_InventarioValorizado": inventario_valorizado.Show 1
'        Case "ID_ResumenporCentrodeCosto": ResuCenCosto2.Show 1
'        Case "ID_SaldosIniciales": FrmSaldoAlm.Show 1
'        Case "ID_LibroMensual": cons_libro_compret.Show 1
'        Case "ID_LibroMensualReg": FrmRepRegistro.Show 1
'        Case "ID_NùmerodeComprobante": frmimpco.Show 1
'        Case "ID_Proveedor/Gastos": FrmProveGasto.Show 1
'        Case "ID_CuentasContables": c_cuenta.Show 1
'        Case "ID_ResumendeProvisiones": frmprovi.Show 1
'        Case "ID_General": frmrcege.Show 1
'        Case "ID_Detallado": frmrepce.Show 1
'        Case "ID_Honorarios": frmhonor.Show 1
'        Case "ID_ComprasporMes": frmcomxmes.Show 1
'        Case "ID_Importaciones": FrmRepImporta.Show 1
'        Case "ID_Generar": frm_gen.Show 1
'        Case "ID_Consultas_as": frm_cons.Show 1
'        Case "ID_Transferir": frmtrans.Show 1
'        Case "ID_SinCuadrar": frmdesc.Show 1
'        Case "ID_CierredelMes": frciemes.Show 1
'        Case "ID_TiposdeCambio": frmcontc.Show 1
'        Case "ID_SinCuadrar_Com": frmconbd.Show 1
'        Case "ID_Generar_Ret": generar_ret.Show 1
'        Case "ID_Transferir_Ret": transferir_ret.Show 1
'        'Case "ID_ActualizaciòndeProveedores/Productos": actualiza_prov_prod.Show 1
'        Case "ID_RegistroControldeRetenciones": cons_reg_cnt_ret.Show 1
'        Case "ID_ParàmetrosdelSistema": parametros.Show 1
'        Case "ID_InterfazComprobantesdeRetenciòn": compret_interface_pdt.Show 1
'        Case "ID_ActualizaciondeProductos/Proveedores": actualiza_producto_prov.Show 1
'
'        Case "ID_Proveedores": Lista_proveedores.Show 1
'
'        Case "ID_Regenerar_Saldos": FrmRegenerar.Show 1
        
    End Select

End Sub
