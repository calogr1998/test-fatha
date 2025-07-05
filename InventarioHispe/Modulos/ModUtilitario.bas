Attribute VB_Name = "ModUtilitario"
Option Explicit
'-------------------------------------------------------------
'Declaraciones del api - Inhabilitar Boton Cerrar de Ventana
'-------------------------------------------------------------
' PAra deshabilitar el menú y otros
Public Declare Function DeleteMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long

' Obtiene el Handle al menú del sistema de la ventana
Public Declare Function GetSystemMenu Lib "user32" ( _
    ByVal HWnd As Long, _
    ByVal bRevert As Long) As Long
''***************************************************************************************************************
''***************************************************************************************************************


'-------------------------------------------------------------
'Declaraciones del api - Para envio de Funcion de Teclas
'-------------------------------------------------------------
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
                                                ByVal bScan As Byte, _
                                                ByVal dwFlags As Long, _
                                                ByVal dwExtraInfo As Long)
''***************************************************************************************************************


'Declaración de Function para Lectura de Ficheros en win32
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declaración de Function para Escritura de Ficheros en win32
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long
''***************************************************************************************************************
'***************************************************************************************************************
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, _
                                                                                ByVal lpOperation As String, _
                                                                                ByVal lpFile As String, _
                                                                                ByVal lpParameters As String, _
                                                                                ByVal lpDirectory As String, _
                                                                                ByVal nShowCmd As Long) As Long




'Función SendMessage
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long


'Para deshabilitar boton cerrar
Public Const MF_BYPOSITION = &H400&
'Para Envio de Teclas
Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_EXTENDEDKEY = &H1

Rem Objetos de Conexion
Public cnRutas As New ADODB.Connection
Rem Objetos de Formulario
Public objLogin As frmLogin
Public cnControl As New ADODB.Connection
Public cnBancos As New ADODB.Connection
Rem Objetos de Conexion
Public rsRutas As New ADODB.Recordset
'Public rscontrol As New ADODB.Recordset
Public rsTipoCambio As New ADODB.Recordset
Public rsAccesos As New ADODB.Recordset
Public rsPermisos As New ADODB.Recordset


Public srazonSocial As String
Public snombreComercial As String
Public stelefono As String
Public scondicion As String
Public ssistElectronica As String
Public sdireccion As String
'                snombreComercial = datos(i)
'            Case "telefono"
'                stelefono = datos(i)
'            Case "condicion"
'                scondicion = datos(i)
'            Case "sistElectronica"
'                ssistElectronica = datos(i)
'            Case "actEconomicas"
'                sactEconomicas = datos(i)
'            Case "direccion"
'                sdireccion

Public datosUser As Usuario
Rem Estructura de Usuario
Type Usuario
    codUser As String
    nomUser As String
    codNivel As String * 3
    CentroCosto As String * 3
End Type

Type a_numero
    valor As Integer
End Type

Type a_solicitud
    TOTAL As Byte
    numero As String
End Type

Type a_grabacion
    campo   As String
    valor   As String
    Tipo    As String
End Type
Public sw_GRABA_REGISTRO_logistica    As Boolean
Public Sw_Ejecuta_Sentencia As Boolean
Global lista() As a_solicitud
Global wllamada As Byte
Global oTipoRequerimiento As String
Public TCCompra      As String
Public TCVenta     As String
Public ruc_rsocial      As String
Public ruc_direccion    As String
Public ruc_estado       As String
Public ruc_situacion    As String
Public ruc_telefono     As String
Public strODBCtablas As String
Public strTABLAempant As String
Public strODBCsiete As String
Public strODBCsgte As String
Public csql As String
Public nGridActive As Integer
Public SqlCad As String
Public est1                 As String

Public wMailNoReplyCia      As String
Public wPassNoReplyCia      As String
Public wMailReplyCia        As String
Public wPassReplyCia        As String


Public ECS_otipo            As String
Public Cantidad()           As Integer

Public cconex_dbbancos      As String
Public ctipoadm_bd          As String  '----> Mysql
Public cconex_paramcom      As String

Public xvale                As Integer
Public cnn_form             As New ADODB.Connection
Public wNameImage           As String
'envia
Public wDestinatarios       As String
Public wDestinatarioOculto  As String
Public wSubject             As String
Public wFileName            As String

Public wRutaEnvia           As String
Public wIndEnvia            As String

Public ccodprod             As String
Public wLocal               As String
Public wObra                As String
Public Cod_Prove            As String
Public cCodConcar           As String
Public cCodDocConcar        As String

Public cnn_dbbancos         As New ADODB.Connection
Public cnn_dbEnvia          As New ADODB.Connection
Public dbcontrol            As New ADODB.Connection
Public rsalmacen            As New ADODB.Recordset
Public rsif5pla             As New ADODB.Recordset
Public rsconcepto_inv       As New ADODB.Recordset
Public RsProveedor          As New ADODB.Recordset
Public rscambios            As New ADODB.Recordset
Public rsccosto             As New ADODB.Recordset
Public rsdocumentos         As New ADODB.Recordset
Public rsif4orden           As New ADODB.Recordset
Public rsif3orden           As New ADODB.Recordset
Public rsif4vales           As New ADODB.Recordset
Public rsif3vales           As New ADODB.Recordset
Public rscontrol            As New ADODB.Recordset
Public rsif6alma            As New ADODB.Recordset
Public rsusuarios           As New ADODB.Recordset
Public rsordtrab            As New ADODB.Recordset
Public rslista              As New ADODB.Recordset
Public rsparam_com          As New ADODB.Recordset
Public RSCTA_DCTO           As New ADODB.Recordset
Public RSPAG_DCTO           As New ADODB.Recordset
Public rsbancos             As New ADODB.Recordset
Public rsmarcas             As New ADODB.Recordset
Public rsformpag            As New ADODB.Recordset

Public rsCabFormula         As New ADODB.Recordset
Public rsDetFormula         As New ADODB.Recordset

Public wcod_alm             As String
Public wnomalmacen          As String
Public wcodprov             As String
Public wrucprov             As String
Public wnomprov             As String
Public wcontacto            As String


Public gtipmov              As String * 1
Public wconcepto            As String
Public wnomconcepto         As String
Public wcodcosto            As String
Public wdescosto            As String
Public wunicosto            As String
Public wclicosto            As String
Public wIgv                 As Double
Public wMonto2doVb          As Double
Public wFob                 As Double
Public wDesaduana           As Double
Public wAdela               As Double
Public wwigv                As Double
Public wcodmar              As String
Public wcodmed              As String
Public wtipmov              As String * 1
Public wcodban              As String
Public wnomban              As String
Public wactprod             As String * 1
Public wuseractprod         As String * 1
Public wuserempresa         As String

Public wcodigo_barra        As String
Public wsw_codbarra         As Integer

Public sw_nuevo_documento   As Boolean
Public sw_ayuda_provee      As Boolean
Public wtipoguia            As String * 1

Public WBD                  As String
Public wrutaconta           As String
Public wrutabancos          As String
Public wrutatemp            As String
Public wusuario             As String
Public wempresa             As String

Public wconc_compra         As String
Public wconc_salxtransf     As String
Public wconc_salida         As String
Public wconc_ing_obra       As String
Public wtiposalida          As String * 1
Public wmoneda_productos    As String * 1

Public CodFirmaSolicitud(2)     As String * 20
Public CodFirmaAprobacion(2)    As String * 20

Public wcodsolicitud            As String
Public wcodproducto             As String
Public wcodpartida              As String
Public wcodpresupuesto          As String
Public wcodfab                  As String
Public wdesproducto             As Variant
Public wmedida                  As String
Public wcantidad                As Double
Public wcantidadmaxima          As Double
Public wf5partara               As String
Public wmarca                   As String
Public wnomsolicitante          As String
Public wcodsolicitante          As String
Public wpartar                  As String

Public wemailsol                As String
Public wemailccsol              As String
Public wasuntosol               As String
Public wtextosol                As String


Public wemailoc                 As String
Public wemailccoc               As String
Public wasuntooc                As String
Public wtextooc                 As String
Public wnomcia                  As String

Public codmprima                As String
Public nommprima                As String
Public uniprima                 As String
Public sw                       As Integer
Public sw_creabd                As Boolean

Public WMONEDAX                 As String
Public wemail_prove             As String

Public cconex_control       As String
Public cnn_control          As New ADODB.Connection

Public cconex_ctrcom        As String
Public cnn_ctrcom           As New ADODB.Connection

Public cconex_ctaspag       As String
Public cnn_ctaspag          As New ADODB.Connection

Public sw_cabecera          As Boolean

Public cnombase             As String
Public cnomtabla            As String
Public cnomtabla2           As String
Public wCost                As Double
Public wCost2                As Double
Public wcant                As Double
Public wf1visualiza_precio_hlp      As String
Public wf1elim_item_dcto    As String
Public wf1mant_productos    As String

Public DBTable              As String

Public gserie               As String

Public rsif5placli          As New ADODB.Recordset
Public rstempo              As New ADODB.Recordset
Public DESCARGADO           As Integer
Public codprod              As String
Public sw_nuevo_doc         As Boolean
Public sw_elimina           As Boolean
Public gcodzon              As String
Public gnomprov             As String
Public gtipcam              As Double
Public TbStockCab           As DAO.Recordset
Public TbStockDet           As DAO.Recordset
Public wtipoayuda           As String
Public wdescripcion         As String
Public vcodigoalmacen       As String
Public vcodigocentro        As String
Public TB_CabOC             As DAO.Recordset
Public TB_Proveedor         As DAO.Recordset
Public Tb_Personal          As DAO.Recordset
Public Tb_Formapago         As DAO.Recordset
Public Tb_costo             As DAO.Recordset
Public wdirprov             As String
Public wcodlocalidad        As String
Public wdeslocalidad        As String
Public wdireccion           As String
Public wDistrito            As String
Public wCosPostal           As String
Public wPais                As String
Public wtelefono            As String
Public wfax                 As String
Public wnumord              As String

Public wcodcli              As String
Public wnomcli              As String
Public cmes(12)             As String * 30
Public Rs                   As New ADODB.Recordset
Public gfilas               As Integer
Public mostrar              As Boolean
Public codigorucs           As String
Public ruc                  As String
Public codpro               As String
'Public nombre               As String
Public Direccion            As String
Public sw_nuevo_mant        As Boolean
Public sw_mant_ayuda        As Boolean
Public sw_load_mant         As Boolean
Public wvv_prod             As Double
Public wpv_prod             As Double
Public wstockact            As Double
Public wprecos              As Double
Public wprecosdol           As Double
Public wf1uupp              As String
Public wrucempresa          As String
Public wfile                As String
Public wf1show_ccosto       As String
Public wingreso             As Boolean
Public sw_ayuda_marca       As String
Public wcodgasto            As String
Public wnomgasto            As String
Public wctagasto            As String
Public wf1direc1            As String
Public wf1direc2            As String
Public wf1visualiza_dctos   As String

Public cconex_dbtabla           As String
Public cnn_dbtabla              As New ADODB.Connection
Public cconex_analisis          As String
Public cnn_analisis             As New ADODB.Connection
Public wfpagoprov               As String
Public wprecio_prod             As Double
Public wtipo_orden              As String
Public wcodusuario              As String
Public wnomusuario              As String
Public wnumordentrab            As String
Public wobservacion             As String
Public sw_ayuda_prod            As Boolean
Public wf1visualiza_advalorem       As String
Public wf1visualiza_import_venta    As String
Public sw_importa_valedeingreso     As Boolean
Public wafecto                      As String
Public wtipocc                      As Double
Public wFORMULA                     As Integer
'--------- PARA EL CONTROL DE LA ELIMINACION DE ITEM'S ------------------'
Public wdxcodigo                As String
Public wdxcodfab                As String
Public wdxdescripcion           As String
Public wdxcantidad              As Double
Public wdxnroitems              As Double
'-------------------------------------------------------------------------'
Public wcodzona                 As String
Public wcodmarca                As String
Public wexito                   As Boolean
Public wnomlinea                As String
Public wcodlinea                As String
Public wf1evalua_stock          As String
Public whelpoc                  As String
Public wf1formato_inventario    As String
Public wf1decimal_cantidad      As Integer
Public wf1decimal_costo         As Integer
Public wf5codpro                As String

Type S_Importa
     Orden    As String
     f4falta  As String
End Type


Global wrutabanco1              As String
Global wrutatemp1               As String
Global wempresa1                As String
Global WNUMERO1                 As Integer
Global WTipoCambio              As Double
Global wimporta(0 To 10)        As S_Importa


Public TbCabOrden1              As New ADODB.Recordset
Public TbDetImport1             As New ADODB.Recordset
Public TbDetTmpImp1             As New ADODB.Recordset
Public rst                      As New ADODB.Recordset


'-----------------------RECORDSET ---------------

Global RSCONSULTA     As New ADODB.Recordset
Global Temp           As New ADODB.Connection
Global RsMovAlmacen   As ADODB.Recordset
Global RsAlmacenes    As ADODB.Recordset
Global RsStockDet     As ADODB.Recordset
Global RsStockCab     As ADODB.Recordset
Global RsProducto     As ADODB.Recordset
Global RsParametro    As ADODB.Recordset
Global RsCentros      As ADODB.Recordset
Global RsMedida       As ADODB.Recordset
Global RsOrigen       As ADODB.Recordset
Global RsNivel01      As ADODB.Recordset
Global RsNivel02      As ADODB.Recordset
Global RsNivel03      As ADODB.Recordset
Global RsNivel04      As ADODB.Recordset
Global RsPartida      As ADODB.Recordset
Global RsParain       As ADODB.Recordset

'------------------ VARIABLES -------------------

Public cconex_inventa                   As String
Global F                            As Integer
Global wmes                         As String
Global stockact                     As Double
Global sal                          As Double
Global sald                         As Double
Global ing                          As Double
Global ingd                         As Single
Global Cospro                       As Double
Global Cosprod                      As Double
Global Debm                         As Double
Global Habm                         As Double
Global StockLog                     As Double
Global FecUlt                       As Date
Global sql                          As String
Global SQL1                         As String
Global wnumval                      As String
Global Gtipval                      As String
Global wcodPrv                      As String
Global wnomPrv                      As String
Global WcodPar                      As String
Global wnomcosto                    As String
Global WNomPar                      As String
Global wcodori                      As String
Global wnomori                      As String
Global wparact_stock                As String
Global wtipcam                      As Double
Global ctipo                        As String * 1
Global wstock                       As Double
Global wvalvta                      As Double
Global sw_ocompra                   As Boolean
Global DBTable3                     As String
Global wtipprov                     As String
Global wtipoc                       As String
Global tbcodigos                    As DAO.Recordset
Public wcodcliprov                  As String
Public wnomcliprov                  As String
Public wruccli                      As String
Public TipoExporta                  As Integer
Public sw_nuevo_item                As Boolean
Public sw_e_ordenpago               As Boolean
Public sw_est_orden_pago            As Boolean
Public sw_ordendepago               As Boolean
'''SE Agrega La condicion de Ayuda_Producto
Global Con_Ayu              As Integer
Public strMessageHTML       As String
Global Codigo_producto      As String
Public Sw_Act As Boolean
Public StrConexDbBancos     As String
Public FrmName              As String
Public cconex_formp         As String
Public strFilePath          As String
Public swActOrden           As Boolean
Public strOrdenCompra       As String
Public wRucCliProv          As String
Public wnomabrevcliprov     As String
Public wCodConcar           As String
Public StrConexControl      As String
Public wF1Dir      As String
Public wTipoReq As Byte
Public wTipoOC As Byte
Public strODBC As String
Public strTabla As String
Public cuerpo As String
Public nombre As String
Public codUser As String
Public destino As String
Public asunto As String

Public wprodfactor As String

Public rstvp       As New ADODB.Recordset


Public sw_ayuda             As Boolean
Public wdesgrupo            As String
Public wcodgrupo            As String


Rem SK ADD:
Public StrConexPlanilla         As String
Public StrConexControlCompra    As String
Public StrConexControlBanco    As String
Public StrConexCntCont      As String
Public StrConexContawinTabla As String
Public cnn_Planilla                 As New ADODB.Connection
Public cnn_ControlCompra            As New ADODB.Connection
Public cnn_ControlBanco             As New ADODB.Connection
Public cnn_CntCont                  As New ADODB.Connection
Public cnn_ContawinTabla            As New ADODB.Connection
Public cnn_ConcarVinculada          As New ADODB.Connection

Public x                            As Integer

Global wprecio                      As String * 1
Global wcolvalvta                   As String * 1
Global walmagui                     As String
Global walmafac                     As String
Global walmabol                     As String
Global walmadeb                     As String
Global walmacre                     As String
Global wpartipcam                   As String
Global wpardecimal                  As String
Global wbasetemp                    As String
Public wvisualiza_act               As String
Public wruc                         As String
Public wimprimir_ruc                As String
Public wf1codcli_anulado            As String
Public wf1codcon_anulado            As String
Public wf1formapago_contado         As String
Public wf1dscto_contado             As Double
Public wf1puntodeventa              As String
Public wf1grababoleta               As String
Public wflete                       As String
Public wpreciofac                   As String
Public wf1monedalistaprecios        As String
Public wf1visualiza_det_lista       As String
Public wf1evalua_linea_venc         As String
Public wf1formato_rv                As String
Public wparamcliente                As String
Public wparamprod                   As String
Public wf1anno                      As String
Public wf1distrito                  As String
Public wf1impresoras                As String
Public wf1facturar_diario_correla   As String
Public wf1sistema_venta_proyectos   As String
Public wf1control_menu              As String
Public wf1redondeo_dec1             As String
Public wf1gen_vale_nc               As String
Public wf1trasladactasxcob          As String
Public almaTrans                    As String

Public comisionVen      As Double
Public wbasetempND                  As String
Global wfestemp         As String

Public conCompras      As New ADODB.Connection
Public rsCompras       As New ADODB.Recordset
Public Const strNombreFicheroConfigCPgeneral = "\configCP.ini"
Public Const strNombreFicheroConfigCPusuario = "\configCPusuario.ini"
Public Const strNombreFicheroConfigSQLCliente = "\configServidorSQLCliente.ini"

Public strCadenaConexioBdCPlus      As String

Public cnDBTemp                     As ADODB.Connection
Public cnDBContaTabla               As ADODB.Connection

Public cnBdCPlus                    As ADODB.Connection

'-------------------------------------------------------------
Rem Para recuperar informacion desde Backup Db_Bancos.mdb
Public cnn_dbbancos_backup          As New ADODB.Connection
Public StrConexDbBancos_backup      As String
Public cconex_dbbancos_backup       As String
'-------------------------------------------------------------

'Public objAyudaBien                 As New ClsBien
'Public objAyudaBienAlterno          As New ClsBienAlterno
'Public objAyudaBienColor            As New ClsBienColor
'Public objAyudaCategoria            As New ClsCategoria
'Public objAyudaCliente              As New ClsCliente
'Public objAyudaCompra               As New ClsCompra
'Public objAyudaComprobante          As New ClsComprobante
'Public objAyudaDistrito             As New ClsDistrito
'Public objAyudaFamilia              As New ClsFamilia
'Public objAyudaEmpresa              As New ClsEmpresa
'Public objAyudaFormaPago            As New ClsFormaPago
'Public objAyudaGasto                As New ClsGasto
Public objAyudaOrden                As New ClsOrden
'Public objAyudaOrdenTrabajo         As New ClsOrdenTrabajo
'Public objAyudaOrigen               As New ClsOrigen
'Public objAyudaPagDcto              As New ClsPagDcto
'Public objAyudaProvDscto            As New ClsProvDscto
Public objAyudaProveedor            As New ClsProveedor
'Public objAyudaSector               As New ClsSectorEmpresarial
'Public objAyudaSolicitud            As New ClsSolicitud
'Public objAyudaSubFamilia           As New ClsSubFamilia
'Public objAyudaTCambio              As New ClsTCambio
'Public objAyudaTipoDocID            As New ClsTipoDocumentoID
'Public objAyudaTipoExistencia       As New ClsTipoExistencia
'Public objAyudaTomaInventario       As New ClsTomaInventario
'Public objAyudaUM                   As New ClsUM
Public objAyudaVale                 As New ClsVale
'Public objAyudaAlmacen              As New ClsAlmacen
'Public objAyudaCentroCosto          As New ClsCentroCosto
'Public objAyudaMarca                As New ClsMarca
'Public objAyudaUsuario              As New ClsUsuario
'Public objAyudaTarea                As New ClsTarea
'Public objAyudaTareaUsuario         As New ClsTareaUsuario
'Public objAyudaCuentaContable       As New ClsCuentaContable
'
'
'Public objSqlAyudaSolicitud         As New SqlClsSolicitud
'Public objSqlAyudaOrden             As New SqlClsOrden
'Public objSqlAyudaVale              As New SqlClsVale
'Public objSqlAyudaOrigen            As New SqlClsOrigen
'Public objSqlAyudaComprobante       As New SqlClsComprobante
'Public objSqlAyudaAlmacen           As New SqlClsAlmacen
'Public objSqlAyudaTipoExistencia    As New SqlClsTipoExistencia
'Public objSqlAyudaCentroCosto       As New SqlClsCentroCosto
'Public objSqlAyudaMarca             As New SqlClsMarca
'Public objSqlAyudaUM                As New SqlClsUM
'Public objSqlAyudaBienColor         As New SqlClsBienColor
'Public objSqlAyudaFamilia           As New SqlClsFamilia
'Public objSqlAyudaSubFamilia        As New SqlClsSubFamilia
'Public objSqlAyudaBien              As New SqlClsBien
'Public objSqlAyudaBienAlterno       As New SqlClsBienAlterno
'Public objSqlAyudaTipoDocID         As New SqlClsTipoDocumentoID
'Public objSqlAyudaDistrito          As New SqlClsDistrito
'Public objSqlAyudaFormaPago         As New SqlClsFormaPago
''Public objSqlAyudaTCambio           As New SqlClsTCambio
''Public objSqlAyudaProveedor         As New SqlClsProveedor
'Public objSqlAyudaCliente           As New SqlClsCliente
'Public objSqlAyudaUsuario           As New SqlClsUsuario
'Public objSqlAyudaTarea             As New SqlClsTarea
'Public objSqlAyudaTareaUsuario      As New SqlClsTareaUsuario
'Public objSqlAyudaTomaInventario    As New SqlClsTomaInventario


Public bolCorreoEnviado             As Boolean


Type ParametrosContables
    dblDetraccionPorcentaje As Double
    bolDetraccionImpMsj As Boolean
    strDetraccionMensaje As String
    strDetraccionCtaSoles As String
    strDetraccionCtaDolar As String
End Type

Public infoPlusParConta As ParametrosContables

Global Const HA = &H80000005 'CONSTANTE PARA EL COLOR ACTIVADO
                                 'DEL CONTROL TEXTBOX
Global Const DH = &HC0C0C0 'CONSTANTE PARA EL COLOR DESACTIVADO
                           'DEL CONTROL TEXTBOX
Global Const DF = &HC0FFC0    'CONSTANTE PARA EL COLOR DESACTIVADO
                           'DEL CONTROL TEXTBOX
Public wcontacnt        As String
Public wexcel           As String
Public dbcntcont        As DAO.Database
Public tbcntcont        As DAO.Recordset
Public mes              As String
Public wanno            As String
Public gcodppp          As String
Public dbbanco          As DAO.Database
Public tbtalon          As DAO.Recordset
Public dbempresa        As DAO.Database
Public Tbproveedor      As DAO.Recordset
Public dbcompras        As DAO.Database
Public TbCabRegis_new   As New ADODB.Recordset
Public TbCabRegis       As DAO.Recordset
'Public TbOfiRegis       As DAO.Recordset
Public TbOfiRegis       As New ADODB.Recordset
Public TbDetRegis       As New ADODB.Recordset
Public tbmes            As DAO.Recordset
Public dbcomtabla       As DAO.Database
Public tbcomtab         As DAO.Recordset
Public tbdepa           As DAO.Recordset
Public tbcambios        As DAO.Recordset
Public tbgastos         As DAO.Recordset
Public TbDocumento      As DAO.Recordset
Public dbconta          As DAO.Database
Public tbcf5pla         As DAO.Recordset
Public tbcf5costo       As DAO.Recordset
Public dbcentros        As DAO.Database
Public tbcentros        As DAO.Recordset
Public tbparametro1     As DAO.Recordset
Public Wnuevo           As Integer
Public wf1viscod        As String * 1
Public wocompra         As String
Public wf1tipdoc_asoc   As String
Public wf1inggasto      As String
Public wbancos          As String * 1
Public wf1traslado      As String
Public gctapag          As String * 1
Public wf1formatov      As String
Public wf1numera        As String * 1
Public wf1trasoc        As String
Public DbCtaPag         As DAO.Database
Public gorden_cs        As String
Public temp_deta        As DAO.Recordset
Public tbcabcta         As DAO.Recordset
Public tbdetcta         As DAO.Recordset
Public gcorre           As Double
Public gretenc          As Double
Public gfonavi          As Double
Public gnummov          As String
Public wingobra         As String
Public GCODGAS          As String
Public gcodcon          As String
Public gnomgas          As String
Public gcodcen          As String
Public wcentro          As String
Public worigen          As String
Public wctaigv          As String
Public wctaotros        As String
Public wredsuma         As String
Public wvisualiza_cod   As String  'visualizar la columna cod.,fab,ambos
Public wredresta        As String
Public wctaret          As String
Public wctafon          As String
Public llampro          As Integer
Public gcueppp          As String
Public gsegppp          As String
Public gmoneda          As String
Public gcodord          As Long
Public gcodprov         As String
Public gcoddepa         As String
Public gcodcta          As String * 3
Public wayudamov        As Integer
Public rucprv           As String
Public XITEM            As String
Public xgasto           As String
Public dbtempcomp       As DAO.Database
Public Tbtemp_regis     As DAO.Recordset
Public F4MES            As String
Public wf1formato       As String
Public xmodo            As String
Public xmes             As String
Public xmoneda          As String
Public xpro             As String
Public xtipo            As String
Public dbcomtab         As DAO.Database
Public grucprov         As String
Public tbregisdoc       As DAO.Recordset
Public tbregismov       As DAO.Recordset
Public gmonedacta       As String
Public gmondes          As String
Public dbplancta        As DAO.Database
Public tbplancta        As DAO.Recordset
Public dbtemp           As DAO.Database
Public NCTA             As Integer
Public gnomcen          As String
Public dbbancos         As DAO.Database
Public tbcta            As DAO.Recordset
Public TBBANCO          As DAO.Recordset
Public tbcheques        As DAO.Recordset
Public TBBF4TCO         As DAO.Recordset
Public dbmovis          As DAO.Database
Public nmovbanco        As String
Public xcaja            As Integer
Public xmovcaja         As Integer
Public xmescaja         As Integer
Public xctacaja         As Integer
Public gcodtip          As String
Public xtalon           As Integer
Public DbInventa        As DAO.Database
Public tbcabtmp         As DAO.Recordset
Public TbDetImport      As DAO.Recordset
Public TBPRODUCTO       As DAO.Recordset
Public TbCabImport      As DAO.Recordset
Public dvinvtemp        As DAO.Database
Public tbcf9saldo       As DAO.Recordset
Public tbcf9cta         As DAO.Recordset
Public tbmovcf4         As DAO.Recordset
Public tbcontrol        As DAO.Recordset
Public dbtemconta       As DAO.Database
Public dbmovconta       As DAO.Database
Public tbmovcf3         As DAO.Recordset
Public dbtabla          As DAO.Database
Public dbanalisis       As DAO.Database
Public wdcto            As String
Public gtipo            As String
Public TbDetOrdenes     As New ADODB.Recordset
Public TbCabOrdenes     As New ADODB.Recordset
Public gnomcon          As String
Public gcodnom          As String
Public wgastos          As String
Public wcodigo          As String
Public tbconta          As DAO.Recordset
Public wctacont         As String
Public wnomctacont      As String
Public LLAMADA          As String
Public des_grupo        As String
Public cod_grupo        As String
Public wf1renovacion    As String
Public wdestino         As String
Public wf1cnting        As String
Public wload_usuario    As String
Public wcodigos         As String
Public tbgrupos         As DAO.Recordset
Public tbmovim          As DAO.Recordset
Public tbcompra         As DAO.Recordset
Public TBCOMPraDOC      As DAO.Recordset
Public dbcompra         As DAO.Database
Public dbmovim          As DAO.Database
Public dbgrupos         As DAO.Database
Public WF2CODGAS        As String
Public wf1agente        As String
Public Sw_Graba_Registro As Boolean
Public ntc              As Double
Function CADENANUM(PNUM As Double, PMON As String, ptipo As String) As String
    
    Dim WDECIMAL    As String * 2
    Dim WENTERO     As String
    Dim WMONEDA     As String
    Dim WCADENA     As String
    Dim WCONT       As Integer
    Dim WSUBENT     As String
    'Global Numord As String

    WDECIMAL = right(Format$(PNUM, "#0.00"), 2)
    WENTERO = left(Format$(PNUM, "#0.00"), Len(Format$(PNUM, "#0.00")) - 3)
    'WMONEDA = IIf(PMON = "S", "NUEVOS SOLES", "DOLARES AMERICANOS")
    WMONEDA = IIf(PMON = "S", "SOLES", IIf(PMON = "E", "EUROS", "DOLARES AMERICANOS"))
    
    WCADENA = ""
    WCONT = 0
    WSUBENT = WENTERO
    Do While WCONT < Len(WENTERO)
        WSUBENT = right(WENTERO, Len(WENTERO) - WCONT)
        Select Case Len(WSUBENT)
        Case Is = 3, 6, 9: WCADENA = WCADENA & FCENTENA(Mid(WSUBENT, 1, 3))
        Case Is = 2, 5, 8
            If Val(Mid(WSUBENT, 1, 2)) > 15 Then
                WCADENA = WCADENA & FDECENA(Mid(WSUBENT, 1, 2))
            Else
                WCADENA = WCADENA & FUNIDAD(Mid(WSUBENT, 1, 2), Len(WSUBENT))
                WCONT = WCONT + 1
            End If
        Case Is = 1, 4, 7: WCADENA = WCADENA & FUNIDAD(Mid(WSUBENT, 1, 1), Len(WSUBENT))
        End Select
        WCONT = WCONT + 1
    Loop

    If ptipo = "*" Then '--- NO DEVUELVE LA MONEDA
        CADENANUM = WCADENA & " CON " & WDECIMAL & "/100 "
    Else
        CADENANUM = WCADENA & " CON " & WDECIMAL & "/100 " & WMONEDA
    End If

End Function

Function FCENTENA(PCAD As String)
    
    ReDim WUNI(10) As String
    Dim WSUBCAD     As String
    
    WUNI(0) = " "
    WUNI(1) = "CIENTO "
    WUNI(2) = "DOSCIENTOS "
    WUNI(3) = "TRESCIENTOS "
    WUNI(4) = "CUATROCIENTOS "
    WUNI(5) = "QUINIENTOS "
    WUNI(6) = "SEISCIENTOS "
    WUNI(7) = "SETECIENTOS "
    WUNI(8) = "OCHOCIENTOS "
    WUNI(9) = "NOVECIENTOS "

    If PCAD = "100" Then
        WSUBCAD = "CIEN"
    Else
        WSUBCAD = WUNI(Val(left(PCAD, 1)))
    End If

    FCENTENA = WSUBCAD

End Function

Function FDECENA(PCAD As String) As String
Dim WCAD        As String

    ReDim WUNI(10) As String
    Dim WSUBCAD     As String
    
    WCAD = left(PCAD, 2)
    WUNI(0) = " "
    WUNI(1) = "DIEZ "
    WUNI(2) = "VEINTE "
    WUNI(3) = "TREINTA "
    WUNI(4) = "CUARENTA "
    WUNI(5) = "CINCUENTA "
    WUNI(6) = "SESENTA "
    WUNI(7) = "SETENTA "
    WUNI(8) = "OCHENTA "
    WUNI(9) = "NOVENTA "

    If right(PCAD, 1) = 0 Then
        WSUBCAD = WUNI(Val(left(PCAD, 1)))
    Else
        WSUBCAD = WUNI(Val(left(PCAD, 1))) & "Y "
    End If

    FDECENA = WSUBCAD

End Function


Function FUNIDAD(PCAD As String, PLEN As Integer) As String

    ReDim WUNI(16) As String
    Dim WSUBCAD     As String
    
    WUNI(0) = " "
    WUNI(1) = "UN "
    WUNI(2) = "DOS "
    WUNI(3) = "TRES "
    WUNI(4) = "CUATRO "
    WUNI(5) = "CINCO "
    WUNI(6) = "SEIS "
    WUNI(7) = "SIETE "
    WUNI(8) = "OCHO "
    WUNI(9) = "NUEVE "
    WUNI(10) = "DIEZ "
    WUNI(11) = "ONCE "
    WUNI(12) = "DOCE "
    WUNI(13) = "TRECE "
    WUNI(14) = "CATORCE "
    WUNI(15) = "QUINCE "
           
    Select Case PLEN
        Case Is = 1, 2: WSUBCAD = WUNI(Val(PCAD))
        Case Is = 4, 5: WSUBCAD = WUNI(Val(PCAD)) & "MIL "
        Case Is = 7, 8: WSUBCAD = WUNI(Val(PCAD)) & IIf(PCAD = "1", "MILLON", "MILLONES ")
    End Select

    FUNIDAD = WSUBCAD

End Function



Public Function dev_mes(mes)
Dim nmes    As Integer
Dim cmes    As String
   
   nmes = Val(mes)
   Select Case mes
      Case 0: cmes = "Apertura"
      Case 1: cmes = "Enero"
      Case 2: cmes = "Febrero"
      Case 3: cmes = "Marzo"
      Case 4: cmes = "Abril"
      Case 5: cmes = "Mayo"
      Case 6: cmes = "Junio"
      Case 7: cmes = "Julio"
      Case 8: cmes = "Agosto"
      Case 9: cmes = "Setiembre"
      Case 10: cmes = "Octubre"
      Case 11: cmes = "Noviembre"
      Case 12: cmes = "Diciembre"
      Case 13: cmes = "Cierre 1"
      Case 14: cmes = "Cierre 2"
   End Select
   dev_mes = cmes

End Function

Public Sub Actualiza_Log(CadSql As String, conexion As String)
    On Error Resume Next
    
    Dim NomFile As String, StrLine As String
    
    Rem SK ADD:
    Dim NomFileUsuario As String, StrLineUsuario As String, strCadSqlUsuario As String
    Dim intNumSlot As Integer
    
    NomFile = wrutabancos & "\Control_Plus_Logistica_" & dev_mes(Month(Date)) & "_" & Year(Date) & ".log"
    
    Rem SK ADD:
    NomFileUsuario = wrutatemp & "\Control_Plus_Logistica_" & ComputerName & "_" & dev_mes(Month(Date)) & "_" & Year(Date) & ".log"
    strCadSqlUsuario = CadSql
    
    intNumSlot = FreeFile
    
    If (InStr(UCase(conexion), "DB_TABLA") > 0 Or InStr(UCase(conexion), "DB_BANCOS") > 0) Then
        '------------------------------------------------------------------------------------------------------------
        Close #intNumSlot
    
        Open Trim(NomFile) For Append As #intNumSlot
        
        '***genera sql
        CadSql = UCase(Replace(CadSql, "'", "|"))
        
        StrLine = "<Fecha Hora:" & Format(Now, "MM/DD/YYYY HH:MM:SS AM/PM") & ">" & "<Usuario:" & wusuario & ">" & "<Pc:" & ComputerName & ">" & "<Sentencia:" & CadSql & ">"
        
        Print #intNumSlot, StrLine
        
        '------------------------------------------------------------------------------------------------------------
        
        Rem SK ADD:
        If Dir(NomFileUsuario, vbArchive) = vbNullString Then
            Close #intNumSlot
            
            Open Trim(NomFileUsuario) For Append As #intNumSlot
            
            StrLineUsuario = "FECHA|USUARIO|NOMBREPC|SENTENCIAEJECUTADA"
            
            Print #intNumSlot, StrLineUsuario
        End If
        
        Close #intNumSlot
        
        Open Trim(NomFileUsuario) For Append As #intNumSlot
        
        StrLineUsuario = Format(Now, "MM/DD/YYYY HH:MM:SS AM/PM") & "|" & wusuario & "|" & ComputerName & "|" & strCadSqlUsuario
        
        Print #intNumSlot, StrLineUsuario
        
        Close #intNumSlot
    End If
    
    Exit Sub
End Sub

Public Sub abrirCnnDbBancos()
    If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
    
    cnn_dbbancos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos.mdb;Persist Security Info=False"
End Sub

Public Sub abrirCnnnDbBancosBackup()
    If Dir(wrutabancos & "\db_bancos_backup.mdb", vbArchive) <> "" Then
        StrConexDbBancos_backup = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos_backup.mdb;Persist Security Info=False"
        cconex_dbbancos_backup = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_bancos_backup.mdb" & ";Persist Security Info=False"
            
        If cnn_dbbancos_backup.State = 1 Then cnn_dbbancos_backup.Close
        
        cnn_dbbancos_backup.Open cconex_dbbancos_backup
    End If
End Sub

Public Sub abrirCnContaTabla()
    Set cnDBContaTabla = New ADODB.Connection
    
    If cnDBContaTabla.State = 1 Then cnDBContaTabla.Close
    
    cnDBContaTabla.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutaconta & "\db_tabla.mdb;Persist Security Info=False"
End Sub

Public Sub abrirCnTemporal()
    Set cnDBTemp = New ADODB.Connection
    
    If cnDBTemp.State = 1 Then cnDBTemp.Close
    
    cnDBTemp.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnDBTemp.CursorLocation = adUseClient
    cnDBTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0" & _
                                ";Data Source= " & wrutatemp & "Templus.mdb" & _
                                ";Persist Security Info=False"
    
    cnDBTemp.Open
End Sub

Public Sub abrirCnBdCPlus()
    On Error GoTo errAbrirCnBdCPlus
    
    Set cnBdCPlus = New ADODB.Connection
    
    strCadenaConexioBdCPlus = sGetINI(App.Path & "\configServidorSQLCliente.ini", "ConfigServidorSQLCliente", "CadenaConexionBdCPlus", "l")
    
    If cnBdCPlus.State = 1 Then cnBdCPlus.Close
    
    cnBdCPlus.Open strCadenaConexioBdCPlus
    
    Exit Sub
errAbrirCnBdCPlus:
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & " - AbrirCnBdCPlus"
    
    Err.Clear
End Sub


Public Function traerCampo(tabla As String, campo As String, campoCom As String, valor As String, Optional condicion As String) As String
    Dim cad As String
    Dim rst As New Recordset
    If IsDate(valor) Then
        cad = "select " & campo & " from " & tabla & " where CVDATE(" & campoCom & ") = '" & valor & "' " & condicion
    'ElseIf IsNumeric(valor) Then
        'cad = "select " & campo & " from " & tabla & " where " & campoCom & " = " & valor & " " & condicion
    Else
        cad = "select " & campo & " from " & tabla & " where " & campoCom & " = '" & valor & "' " & condicion
    End If
    If tabla = "srutas" Then
        cconex_control = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RUTAS.MDB" & ";Persist Security Info=False"
        If cnn_control.State = 1 Then cnn_control.Close
        cnn_control.Open cconex_control
        rst.Open cad, cnn_control, adOpenForwardOnly, adLockReadOnly
    Else
        rst.Open cad, cnn_dbbancos, adOpenForwardOnly, adLockReadOnly
    End If
    traerCampo = ""
    If Not rst.EOF And Not IsNull(rst.Fields(0)) Then traerCampo = rst.Fields(0)
End Function

Public Function OBTIENE_CORRELA_XPAGAR(pconexion As ADODB.Connection)
Dim rspagcab            As New ADODB.Recordset
Dim ncorre              As Double

    If rspagcab.State = adStateOpen Then rspagcab.Close
    rspagcab.Open "SELECT CORRELA FROM PAG_DCTO ORDER BY CORRELA DESC", pconexion, adOpenDynamic, adLockOptimistic
    If Not rspagcab.EOF Then
        rspagcab.MoveFirst
        ncorre = rspagcab.Fields("correla") + 1
    Else
        ncorre = 1
    End If
    rspagcab.Close
    OBTIENE_CORRELA_XPAGAR = ncorre

End Function

Public Sub TRANS_CTASXPAGAR_NEW(ptipo As String, pvia As String, pcorrela As Double, ptipdocu As String, pserdoc As String, pdocum As String, pfecha As Date, pruc As String, pcodigo As String, pmoneda As String, ptipcam As Double, ptotal As Double, pdebhab As String, prefer As String, pfechavenc As Date, pcentro As String, pnomcodigo As String, pconexion As String, preg_com As String, pannorc As String, pordcompra As String)
'On Error GoTo error_trans
'Dim CnTmp As New ADODB.Connection
'Dim amovs_cab(0 To 21)  As a_grabacion
'Dim sw_llena            As Boolean
'Dim ctipo_graba         As String * 1
'Dim cwhere              As String
'Dim ncorrela            As Double
'Dim rspagcab            As New ADODB.Recordset
'    '***abre conexion
'    If CnTmp.State = 1 Then CnTmp.Close
'    CnTmp.Open pconexion
'    '****************
'    sw_llena = False
'    If pcorrela > 0 Then
'        '--------- MODIFICACION
'        If rspagcab.State = adStateOpen Then rspagcab.Close
'        rspagcab.Open "SELECT * FROM PAG_DCTO WHERE CORRELA=" & pcorrela & "", CnTmp, adOpenDynamic, adLockOptimistic
'        If Not rspagcab.EOF Then
'            If Val(Format(rspagcab.Fields("TOTAL"), "0.00")) = Val(Format(rspagcab.Fields("SALDO"), "0.00")) Then
'                sw_llena = True
'                ctipo_graba = "M"
'                cwhere = "CORRELA = " & pcorrela & ""
'                ncorrela = pcorrela
'                CorrelaPagDcto = ncorrela
'            Else
'                sw_llena = False
'                'MsgBox "El documento ya ha sido aplicado.", vbInformation, "CONTROL Plus!"
'
'                MsgBox "El documento ya ha sido aplicado.", vbInformation, wnomcia
'                ncorrela = pcorrela
'                CorrelaPagDcto = ncorrela
'            End If
'        End If
'        rspagcab.Close
'    Else
'        '--------- NUEVO
'        sw_llena = True
'        ctipo_graba = "A"
'        cwhere = ""
'        ncorrela = OBTIENE_CORRELA_XPAGAR(CnTmp)
'        CorrelaPagDcto = ncorrela
'    End If
'    'wcorrela = ncorrela
'    If sw_llena = True Then
'        amovs_cab(0).campo = "VIA_INGR": amovs_cab(0).valor = pvia: amovs_cab(0).Tipo = "T"
'        amovs_cab(1).campo = "CORRELA": amovs_cab(1).valor = ncorrela: amovs_cab(1).Tipo = "N"
'        If Len(Trim(pserdoc)) > 0 Then
'            amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = ptipdocu & pserdoc & "/" & pdocum: amovs_cab(2).Tipo = "T"
'        Else
'            amovs_cab(2).campo = "NRO_COMP": amovs_cab(2).valor = ptipdocu & pdocum: amovs_cab(2).Tipo = "T"
'        End If
'
'        amovs_cab(3).campo = "FCH_COMP": amovs_cab(3).valor = pfecha: amovs_cab(3).Tipo = "F"
'        amovs_cab(4).campo = "PROVEEDOR": amovs_cab(4).valor = pcodigo: amovs_cab(4).Tipo = "T"
'        amovs_cab(5).campo = "PROVEEDORO": amovs_cab(5).valor = pcodigo: amovs_cab(5).Tipo = "T"
'        amovs_cab(6).campo = "MONEDA": amovs_cab(6).valor = pmoneda: amovs_cab(6).Tipo = "T"
'        amovs_cab(7).campo = "MONEDAO": amovs_cab(7).valor = pmoneda: amovs_cab(7).Tipo = "T"
'        amovs_cab(8).campo = "TCAMBIO": amovs_cab(8).valor = ptipcam: amovs_cab(8).Tipo = "N"
'        amovs_cab(9).campo = "TCAMBIOO": amovs_cab(9).valor = ptipcam: amovs_cab(9).Tipo = "N"
'        amovs_cab(10).campo = "TOTAL": amovs_cab(10).valor = ptotal: amovs_cab(10).Tipo = "N"
'        amovs_cab(11).campo = "TOTALO": amovs_cab(11).valor = ptotal: amovs_cab(11).Tipo = "N"
'        amovs_cab(12).campo = "SALDO": amovs_cab(12).valor = ptotal: amovs_cab(12).Tipo = "N"
'        amovs_cab(13).campo = "DEB_HAB": amovs_cab(13).valor = pdebhab: amovs_cab(13).Tipo = "T"
'        amovs_cab(14).campo = "REFERENCIA": amovs_cab(14).valor = prefer: amovs_cab(14).Tipo = "T"
'        amovs_cab(15).campo = "FCH_VCTO": amovs_cab(15).valor = pfechavenc: amovs_cab(15).Tipo = "F"
'        amovs_cab(16).campo = "F4CENTRO": amovs_cab(16).valor = pcentro: amovs_cab(16).Tipo = "T"
'        amovs_cab(17).campo = "NOMPROV": amovs_cab(17).valor = pnomcodigo: amovs_cab(17).Tipo = "T"
'        amovs_cab(18).campo = "REG_COM": amovs_cab(18).valor = Mid(preg_com, 5, 9): amovs_cab(18).Tipo = "T"
'        amovs_cab(19).campo = "RUC": amovs_cab(19).valor = pruc: amovs_cab(19).Tipo = "T"
'        amovs_cab(20).campo = "F4ANNORC": amovs_cab(20).valor = pannorc: amovs_cab(20).Tipo = "T"
'        amovs_cab(21).campo = "F4OCOMPRA": amovs_cab(21).valor = pordcompra: amovs_cab(21).Tipo = "T"
'
'        GRABA_REGISTRO amovs_cab(), "PAG_DCTO", ctipo_graba, 21, pconexion, cwhere
'
'        csql = "UPDATE REGISDOC SET F4CORRELA=" & ncorrela & " WHERE F4MESMOV='" & (Mid(preg_com, 1, 6)) & "' AND F4NUMMOV='" & Format(right(preg_com, 7), "0000000") & "'"
'        CnTmp.Execute (csql)
'
'
'    End If
'    '****cierra conex
'    If CnTmp.State = 1 Then CnTmp.Close
'    Set CnTmp = Nothing
'    'validando registro de compras
'    Exit Sub
'error_trans:
'    MsgBox "Se ha producido el sgte. error : " & Err.Description, 16, wnomcia
'    Resume
'    Exit Sub
End Sub

 


Sub TRANSCTACTE(ptipo As String, pvia As String, pcorrela As Double, ptipdocu As String, pserdoc As String, pdocum As String, pfecha As Date, pruc As String, pcodigo As String, pmoneda As String, ptipcam As Double, ptotal As Double, pdebhab As String, prefer As String, pfechavenc As Date, pcentro As String, pnomcodigo As String, pconexion As ADODB.Connection)
On Error GoTo error_trans
Dim amovs_cab(0 To 17)  As a_grabacion
Dim sw_llena            As Boolean
Dim ctipo_graba         As String * 1
Dim cwhere              As String
Dim ncorrela            As Double

    sw_llena = False
    If pcorrela > 0 Then
        '--------- MODIFICACION
        RSCTA_DCTO.Open "SELECT * FROM CTA_DCTO WHERE TIPO='" & ptipo & "' AND CORRELA=" & pcorrela & "", pconexion, adOpenDynamic, adLockOptimistic
        If Not RSCTA_DCTO.EOF Then
            If Val(Format(RSCTA_DCTO.Fields("TOTAL"), "0.00")) = Val(Format(RSCTA_DCTO.Fields("SALDO"), "0.00")) Then
                sw_llena = True
                ctipo_graba = "M"
                cwhere = "TIPO = '" & ptipo & "' AND CORRELA = " & pcorrela & ""
                ncorrela = pcorrela
            Else
                sw_llena = False
                MsgBox "El documento ya ha sido aplicado.", vbInformation, "Atención"
            End If
        End If
        RSCTA_DCTO.Close
    Else
        '--------- NUEVO
        sw_llena = True
        ctipo_graba = "A"
        cwhere = ""
        ncorrela = OBTIENE_CORRELA(pconexion)
    End If
    
    If sw_llena = True Then
        amovs_cab(0).campo = "TIPO": amovs_cab(0).valor = ptipo: amovs_cab(0).Tipo = "T"
        amovs_cab(1).campo = "VIA_INGR": amovs_cab(1).valor = pvia: amovs_cab(1).Tipo = "T"
        amovs_cab(2).campo = "CORRELA": amovs_cab(2).valor = ncorrela: amovs_cab(2).Tipo = "N"
        amovs_cab(3).campo = "TIPDOCU": amovs_cab(3).valor = ptipdocu: amovs_cab(3).Tipo = "T"
        amovs_cab(4).campo = "SERDOC": amovs_cab(4).valor = pserdoc: amovs_cab(4).Tipo = "T"
        amovs_cab(5).campo = "DOCUM": amovs_cab(5).valor = pdocum: amovs_cab(5).Tipo = "T"
        amovs_cab(6).campo = "FECHA": amovs_cab(6).valor = pfecha: amovs_cab(6).Tipo = "F"
        amovs_cab(7).campo = "RUC": amovs_cab(7).valor = pruc: amovs_cab(7).Tipo = "T"
        amovs_cab(8).campo = "CODIGO": amovs_cab(8).valor = pcodigo: amovs_cab(8).Tipo = "T"
        amovs_cab(9).campo = "MONEDA": amovs_cab(9).valor = pmoneda: amovs_cab(9).Tipo = "T"
        amovs_cab(10).campo = "TIPCAM": amovs_cab(10).valor = ptipcam: amovs_cab(10).Tipo = "N"
        amovs_cab(11).campo = "TOTAL": amovs_cab(11).valor = ptotal: amovs_cab(11).Tipo = "N"
        amovs_cab(12).campo = "SALDO": amovs_cab(12).valor = ptotal: amovs_cab(12).Tipo = "N"
        amovs_cab(13).campo = "DEB_HAB": amovs_cab(13).valor = pdebhab: amovs_cab(13).Tipo = "T"
        amovs_cab(14).campo = "REFERENCIA": amovs_cab(14).valor = prefer: amovs_cab(14).Tipo = "T"
        amovs_cab(15).campo = "FECHA_VCTO": amovs_cab(15).valor = pfechavenc: amovs_cab(15).Tipo = "F"
        amovs_cab(16).campo = "F4CENTRO": amovs_cab(16).valor = pcentro: amovs_cab(16).Tipo = "T"
        amovs_cab(17).campo = "NOMCODIGO": amovs_cab(17).valor = pnomcodigo: amovs_cab(17).Tipo = "T"
        
        GRABA_REGISTRO_logistica amovs_cab(), "CTA_DCTO", ctipo_graba, 17, pconexion, cwhere
        
    End If
        
    Exit Sub
    
error_trans:
    MsgBox "Se ha producido el sgte. error : " & Err.Description, 16, "Error: " & Err.Number
    Exit Sub
    
End Sub

Public Sub CREATETABLE_N(ptabla As String, pcadena As String, pconexion As ADODB.Connection)
On Error Resume Next
Dim query As String
    
    query = "Create table " & ptabla & " " & pcadena
    pconexion.Execute (query)
    ''AlmacenaQuery_sql query, pconexion
End Sub

Public Sub DELETEREC_LOG(ptabla As String, pconexion As ADODB.Connection)
On Error Resume Next
    pconexion.Execute ("DELETE * FROM " & ptabla)
    ''AlmacenaQuery_sql "DELETE * FROM " & ptabla, pconexion
End Sub
Public Sub DELETEREC_BANCOS(ptabla As String, pconexion As String, pwhere As String)
On Error GoTo ErrorDelete

Dim Cn_Tmp As New ADODB.Connection
Dim CadSql As String
    Cn_Tmp.Open pconexion
    
    
    CadSql = "DELETE * FROM " & ptabla
    
    If Len(Trim(pwhere)) > 0 Then
        CadSql = CadSql & " where " & pwhere
    End If
    
    Cn_Tmp.Execute CadSql
    
    Actualiza_Log CadSql, pconexion
    
    
    If Cn_Tmp.State = 1 Then Cn_Tmp.Close
    Set Cn_Tmp = Nothing
    
    
    Exit Sub

ErrorDelete:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Exit Sub
End Sub



Public Sub CREATEDATABASE_N(pruta As String, pnombase As String)
Dim DBName      As String
Dim cat         As ADOX.Catalog
Dim Tbl         As ADOX.Table

    DBName = pruta & pnombase
    Set cat = New ADOX.Catalog
    Set Tbl = New ADOX.Table
    
    If Len(Dir$(DBName)) Then
        Kill DBName
    End If
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ";"
    cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ";"
            
End Sub

Public Sub GRABA_REGISTRO_logistica(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As ADODB.Connection, pwhere As String)
    On Error GoTo CapturaError
    Dim i           As Integer
    Dim ccampos     As String
    Dim cvalores    As String
    Dim csql        As String
    Dim prSol As New Recordset
    
    sw_GRABA_REGISTRO_logistica = False
    
    ccampos = "": cvalores = ""
    For i = 0 To pcantidad
        If ptipo = "A" Then
            If Len(ccampos) = 0 Then
                ccampos = ccampos & parreglo(i).campo
                If parreglo(i).Tipo = "T" Then
                    If parreglo(i).valor <> vbNullString Then
                        cvalores = "'" & parreglo(i).valor & "'"
                    Else
                        cvalores = "Null"
                    End If
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = "Null"
                    Else
                        cvalores = parreglo(i).valor
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = "'" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = "'" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = "Null"
                        Else
                            cvalores = "CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            Else
                ccampos = ccampos & "," & parreglo(i).campo
                If parreglo(i).Tipo = "T" Then
                    If parreglo(i).valor <> vbNullString Then
                        cvalores = cvalores & ",'" & parreglo(i).valor & "'"
                    Else
                        cvalores = cvalores & ",Null"
                    End If
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & ",Null"
                    Else
                        cvalores = cvalores & "," & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & ",'" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & ",'" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & ",Null"
                        Else
                            cvalores = cvalores & ",CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
        If ptipo = "M" Then
            If Len(Trim(cvalores)) = 0 Then
                If parreglo(i).Tipo = "T" Then
                    If parreglo(i).valor <> vbNullString Then
                        cvalores = cvalores & parreglo(i).campo & "='" & parreglo(i).valor & "'"
                    Else
                        cvalores = cvalores & parreglo(i).campo & "=Null"
                    End If
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & parreglo(i).campo & "=Null"
                    Else
                        cvalores = cvalores & parreglo(i).campo & "=" & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & parreglo(i).campo & "='" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & parreglo(i).campo & "=" & Format(parreglo(i).valor, "yyyy-mm-dd")
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & parreglo(i).campo & "=Null"
                        Else
                            cvalores = cvalores & parreglo(i).campo & "=CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            Else
                If parreglo(i).Tipo = "T" Then
                    If parreglo(i).valor <> vbNullString Then
                        cvalores = cvalores & "," & parreglo(i).campo & "='" & parreglo(i).valor & "'"
                    Else
                        cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                    End If
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                    Else
                        cvalores = cvalores & "," & parreglo(i).campo & "=" & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                        Else
                            cvalores = cvalores & "," & parreglo(i).campo & "=CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ptipo = "A" Then
        With objAyudaVale
            .SQLSelectAlter = .NumeroVale
            
            '.TipoVale = left(cnumvale, 1)
            '.CodigoAlmacen = IIf(Not sw_ingreso, Trim(txtalmacen.Text), Trim(txtalmacendes.Text))
            '.NumeroVale = IIf(Not sw_ingreso, Trim(txtnumero.Text), Trim(txtdestino.Text))
            '.Fecha = abofecha.value
            
            If .verificarExistencia Then
                .NumeroVale = .generarNumeroVale
                '.NumeroVale = "I-101108"
                
                cvalores = Replace(cvalores, .SQLSelectAlter, .NumeroVale, 1)
            End If
        End With
        
        csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
        csql = csql & cvalores & ")"
    End If
  
    If ptipo = "M" Then
        csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
    End If
    
    pconexion.Execute csql
    sw_GRABA_REGISTRO_logistica = True
'    'AlmacenaQuery_sql csql, pconexion
    Actualiza_Log csql, pconexion.ConnectionString
    'AlmacenaQuery_sql csql, pconexion
    
    
    Exit Sub
    Resume
CapturaError:

    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Atencion"
    Resume Next
End Sub

Public Sub GRABA_REGISTRO(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As String, pwhere As String)
On Error GoTo Error_Graba_Registro
Dim CnSave As New ADODB.Connection
Dim i           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
    
    Sw_Graba_Registro = False
    
    CnSave.Open pconexion
    
    ccampos = "": cvalores = ""
    For i = 0 To pcantidad
        If ptipo = "A" Then
            If Len(ccampos) = 0 Then
                ccampos = ccampos & parreglo(i).campo
                If parreglo(i).Tipo = "T" Then
                    cvalores = "'" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = "Null"
                    Else
                        cvalores = 0 & parreglo(i).valor
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = "'" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = "'" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = "Null"
                        Else
                            cvalores = "CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            Else
                ccampos = ccampos & "," & parreglo(i).campo
                If parreglo(i).Tipo = "T" Then
                    cvalores = cvalores & ",'" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & ",Null"
                    Else
                        cvalores = cvalores & "," & 0 & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & ",'" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & ",'" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & ",Null"
                        Else
                            cvalores = cvalores & ",CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
        If ptipo = "M" Then
            If Len(Trim(cvalores)) = 0 Then
                If parreglo(i).Tipo = "@" Then
                    cvalores = cvalores & parreglo(i).campo & "=" & parreglo(i).valor & " "
                End If
                If parreglo(i).Tipo = "T" Then
                    cvalores = cvalores & parreglo(i).campo & "='" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & parreglo(i).campo & "=Null"
                    Else
                        cvalores = cvalores & parreglo(i).campo & "=" & 0 & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & parreglo(i).campo & "='" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & parreglo(i).campo & "=" & Format(parreglo(i).valor, "yyyy-mm-dd")
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & parreglo(i).campo & "=Null"
                        Else
                            cvalores = cvalores & parreglo(i).campo & "=CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            Else
                If parreglo(i).Tipo = "@" Then
                    cvalores = cvalores & "," & parreglo(i).campo & "=" & parreglo(i).valor & " "
                End If
                If parreglo(i).Tipo = "T" Then
                    cvalores = cvalores & "," & parreglo(i).campo & "='" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                    Else
                        cvalores = cvalores & "," & parreglo(i).campo & "=" & 0 & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                        Else
                            cvalores = cvalores & "," & parreglo(i).campo & "=CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ptipo = "A" Then
        csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
        csql = csql & cvalores & ")"
    End If

    
    If ptipo = "M" Then
        csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
    End If
    
    CnSave.Execute csql
    
    
    Actualiza_Log csql, pconexion
    Sw_Graba_Registro = True
    
    If CnSave.State = 1 Then CnSave.Close
    Set CnSave = Nothing
    
    Exit Sub
    
Error_Graba_Registro:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Select Case Err.Number
    Case 3704
        If CnSave.State = 0 Then CnSave.Open
        Resume
    Case -2147467259
        If MsgBox("¿Desea volver a intentar?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
            For J = 0 To 10000
                If CnSave.State = 1 Then CnSave.Close
                Set CnSave = Nothing
            Next
            'If CnSave.State = 1 Then CnSave.Close
            CnSave.Open pconexion
            Resume
        Else
            Exit Sub
        End If
    End Select
    Sw_Graba_Registro = False
    
    Exit Sub
    
End Sub


Public Sub GRABA_REGISTRO_noenvia(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As ADODB.Connection, pwhere As String)
On Error GoTo CapturaError
Dim i           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
Dim prSol As New Recordset
    sw_GRABA_REGISTRO_logistica = False
    
    ccampos = "": cvalores = ""
    For i = 0 To pcantidad
        If ptipo = "A" Then
            If Len(ccampos) = 0 Then
                ccampos = ccampos & parreglo(i).campo
                If parreglo(i).Tipo = "T" Then
                    cvalores = "'" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = "Null"
                    Else
                        cvalores = 0 & parreglo(i).valor
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = "'" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = "'" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = "Null"
                        Else
                            cvalores = "CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            Else
                ccampos = ccampos & "," & parreglo(i).campo
                If parreglo(i).Tipo = "T" Then
                    cvalores = cvalores & ",'" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & ",Null"
                    Else
                        cvalores = cvalores & "," & 0 & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & ",'" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & ",'" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & ",Null"
                        Else
                            cvalores = cvalores & ",CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
        If ptipo = "M" Then
            If Len(Trim(cvalores)) = 0 Then
                If parreglo(i).Tipo = "T" Then
                    cvalores = cvalores & parreglo(i).campo & "='" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & parreglo(i).campo & "=Null"
                    Else
                        cvalores = cvalores & parreglo(i).campo & "=" & 0 & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & parreglo(i).campo & "='" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & parreglo(i).campo & "=" & Format(parreglo(i).valor, "yyyy-mm-dd")
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & parreglo(i).campo & "=Null"
                        Else
                            cvalores = cvalores & parreglo(i).campo & "=CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            Else
                If parreglo(i).Tipo = "T" Then
                    cvalores = cvalores & "," & parreglo(i).campo & "='" & parreglo(i).valor & "'"
                End If
                If parreglo(i).Tipo = "N" Then
                    If Not IsNumeric(parreglo(i).valor) Then
                        cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                    Else
                        cvalores = cvalores & "," & parreglo(i).campo & "=" & 0 & parreglo(i).valor & ""
                    End If
                End If
                If parreglo(i).Tipo = "H" Then
                    cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parreglo(i).valor, "HH:MM:SS") & "'"
                End If
                If parreglo(i).Tipo = "F" Then
                    If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                        cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parreglo(i).valor, "yyyy-mm-dd") & "'"
                    Else
                        If Not IsDate(parreglo(i).valor) Then
                            cvalores = cvalores & "," & parreglo(i).campo & "=Null"
                        Else
                            cvalores = cvalores & "," & parreglo(i).campo & "=CVDATE('" & parreglo(i).valor & "')"
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ptipo = "A" Then
        csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
        csql = csql & cvalores & ")"
    End If
  
    If ptipo = "M" Then
        csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
    End If
    
    pconexion.Execute csql
    sw_GRABA_REGISTRO_logistica = True
    Actualiza_Log csql, pconexion.ConnectionString
'    'AlmacenaQuery_sql csql, pconexion
    ''AlmacenaQuery_sql csql, pconexion
    
    
    Exit Sub
CapturaError:

    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Atencion"
    
End Sub


'Public Sub 'AlmacenaQuery_sql(ByVal sql As String, ConeccionAdodb As ADODB.Connection)
'On Error GoTo CapturaError
'If InStr(UCase(ConeccionAdodb), "DB_BANCOS.MDB") > 0 Or InStr(UCase(ConeccionAdodb), "DB_TABLA.MDB") > 0 Then
'    Dim cnEnvia As New ADODB.Connection
'    If cnEnvia.State = 1 Then cnEnvia.Close
'    cnEnvia.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & wrutabancos & "\db_envia.mdb;Persist Security Info=False"
'    sql = Replace(sql, "'", "|")
'
'    csql = "insert into querys (wquery) values('" & sql & "')"
'    cnEnvia.Execute csql
'    'AlmacenaQuery_sql csql, cnEnvia
'
'    If cnEnvia.State = 1 Then cnEnvia.Close
'    Set cnEnvia = Nothing
'End If
'Exit Sub
'CapturaError:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Atención"
'    Exit Sub
'End Sub



Public Sub GRABA_REGISTRO_logistica_DET(parreglo() As a_grabacion, ptabla As String, ptipo As String, pcantidad As Integer, pconexion As ADODB.Connection, pwhere As String, parr_det() As Variant, pnumfilas As Integer, pvalores As String, pmes As String, pgraba_saldo As String)

Dim i           As Integer
Dim ccampos     As String
Dim cvalores    As String
Dim csql        As String
Dim nfila       As Integer
    
        
    sw_GRABA_REGISTRO_logistica = False
    
    ccampos = "": cvalores = ""
    For nfila = 0 To pnumfilas
        For i = 0 To pcantidad
            If Mid(pvalores, i + 1, 1) = "1" Then '--- EL CAMPO ESTA ACTIVO PARA GRABARLO
                If ptipo = "A" Then
                    If Len(ccampos) = 0 Then
                        ccampos = ccampos & parreglo(i).campo
                        If parreglo(i).Tipo = "T" Then
                            cvalores = "'" & parr_det(i, nfila) & "'"
                        End If
                        If parreglo(i).Tipo = "N" Then
                            cvalores = parr_det(i, nfila)
                        End If
                        If parreglo(i).Tipo = "H" Then
                            cvalores = "'" & Format(parr_det(i, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(i).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = "'" & Format(parr_det(i, nfila), "yyyy-mm-dd") & "'"
                            Else
                                cvalores = "CVDATE('" & parr_det(i, nfila) & "')"
                            End If
                        End If
                    Else
                        ccampos = ccampos & "," & parreglo(i).campo
                        If parreglo(i).Tipo = "T" Then
                            cvalores = cvalores & ",'" & parr_det(i, nfila) & "'"
                        End If
                        If parreglo(i).Tipo = "N" Then
                            cvalores = cvalores & "," & parr_det(i, nfila) & ""
                        End If
                        If parreglo(i).Tipo = "H" Then
                            cvalores = cvalores & ",'" & Format(parr_det(i, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(i).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & ",'" & Format(parr_det(i, nfila), "yyyy-mm-dd") & "'"
                            Else
                                cvalores = cvalores & ",CVDATE('" & parr_det(i, nfila) & "')"
                            End If
                        End If
                    End If
                End If
                If ptipo = "M" Then
                    If Len(Trim(cvalores)) = 0 Then
                        If parreglo(i).Tipo = "T" Then
                            cvalores = cvalores & parreglo(i).campo & "='" & parr_det(i, nfila) & "'"
                        End If
                        If parreglo(i).Tipo = "N" Then
                            cvalores = cvalores & parreglo(i).campo & "=" & parr_det(i, nfila) & ""
                        End If
                        If parreglo(i).Tipo = "H" Then
                            cvalores = cvalores & parreglo(i).campo & "='" & Format(parr_det(i, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(i).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & parreglo(i).campo & "=" & Format(parr_det(i, nfila), "yyyy-mm-dd")
                            Else
                                cvalores = cvalores & parreglo(i).campo & "=CVDATE('" & parr_det(i, nfila) & "')"
                            End If
                        End If
                    Else
                        If parreglo(i).Tipo = "T" Then
                            cvalores = cvalores & "," & parreglo(i).campo & "='" & parr_det(i, nfila) & "'"
                        End If
                        If parreglo(i).Tipo = "N" Then
                            cvalores = cvalores & "," & parreglo(i).campo & "=" & parr_det(i, nfila) & ""
                        End If
                        If parreglo(i).Tipo = "H" Then
                            cvalores = cvalores & "," & parreglo(i).campo & "='" & Format(parr_det(i, nfila), "HH:MM:SS") & "'"
                        End If
                        If parreglo(i).Tipo = "F" Then
                            If ctipoadm_bd = "M" Then '----- ADM. B.D. MYSQL
                                cvalores = cvalores & "," & parreglo(i).campo & "=" & Format(parr_det(i, nfila), "yyyy-mm-dd")
                            Else
                                cvalores = cvalores & "," & parreglo(i).campo & "=CVDATE('" & parr_det(i, nfila) & "')"
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        If ptipo = "A" Then
            cvalores = Replace(cvalores, cnumvale, objAyudaVale.NumeroVale, 1)
            
            csql = "INSERT INTO " & ptabla & " (" & ccampos & ") VALUES ("
            csql = csql & cvalores & ")"
            pconexion.Execute csql
        '    If pgraba_saldo = "*" Then '---- ACTUALIZA CANTIDAD Y PESO
        '        GRABA_SALDO parr_det(1, nfila), parr_det(4, nfila), parr_det(5, nfila), pmes, "I", pconexion
        '    End If
        '    If pgraba_saldo = "A" Then '---- ACTUALIZA CANTIDAD Y COSTO X ALMACEN
        '        GRABA_SALDO_ALM parr_det(1, nfila), Format(parr_det(2, nfila) & "", "0.000"), parr_det(5, nfila), pmes, wtipoguia, pconexion, parr_det(6, nfila), parr_det(10, nfila), "S"
        '    End If
        End If
        
        If ptipo = "M" Then
            csql = "UPDATE " & ptabla & " SET " & cvalores & " WHERE " & pwhere
            pconexion.Execute csql
        End If
        'AlmacenaQuery_sql csql, pconexion
        Actualiza_Log csql, pconexion.ConnectionString
        ''AlmacenaQuery_sql csql, pconexion
        
        ccampos = "": cvalores = ""
    Next
    
    sw_GRABA_REGISTRO_logistica = True
End Sub

Public Sub GRABA_SALDO(pcodprod As Variant, pcantidad As Variant, ppeso As Variant, pmes As String, pdestino As String, pconexion As ADODB.Connection)
Dim csql    As String

    rsif5placli.Open "SELECT * FROM IF5PLACLI WHERE F5CODPRO = '" & pcodprod & "'", pconexion, adOpenDynamic, adLockOptimistic
    If Not rsif5placli.EOF Then
        If pdestino = "I" Then
            csql = "UPDATE IF5PLACLI SET F5DEBM" & pmes & " = F5DEBM" & pmes & " + " & Val("" & pcantidad) & ",F5INGPESO" & pmes & " =  F5INGPESO" & pmes & " + " & Val("" & ppeso) & " WHERE F5CODPRO = '" & pcodprod & "'"
        Else
            csql = "UPDATE IF5PLACLI SET F5DEBM" & pmes & " = F5DEBM" & pmes & " - " & Val("" & pcantidad) & ",F5INGPESO" & pmes & " =  F5INGPESO" & pmes & " - " & Val("" & ppeso) & " WHERE F5CODPRO = '" & pcodprod & "'"
        End If
        pconexion.Execute csql
        ''AlmacenaQuery_sql csql, pconexion
    End If
    rsif5placli.Close

End Sub

Public Sub ELIMINA_BD_N(pruta As String, pnombase As String)
Dim DBName      As String

    DBName = pruta & "\" & pnombase
    'Kill DBName
    
End Sub

Public Sub GRABA_SALDO_ALM(pcodprod As Variant, pcantidad As Variant, psoles As Variant, pmes As String, pdestino As String, pconexion As ADODB.Connection, palmacen As Variant, pdolares As Variant, poperacion As String)
Dim csql    As String

    rsif6alma.Open "SELECT F5CODPRO FROM IF6ALMA WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'", pconexion, adOpenDynamic, adLockOptimistic
    If Not rsif6alma.EOF Then
        If pdestino = "I" Then
            If poperacion = "S" Then
                csql = "UPDATE IF6ALMA SET F5DEBM" & pmes & " = F5DEBM" & pmes & " + " & Val("" & pcantidad) & _
                ",F5ING" & pmes & " =  F5ING" & pmes & " + " & Val("" & psoles) & _
                ",F5INGD" & pmes & " =  F5INGD" & pmes & " + " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT + " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT + " & Val("" & pcantidad) & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                ''AlmacenaQuery_sql csql, pconexion
                
            Else   '------ poperacion = "R"
                csql = "UPDATE IF6ALMA SET F5DEBM" & pmes & " = F5DEBM" & pmes & " - " & Val("" & pcantidad) & _
                ",F5ING" & pmes & " =  F5ING" & pmes & " - " & Val("" & psoles) & _
                ",F5INGD" & pmes & " =  F5INGD" & pmes & " - " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT - " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT - " & pcantidad & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                ''AlmacenaQuery_sql csql, pconexion
            End If
        Else
            If poperacion = "S" Then
                csql = "UPDATE IF6ALMA SET F5HABM" & pmes & " = F5HABM" & pmes & " + " & Val("" & pcantidad) & _
                ",F5SAL" & pmes & " =  F5SAL" & pmes & " + " & Val("" & psoles) & _
                ",F5SALD" & pmes & " =  F5SALD" & pmes & " + " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT - " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT - " & pcantidad & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                ''AlmacenaQuery_sql csql, pconexion
            Else   '------ poperacion = "R"
                csql = "UPDATE IF6ALMA SET F5HABM" & pmes & " = F5HABM" & pmes & " - " & Val("" & pcantidad) & _
                ",F5SAL" & pmes & " =  F5SAL" & pmes & " - " & Val("" & psoles) & _
                ",F5SALD" & pmes & " =  F5SALD" & pmes & " - " & Val("" & pdolares) & _
                ",F6STOCKACT = F6STOCKACT + " & Val(pcantidad) & _
                " WHERE F5CODPRO = '" & pcodprod & "' AND F2CODALM = '" & palmacen & "'"
                csql = "UPDATE IF5PLA SET F5STOCKACT=F5STOCKACT + " & Val("" & pcantidad) & " WHERE F5CODPRO='" & pcodprod & "'"
                pconexion.Execute csql
                ''AlmacenaQuery_sql csql, pconexion
            End If
            
        End If
        pconexion.Execute csql
        ''AlmacenaQuery_sql csql, pconexion
    End If
    rsif6alma.Close

End Sub

Public Function OBTIENE_CORRELA(pconexion As ADODB.Connection)
Dim ncorre      As Double

    RSCTA_DCTO.Open "SELECT CORRELA FROM CTA_DCTO ORDER BY CORRELA DESC", pconexion, adOpenDynamic, adLockOptimistic
    If Not RSCTA_DCTO.EOF Then
        RSCTA_DCTO.MoveFirst
        ncorre = RSCTA_DCTO.Fields("correla") + 1
    Else
        ncorre = 1
    End If
    RSCTA_DCTO.Close
    OBTIENE_CORRELA = ncorre

End Function

Public Sub ACUMULA_PRODUCTOS(prutatemp As String, pnombase As String, psql1 As String, psql2 As String, pconexion As ADODB.Connection, pnomtabla As String, pconexion_temp As ADODB.Connection, pcantidad As Double, ptipoalm As String, palmacen As String, PCodOri As String, pmedida As Double, ptime As String)
Dim rsdet       As New ADODB.Recordset
Dim rstemp      As New ADODB.Recordset
Dim rsgrupos    As New ADODB.Recordset
Dim CadSql      As String
Dim calmacen    As String

    If sw_creabd = True Then
        DELETEREC_LOG pnomtabla, pconexion_temp
        sw_creabd = False
    End If
        
    rsdet.Open psql2, pconexion, adOpenDynamic, adLockOptimistic
       
    If Not rsdet.EOF Then
        rsdet.MoveFirst
        Do While Not rsdet.EOF
            calmacen = ""
            If ptipoalm = "*" Then
                calmacen = Format(rsdet.Fields("F3GRUPOINS"), "00")
            Else
                calmacen = palmacen
            End If
            If pmedida > pcantidad Then
                sql = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & rsdet.Fields("F3GRUPOINS") & "','" & calmacen & "','" & rsdet.Fields("F3CODPROINS") & "', " & (Val(Format((pcantidad / pmedida) * Val(rsdet.Fields("F3CANTIDAD")), "0.00000"))) & ",'" & PCodOri & "')"
            ElseIf pmedida < pcantidad Then
                sql = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & rsdet.Fields("F3GRUPOINS") & "','" & calmacen & "','" & rsdet.Fields("F3CODPROINS") & "', " & (Val(Format((pcantidad * rsdet.Fields("F3CANTIDAD")) / pmedida, "0.00000"))) & ",'" & PCodOri & "')"
            Else
                sql = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & rsdet.Fields("F3GRUPOINS") & "','" & calmacen & "','" & rsdet.Fields("F3CODPROINS") & "', " & Val(Format(rsdet.Fields("F3CANTIDAD"), "0.000")) & ",'" & PCodOri & "')"
            End If
            pconexion_temp.Execute sql
            ''AlmacenaQuery_sql sql, pconexion_temp
            rsdet.MoveNext
        Loop
    Else
        MsgBox "Producto " & PCodOri & " no tiene Receta", vbInformation, "AVISO"
        wFORMULA = 1
        'SQL = "INSERT INTO " & pnomtabla & " (GRUPO,ALMACEN,CODPROD,CANTIDAD,CODORI) VALUES('" & dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value & "','" & calmacen & "','" & dxDBGrid1.Columns.ColumnByFieldName("CODIGO").Value & "', '" & Val(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value) & "','" & PCodOri & "')"
        'pconexion_temp.Execute SQL
        Exit Sub
    End If
      
    rsdet.Close

End Sub

Public Function valida(Tipo, car, Optional Texto, Optional mays)
Select Case Tipo
    Case 1  'Todo
        If Not IsMissing(mays) Then
            valida = Asc(UCase(Chr(car)))
        Else
            valida = car
        End If
    Case 2  'solo letras
        If Not ((car >= 65 And car <= 90) Or (car >= 97 And car <= 122) Or car = 8 Or car = 32) Then
            valida = 0
        Else
            If Not IsMissing(mays) Then
                valida = Asc(UCase(Chr(car)))
            Else
                valida = car
            End If
        End If
    Case 3  'numeros sin punto decimal
        If Not (car >= 48 And car <= 57 Or car = 8) Then
            valida = 0
        Else
            valida = car
        End If
    Case 4  'numeros con punto decimal
        If Not (car >= 48 And car <= 57 Or car = 8) Then
            If car = 46 Then
                existedec = InStr(1, Texto, ".")
                If existedec > 0 Then
                    valida = 0
                Else
                    valida = car
                End If
            Else
                valida = 0
            End If
        Else
            valida = car
        End If
End Select
End Function

'Public Function CalculaExistencia(almacen, prod, FECHA)
'Dim rst As New ADODB.Recordset
'Dim rstS As New ADODB.Recordset
'cad = ""
'If Not (almacen = "") Then
'    cad = " and f3vales.f2codalm='" & almacen & "'"
'End If
'
'If ctipoadm_bd = "A" Then
''    'SQL = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr " _
''    '& "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) " _
''    '& "WHERE CVDATE(IF4VALES.F4FECVAL)<=CVDATE('" & fecha & "') And f5codpro='" & prod & "'and f3vales.f2codalm='" & almacen & "'" ' & cad
'    SQL = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) Where CVDATE(if3vales.F4FECVAL)<=CVDATE('" & FECHA & "') AND if3vales.F5CODPRO ='" & Trim(prod) & "' and if3vales.F2CODALM='" & almacen & "'"
'Else
''   ' SQL = "SELECT Sum(If(Left(IF4VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0)) AS ING, Sum(Iff(Left(IF4VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0)) AS EGR "
''   ' SQL = SQL & "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) "
''   ' SQL = SQL & "WHERE IF4VALES.F4FECVAL<='" & fecha & "' And f5codpro='" & prod & "'" & cad
'
'    SQL = "SELECT Left([IF3VALES].[F4NUMVAL],1)='I' AS ING1, IF3VALES.F2CODALM, IF3VALES.F5CODPRO, Sum(IF3VALES.F3CANPRO) AS SumaDeF3CANPRO"
'    SQL = SQL + " FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)"
'   SQL = SQL + " GROUP BY Left([IF3VALES].[F4NUMVAL],1)='I', IF3VALES.F2CODALM, IF3VALES.F5CODPRO"
'    SQL = SQL + " HAVING (((IF3VALES.F2CODALM)='" & almacen & "') AND ((IF3VALES.F5CODPRO)='" & prod & "'))"
'
'End If
'
'If rst.State = adStateOpen Then rst.Close
'rst.Open UCase(SQL), cnn_dbbancos, adOpenStatic, adLockOptimistic
'
'If Not rst.EOF Then
'    nstock = Val("" & rst!ing1) - Val("" & rst!Egr)
'    CalculaExistencia = nstock
'Else
'    CalculaExistencia = -1
'End If
'rst.Close
'End Function
Public Function CalculaExistencia(almacen, prod, Fecha)
Dim rst As New ADODB.Recordset

cad = ""
If Not (almacen = "") Then
    cad = " and if3vales.f2codalm='" & almacen & "'"
End If
If ctipoadm_bd = "M" Then
    sql = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) Where CVDATE(if3vales.F4FECVAL)<=CVDATE('" & Fecha & "') AND if3vales.F5CODPRO ='" & Trim(prod) & "' and if3vales.F2CODALM='" & almacen & "'"
Else
    sql = "SELECT Sum(IIf(Left(if4vales.f4numval,1)='I',f3canpro)) AS ing, Sum(IIf(Left(if4vales.f4numval,1)='S',f3canpro)) AS egr " _
    & "FROM IF4VALES INNER JOIN if3vales ON (IF4VALES.F4NUMVAL = if3vales.F4NUMVAL) AND (IF4VALES.F2CODALM = if3vales.F2CODALM) " _
    & "WHERE CVDATE(IF4VALES.F4FECVAL)<=CVDATE('" & Fecha & "') And f5codpro='" & prod & "'" & cad
End If
If rst.State = adStateOpen Then rst.Close
rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
If Not rst.EOF Then
    nstock = Val("" & rst("ing")) - Val("" & rst("egr"))
    CalculaExistencia = nstock
Else
    CalculaExistencia = -1
End If
rst.Close
End Function


Public Function VALIDA_FPAGO(pfpago As String)
Dim sw      As Boolean

    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select F2DESPAG from ef2forpag where f2forpag='" & Trim(pfpago) & "'", cnn_dbbancos  'cnn_dbbancos
    If rst.EOF = False Then
        wnompag = rst!F2DESPAG & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_FPAGO = sw
        
End Function

Public Function VALIDA_CLIENTE(pcodcli As String)
Dim sw      As Boolean
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select f2CODcli,f2nomcli,F2NEWRUC,f2dircli,f2forpag,f2lista_asig from ef2clientes where f2codcli='" & pcodcli & "' OR f2NEWRUC='" & pcodcli & "'", cnn_dbbancos
    If rst.EOF = False Then
        wcodcli = "" & rst!F2CODCLI
        wnomcli = "" & rst!F2nomcli
        wruccli = "" & rst!F2NEWRUC
        WDIRCLI = "" & rst!F2DIRCLI
        wforpag = "" & rst!F2FORPAG
        nnumlista = Val("" & rst.Fields("f2lista_asig") & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CLIENTE = sw
        
End Function

Public Function VALIDA_CLIENTE_V(pcodcli As String, pcodven As String)
Dim sw      As Boolean
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select f2CODcli,f2nomcli,F2NEWRUC,f2dircli,f2forpag,f2lista_asig from ef2clientes where (f2codcli='" & pcodcli & "' OR f2newruc='" & pcodcli & "') AND (F2CODVEN='" & pcodven & "')", cnn_dbbancos
    If rst.EOF = False Then
        wcodcli = "" & rst!F2CODCLI
        wnomcli = "" & rst!F2nomcli
        wruccli = "" & rst!F2NEWRUC
        WDIRCLI = "" & rst!F2DIRCLI
        wforpag = "" & rst!F2FORPAG
        nnumlista = Val(rst.Fields("f2lista_asig") & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CLIENTE_V = sw
        
End Function

Public Function VALIDA_RESPONSABLE(pcodres As String)
Dim sw      As Boolean

    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "Select * from EF2VENDEDORES where F2CODVEN='" & pcodres & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Rs.EOF = False Then
        wnomven = Rs.Fields("F2NOMVEN") & ""
        wfacmin_nac = Val(Format("" & Rs.Fields("F2FACNAC_MIN"), "0.00"))
        wfacmin_imp = Val(Format("" & Rs.Fields("F2FACIMP_MIN"), "0.00"))
        sw = True
    Else
        sw = False
    End If
    Rs.Close
    VALIDA_RESPONSABLE = sw
        
End Function

Public Function VALIDA_ORIGEN(PCodOri As String)
Dim sw      As Boolean
Set rst = New ADODB.Recordset
    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from Sf1Origenes where F1CodOri='" & Trim(PCodOri) & "'", cnn_dbbancos
    If Not rst.EOF Then
        wcodori = Trim(rst!F1CODORI & "")
        wnomori = Trim(rst!F1NOMORI & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_ORIGEN = sw

End Function
Public Function VALIDA_ALMACENO(pcodalm As String)
Dim sw  As Boolean
Set Rs = New ADODB.Recordset

    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "Select F2NOMALM from ef2almacenes where f2codalm='" & Trim(pcodalm) & "'", cnn_dbbancos
    If Rs.EOF = False Then
        WNomPar = Rs!F2NOMALM & ""
        sw = True
    Else
        sw = False
    End If
    Rs.Close
    VALIDA_ALMACENO = sw
        
End Function

Public Sub GRABAR_ACTUALIZACIONES(Fmes As Integer, Alma As String, cod As String)
Dim amovs_arr(0 To 7) As a_grabacion

amovs_arr(0).campo = "F5DEBM" & Format(Fmes, "00"): amovs_arr(0).valor = Debm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5ING" & Format(Fmes, "00"): amovs_arr(1).valor = ing: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5INGD" & Format(Fmes, "00"): amovs_arr(2).valor = ingd: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"
amovs_arr(4).campo = "F6STOCKLOG": amovs_arr(4).valor = StockLog: amovs_arr(4).Tipo = "N"
amovs_arr(5).campo = "F6FECULT": amovs_arr(5).valor = FecUlt: amovs_arr(5).Tipo = "F"
amovs_arr(6).campo = "F5COSPRO": amovs_arr(6).valor = Cospro: amovs_arr(6).Tipo = "N"
amovs_arr(7).campo = "F5COSPROD": amovs_arr(7).valor = Cosprod: amovs_arr(7).Tipo = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 7, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Public Sub GRABAR_ACTUALIZACIONES1(Fmes As Integer, Alma As String, cod As String)

Dim amovs_arr(0 To 6) As a_grabacion

amovs_arr(0).campo = "F5HABM" & Format(Fmes, "00"): amovs_arr(0).valor = Habm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5SAL" & Format(Fmes, "00"): amovs_arr(1).valor = sal: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5SALD" & Format(Fmes, "00"): amovs_arr(2).valor = sald: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"
amovs_arr(4).campo = "F6FECULT": amovs_arr(4).valor = CVDate(FecUlt): amovs_arr(4).Tipo = "F"
amovs_arr(5).campo = "F5COSPRO": amovs_arr(5).valor = Cospro: amovs_arr(5).Tipo = "N"
amovs_arr(6).campo = "F5COSPROD": amovs_arr(6).valor = Cosprod: amovs_arr(6).Tipo = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 6, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub Actualizar_Almacenes(pcodalm As String, pcodpro As String, pcanpro As Double, PFECMOV As Variant, psoles As Double, pdolares As Double, ptipmov As String, pprecio As Double)
Dim csql, SSQL As String
    
Set RsProducto = New ADODB.Recordset
Set RsMovAlmacen = New ADODB.Recordset
'''''''''''CAMBIOS EN SQL
SSQL = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsMovAlmacen.EOF Then
    wmes = Format(Month(CVDate(PFECMOV)), "00")
    Debm = 0#: stockact = 0#: ing = 0#: ingd = 0#
    Habm = 0#: StockLog = 0#: sal = 0#: sald = 0#
    If ptipmov = "S" Then
        Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) - pcanpro
        ing = RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")) - psoles
        ingd = RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")) - pdolares
        GRABAR_ACTUALIZAALMA wmes, pcodalm, pcodpro
    Else
    '''''''''CORREGIR MES
        Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
        StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
        sal = RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")) - psoles
        sald = RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")) - pdolares
        GRABAR_ACTUALIZAALMA1 wmes, pcodalm, pcodpro
    End If

End If
End Sub

Sub GRABAR_ACTUALIZAALMA(Fmes As String, Alma As String, cod As String)
Dim amovs_arr(0 To 3) As a_grabacion

amovs_arr(0).campo = "F5DEBM" & Format(Fmes, "00"): amovs_arr(0).valor = Debm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5ING" & Format(Fmes, "00"): amovs_arr(1).valor = ing: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5INGD" & Format(Fmes, "00"): amovs_arr(2).valor = ingd: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"

'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 3, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub GRABAR_ACTUALIZAALMA1(Fmes As String, Alma As String, cod As String)
Dim amovs_arr(0 To 4) As a_grabacion

amovs_arr(0).campo = "F5HABM" & Format(Fmes, "00"): amovs_arr(0).valor = Habm: amovs_arr(0).Tipo = "N"
amovs_arr(1).campo = "F5SAL" & Format(Fmes, "00"): amovs_arr(1).valor = sal: amovs_arr(1).Tipo = "N"
amovs_arr(2).campo = "F5SALD" & Format(Fmes, "00"): amovs_arr(2).valor = sald: amovs_arr(2).Tipo = "N"
amovs_arr(3).campo = "F6STOCKACT": amovs_arr(3).valor = stockact: amovs_arr(3).Tipo = "N"
amovs_arr(4).campo = "F6STOCKLOG": amovs_arr(4).valor = StockLog: amovs_arr(4).Tipo = "N"


'------- ACTUALIZAR STOCKS
GRABA_REGISTRO_logistica amovs_arr(), "IF6ALMA", "M", 4, cnn_dbbancos, "F2CODALM = '" & Alma & "' AND F5CODPRO = '" & cod & "'"

End Sub

Sub ImprimendoV(pcodalm As String, PNUMVAL As String, pcosto As Integer)
    
Dim ITEM, Fila As Integer
    
ITEM = 1
Fila = 1

Set RsStockCab = New ADODB.Recordset
Set RsStockDet = New ADODB.Recordset
Set RsProducto = New ADODB.Recordset
Set RSCONSULTA = New ADODB.Recordset

If RsStockCab.State = adStateOpen Then RsStockCab.Close
RsStockCab.Open "SELECT * FROM IF4VALES WHERE F2CODALM = '" & pcodalm & "' AND F4NUMVAL = '" & PNUMVAL & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsStockCab.EOF Then
    
    Printer.ScaleMode = 4
    TituloV PNUMVAL

    If RsStockDet.State = adStateOpen Then RsStockDet.Close
    RsStockDet.Open "SELECT * FROM IF3VALES WHERE F2CODALM = '" & pcodalm & "' AND F4NUMVAL = '" & PNUMVAL & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsStockDet.EOF Then
        Fila = Fila + 6
        Do While Not RsStockDet.EOF
            If RsProducto.State = adStateOpen Then RsProducto.Close
            RsProducto.Open "SELECT * FROM IF5PLA WHERE F5CODPRO ='" & RsStockDet.Fields("F5CODPRO") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            If Not RsProducto.EOF Then
                writexy Trim(RsStockDet.Fields("F5codpro")), Fila, 1, 0
                writexy left(RsProducto.Fields("F5nompro"), 65), Fila, 11, 0
                writexy Val(Format(RsStockDet.Fields("F3canpro"), "#0.000")), Fila, 63, 2
                writexy Trim("" & RsProducto.Fields("f7codmed")), Fila, 80, 0
                Fila = Fila + 1
                If RsProducto.Fields("f5series") = "1" Then
                    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
                    sql = "select * from if3series where f2codalm='" & pcodalm & "' and f4numval='" & PNUMVAL & "' and f5codpro='" & Trim(RsProducto.Fields("f5codpro")) & "' order by f3numser"
                    RSCONSULTA.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not RSCONSULTA.EOF Then
                        RSCONSULTA.MoveFirst
                        Do While Not RSCONSULTA.EOF
                            writexy ITEM & ".-  E.S.N. ==>" & RSCONSULTA.Fields("F3numser"), Fila, 16, 0
                            Fila = Fila + 1
                            RSCONSULTA.MoveNext
                            If Fila >= 60 Then
                                Printer.NewPage
                                TituloV PNUMVAL
                                Fila = 18
                            End If
                            ITEM = ITEM + 1
                        Loop
                    End If
                    Fila = Fila + 1
                End If
                ITEM = 1
                If Fila >= 60 Then
                   Printer.NewPage
                   TituloV PNUMVAL
                   Fila = 18
                End If
                RsProducto.MoveNext
            End If
            RsStockDet.MoveNext
        Loop
    End If

    Printer.Line (60, Fila + 6)-(88, Fila + 6)
    writexy "Almacén", Fila + 6, 68, 0
    
    Printer.Line (5, Fila + 6)-(33, Fila + 6)
    writexy "Solicitado Por", Fila + 6, 12, 0
    Printer.EndDoc
    
End If

End Sub

Sub TituloV(PNUMVAL As String)
Dim Fila As Integer
Set RsAlmacenes = New ADODB.Recordset
Set RsProveedor = New ADODB.Recordset
Set RsOrigen = New ADODB.Recordset
Set RsPartida = New ADODB.Recordset
    

    If RsOrigen.State = adStateOpen Then RsOrigen.Close
    RsOrigen.Open "SELECT * FROM SF1ORIGENES WHERE F1CODORI = '" & RsStockCab.Fields("F1CODORI") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    CabeceraV
    Printer.FontBold = True
    Printer.FontSize = 15
    writexy IIf(RsOrigen.Fields("F1TIPMOV") = "I", "VALE DE INGRESO", "VALE DE SALIDA"), 6, 34, 0
    Printer.FontBold = False
    Printer.FontSize = 9
    
    Fila = 11
    Printer.FontBold = True

    If Trim(nomcentro) <> "" Then
        writexy "C.COSTO:", 10, 5, 0
        writexy Trim(nomcentro), 10, 15, 0
    End If

    writexy "Nº VALE:", Fila, 5, 0
    writexy PNUMVAL, Fila, 15, 0
    
    writexy "Almacén:", Fila, 60, 0
    
    If Not IsNull(RsStockCab.Fields("F2codalm")) Then
        If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
        RsAlmacenes.Open "SELECT * FROM EF2ALMACENES WHERE F2CODALM = '" & RsStockCab.Fields("F2CODALM") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsAlmacenes.EOF Then
            writexy Trim(RsAlmacenes.Fields("f2codalm") & " - " & Mid(RsAlmacenes.Fields("F2nomalm"), 1, 21)), Fila, 70, 0
        End If
        
    End If

    Fila = Fila + 1
    
    If Not IsNull(RsStockCab.Fields("F2codprov")) Then
        If RsProveedor.State = adStateOpen Then RsProveedor.Close
        RsProveedor.Open "SELECT * FROM EF2PROVEEDORES WHERE F2NEWRUC = '" & RsStockCab.Fields("F2CODPROV") & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not RsProveedor.EOF Then
            writexy "Proveedor:", Fila, 5, 0
            writexy Trim(RsProveedor.Fields("f2codprov") & " - " & RsProveedor.Fields("F2nomprov")), Fila, 15, 0
        End If
        
    End If
    
    If Not IsNull(RsStockCab.Fields("F2codpar")) Then
        If RsStockCab.Fields("f1codori") = "XT0" Or RsStockCab.Fields("f1codori") = "XT1" Then
            If RsAlmacenes.State = adStateOpen Then RsAlmacenes.Close
            RsAlmacenes.Open "SELECT * FROM EF2ALMACENES", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsAlmacenes.Find "F2codalm = '" & RsStockCab.Fields("F2codpar") & "'"
            If Not RsAlmacenes.EOF Then
                If RsStockCab.Fields("f1codori") = "XT0" Then
                   writexy "ORIGEN  : ", Fila, 60, 0
                Else
                   writexy "DESTINO : ", Fila, 60, 0
                End If
                writexy UCase(RsAlmacenes.Fields("f2codalm") & "-" & RsAlmacenes.Fields("f2nomalm")), Fila, 70, 0
            End If
        Else
            
            If RsPartida.State = adStateOpen Then RsPartida.Close
            RsPartida.Open "SELECT * FROM EF2PARTIDAS", cnn_dbbancos, adOpenDynamic, adLockOptimistic
            RsPartida.Find "F2codpar = '" & RsStockCab.Fields("F2codpar") & "'"
            If Not RsPartida.EOF Then
                writexy "DESTINO : ", Fila, 60, 0
                writexy UCase(RsPartida.Fields("f2nompar")), Fila, 70, 0
            End If
        End If
    Else
        If Val("" & RsStockCab.Fields("F4numord")) > 0 Then
            writexy "O/Compra : ", Fila, 60, 0
            writexy Format(RsStockCab.Fields("F4numord"), "000000"), Fila, 70, 0
        End If
    End If


    Printer.FontBold = False
    writexy "Concepto:", Fila + 1, 5, 0
    writexy Trim(RsOrigen.Fields("F1nomori")), Fila + 1, 15, 0

    writexy "Documento:", Fila + 1, 60, 0
    writexy Trim(RsStockCab.Fields("F1coddoc") & " (" & RsStockCab.Fields("F4numdoc") & ") "), Fila + 1, 70, 0
    
    writexy "Fecha:", Fila + 2, 5, 0
    writexy Trim(Trim(cmes(Month(CVDate(RsStockCab.Fields("F4fecval"))))) & " " & Format(Day(CVDate(RsStockCab.Fields("F4fecval"))), "##") & ", de " & Year(CVDate(RsStockCab.Fields("F4fecval")))), Fila + 2, 15, 0
    
    Printer.FontBold = True
    Printer.Line (1, Fila + 3)-(90, Fila + 3)
    writexy "Código", Fila + 4, 1, 0
    writexy "Artículo", Fila + 4, 11, 0
    writexy "Cantidad", Fila + 4, 63, 0
    writexy "Uni.", Fila + 4, 80, 0
    Printer.Line (1, Fila + 5)-(90, Fila + 5)
    Printer.FontBold = False
    
End Sub

Sub CabeceraV()
Dim csql As String
Set RsParain = New ADODB.Recordset

csql = "SELECT * FROM SF1PARAIN WHERE F1CODEMP = '" & wempresa & "'"
If RsParain.State = adStateOpen Then RsParain.Close
RsParain.Open csql, cnn_control, adOpenDynamic, adLockOptimistic
  
Printer.ScaleMode = 4
Printer.FontName = "Courier New" 'Printer.Fonts (RsParain.Fields("F1FONNAM"))
Printer.FontSize = 12
Printer.FontBold = True
writexy Trim("" & RsParain.Fields("f1nomemp")), 2, 1, 0
Printer.FontSize = 8
Printer.FontBold = False

writexy "Fecha: ", 2, 80, 0
writexy Format(Now, "dd/mm/yyyy"), 2, 88, 0
writexy "Intersys - Inventario", 3, 1, 0
writexy "Página: ", 3, 80, 0
writexy Format(Printer.Page, "###00"), 3, 88, 0

End Sub

Sub Reactualiza_Almacenes(pcodalm As String, pcodpro As String, pcanpro As Double, PFECMOV As Variant, psoles As Double, pdolares As Double, ptipmov As String, pprecio As Double)
Dim SSQL, csql As String
Set RsProducto = New ADODB.Recordset
Set RsMovAlmacen = New ADODB.Recordset

stockact = 0#: Debm = 0#: ing = 0#: ingd = 0#
StockLog = 0#: Habm = 0#: sal = 0#: sald = 0#
sql = "SELECT * FROM IF6ALMA WHERE F2CODALM = '" & pcodalm & "' AND F5CODPRO = '" & pcodpro & "'"
If RsMovAlmacen.State = adStateOpen Then RsMovAlmacen.Close
RsMovAlmacen.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not RsMovAlmacen.EOF Then
    wmes = Format(Month(CVDate(PFECMOV)), "00")
    SSQL = "SELECT F5STOCKACT FROM IF5PLA WHERE F5CODPRO = '" & pcodpro & "'"
    If RsProducto.State = adStateOpen Then RsProducto.Close
    RsProducto.Open SSQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not RsProducto.EOF Then
       If ptipmov = "S" Then
          stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) + pcanpro
       Else
          stockact = Val(Format(RsProducto.Fields("F5STOCKACT"), "#0.00")) - pcanpro
       End If
       csql = "UPDATE IF5PLA SET F5STOCKACT = " & stockact & " WHERE F5CODPRO = '" & pcodpro & "'"
       cnn_dbbancos.Execute (csql)
       ''AlmacenaQuery_sql csql, cnn_dbbancos
       
    End If
    RsProducto.Close
    
    If ptipmov = "S" Then
        Debm = Val(Format(RsMovAlmacen.Fields("F5DEBM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) - pcanpro
        ing = Val(Format(RsMovAlmacen.Fields("F5ING" & Format(wmes, "00")), "0.00")) - psoles
        ingd = Val(Format(RsMovAlmacen.Fields("F5INGD" & Format(wmes, "00")), "0.00")) - pdolares
        GRABAR_ACTUALIZAALMA wmes, pcodalm, pcodpro
    Else
        Habm = Val(Format(RsMovAlmacen.Fields("F5HABM" & Format(wmes, "00")), "#0.00")) - pcanpro
        stockact = Val(Format(RsMovAlmacen.Fields("F6STOCKACT"), "#0.00")) + pcanpro
        StockLog = Val(Format(RsMovAlmacen.Fields("F6STOCKLOG"), "#0.00")) + pcanpro
        sal = Val(Format(RsMovAlmacen.Fields("F5SAL" & Format(wmes, "00")), "0.00")) - psoles
        sald = Val(Format(RsMovAlmacen.Fields("F5SALD" & Format(wmes, "00")), "0.00")) - pdolares
        GRABAR_ACTUALIZAALMA1 wmes, pcodalm, pcodpro

    End If
    
End If
End Sub

Public Sub BASE_TEMPORAL(Base As String)
Dim con As String
Set Temp = New ADODB.Connection

con = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & Base & "; Persist Security Info=False"
Temp.Open con

End Sub
Public Function VALIDA_PROVEEDOR(pcodprov As String)
Dim RsProveedor As New ADODB.Recordset
On Error Resume Next
Dim sw      As Boolean
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "SELECT F2CODPROV, F2NOMPROV, F2NEWRUC " & _
                "FROM EF2PROVEEDORES " & _
                "where f2codprov='" & pcodprov & "' OR F2NEWRUC='" & pcodprov & "'", cnn_dbbancos, 3, 1
    If Rs.EOF = False Then
        wcodprov = "" & Rs!F2CODPROV
        wnomprov = "" & Rs!F2NOMPROV
        wrucprov = "" & Rs!F2NEWRUC
        sw = True
    Else
        sw = False
        wcodprov = ""
        wnomprov = ""
        wrucprov = ""
    End If
    Rs.Close
    VALIDA_PROVEEDOR = sw
        
End Function

Public Function ObtenerCampo(tabla As String, campo As String, campoCom As String, valor As String, TipoDeComparacion As String, ConexionDeBaseDeDatos As ADODB.Connection) As String
Dim cad As String
Dim rst As New Recordset
If TipoDeComparacion = "F" Then
    cad = "select " & campo & " from " & tabla & " where CVDATE(" & campoCom & ") = '" & valor & "' " & condicion
ElseIf TipoDeComparacion = "T" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = '" & valor & "' " & condicion
ElseIf TipoDeComparacion = "N" Then
    cad = "select " & campo & " from " & tabla & " where " & campoCom & " = " & valor & " " & condicion
End If
If rst.State = 1 Then rst.Close
rst.Open cad, ConexionDeBaseDeDatos, adOpenForwardOnly, adLockReadOnly
ObtenerCampo = ""
If Not rst.EOF And Not IsNull(rst.Fields(0)) Then ObtenerCampo = rst.Fields(0)
End Function


Public Function SeleccionaEnComboLeft(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
For i = 0 To NombreDelCombo.ListCount - 1
    If right(NombreDelCombo.List(i), Len(Trim(DatoBuscado))) = DatoBuscado Then
        NombreDelCombo.ListIndex = i
        Exit For
    Else
    End If
Next
End Function

Public Function SeleccionaEnComboRight(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
For i = 0 To NombreDelCombo.ListCount - 1
    If UCase(right(NombreDelCombo.List(i), Len(Trim(DatoBuscado)))) = UCase(DatoBuscado) Then
        NombreDelCombo.ListIndex = i
        Exit For
    Else
    End If
Next
End Function

Public Function SeleccionaEnCombo(ByVal DatoBuscado As String, ByVal NombreDelCombo As ComboBox)
For i = 0 To NombreDelCombo.ListCount - 1
    If left(NombreDelCombo.List(i), Len(Trim(DatoBuscado))) = DatoBuscado Then
        NombreDelCombo.ListIndex = i
        Exit For
    Else
    End If
Next
End Function

Public Function CargaRsT(ByVal SentenciaSql As String, ByVal ConexionUsada As ADODB.Connection) As ADODB.Recordset
Dim rst As New ADODB.Recordset
If rst.State = 1 Then rst.Close
rst.Open SentenciaSql, ConexionUsada, 3, 1
Set CargaRsT_DbTabla = rst
End Function
Public Function VALIDA_VENDEDOR(pvendedor As String)
Dim sw      As Boolean
    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from EF2USERS where F2CODUSER='" & Trim(pvendedor) & "' AND F2VENDEDOR='*'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If rst.EOF = False Then
        wnomven = rst!F2NOMUSER & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_VENDEDOR = sw
End Function

Public Function VALIDA_COBRADOR(pvendedor As String)
Dim sw      As Boolean

    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from EF2USERS where F2CODUSER='" & Trim(pvendedor) & "' AND F2COBRADOR='*'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If rst.EOF = False Then
        wnomven = rst!F2NOMUSER & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_COBRADOR = sw
End Function

Public Function VALIDA_CANAL(pcanal As String)
Dim sw      As Boolean
    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from EF2CANALES where F2CANCODIGO='" & Trim(pcanal) & "'", cnn_dbbancos, adOpenStatic, adLockReadOnly
    If rst.EOF = False Then
        wnomcanal = rst!F2CANNOMBRE & ""
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CANAL = sw
End Function

Public Function FileExist(ByVal sFileName As String) As Boolean
Dim hFile As Integer

sFileName = Trim$(sFileName)

FileExist = False

If sFileName = "" Then Exit Function

On Error Resume Next
hFile = FreeFile
Open sFileName For Input Access Read Shared As #hFile
If Err.Number = 0 Then FileExist = True
Close #hFile

End Function

Public Function Repetir(ByVal NumeroDeVeces As Integer, ByVal CaracterRepetido As String) As String
Repetir = ""
For i = 1 To NumeroDeVeces
    Repetir = Repetir & CaracterRepetido
Next
End Function



Public Sub Crea_Campo(pCadenaConexion As String, pNombreDeTabla As String, pNombreDeCampo As String, pTipoCampo As String, pEsNull As Boolean, pValorPorDefecto As String)
On Error GoTo CapturaError
Dim StrContenido As String
Dim SwExiste As Boolean
Dim pAf As New ADOFunctions
Dim pRs As New ADODB.Recordset
Dim SqlCad As String, i As Integer
SwExiste = False
If (FileExist(wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl") = False) Then
    SqlCad = "Select * from " & pNombreDeTabla
    Set pRs = pAf.OpenSQLForwardOnly(SqlCad, pCadenaConexion)
    If pRs.State = 1 Then
        For i = 0 To (pRs.Fields.Count - 1)
            If UCase(pRs.Fields(i).Name) = UCase(pNombreDeCampo) Then
                SwExiste = True
                Exit For
            End If
        Next
    End If
Else
    'StrContenido = sGetINI(wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl", wempresa, "Fecha", "l")
    If IsDate(StrContenido) Then
        SwExiste = True
    Else
        SwExiste = False
    End If
End If
If pRs.State = 1 Then pRs.Close
Set pRs = Nothing
If SwExiste = False Then
    csql = "Alter table " & pNombreDeTabla & " ADD COLUMN " & pNombreDeCampo & " " & pTipoCampo & IIf(pEsNull = True, " NULL ", " NOT NULL ") & IIf(Len(Trim(pValorPorDefecto)) > 0, " DEFAULT " & pValorPorDefecto, "")
    Call EJECUTA_SENTENCIA(csql, pCadenaConexion)
    'If Sw_Ejecuta_Sentencia = True Then
    sWrtIni wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl", wempresa, "Fecha", Str(Date)
    'End If
Else
    If (FileExist(wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl") = False) Then
        sWrtIni wrutatemp & UCase(pNombreDeTabla) & "_" & UCase(pNombreDeCampo) & ".tbl", wempresa, "Fecha", Str(Date)
    End If
End If

Exit Sub
CapturaError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Exit Sub
End Sub

Public Sub EJECUTA_SENTENCIA(pSentencia As String, pconexion As String)
On Error GoTo Error_GRABA_REGISTRO_logistica

Dim CnExec As New ADODB.Connection
Sw_Ejecuta_Sentencia = False
CnExec.Open pconexion
CnExec.Execute pSentencia
If CnExec.State = 1 Then CnExec.Close
Set CnExec = Nothing
Sw_Ejecuta_Sentencia = True
Exit Sub

Error_GRABA_REGISTRO_logistica:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, wnomcia
    Select Case Err.Number
    Case 3704
        If CnExec.State = 0 Then CnExec.Open
        Resume
    Case Else
        Sw_Ejecuta_Sentencia = False
        Exit Sub
    End Select
End Sub

Public Function VerificaPermiso(Codigo_de_Permiso As String, Nombre_de_Usuario As String) As Boolean
    Dim Af As New ADOFunctions
    Dim Rs As New ADODB.Recordset
    
    csql = "select * from EF2TAREAUSERS where f2coduser='" & Nombre_de_Usuario & "' and f2codtarea='" & Codigo_de_Permiso & "'"
    
    Set Rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    If Rs.RecordCount > 0 Then
        VerificaPermiso = True
    Else
        VerificaPermiso = False
    End If
    
    If Rs.State = 1 Then Rs.Close
    
    Set Rs = Nothing
End Function

Public Property Get ComputerName() As String
    Dim sName As String
    Dim lRetval As Long
    Dim iPos As Integer
    
    sName = Space$(255)
    lRetval = GetComputerName(sName, 255)
    iPos = InStr(sName, Chr$(0))
    ComputerName = left$(sName, iPos - 1)
End Property

Public Function VerificaAutorizaciones(Codigo_de_Consulta As String, Nombre_de_Usuario As String) As String
    Dim Af As New ADOFunctions
    Dim i As Integer
    Dim Rs As New ADODB.Recordset
    
    'Crea_Campo cconex_dbbancos, "EF2AUTORIZADOS", "F2CODUSER", "String", True, ""
    i = 0
    
    csql = "select F3COSTO from EF2AUTORIZADOS where F2CODUSER='" & Nombre_de_Usuario & "' and F2REPORTE='" & Codigo_de_Consulta & "'"
    
    Set Rs = Af.OpenSQLForwardOnly(csql, cconex_dbbancos)
    
    If Rs.State = 1 Then
        If Rs.RecordCount > 0 Then
            Rs.MoveFirst
            
            Do While Not Rs.EOF
                i = i + 1
                
                If i = 1 Then
                    VerificaAutorizaciones = "'" & Rs!F3COSTO & "'"
                Else
                    VerificaAutorizaciones = VerificaAutorizaciones & ",'" & Rs!F3COSTO & "'"
                End If
                
                Rs.MoveNext
            Loop
        Else
            VerificaAutorizaciones = "''"
        End If
    End If
    
    If Rs.State = 1 Then Rs.Close
    
    Set Rs = Nothing
End Function

'Public Function sGetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
'Dim sTemp As String * 256
'Dim nLength As Integer
'sTemp = Space$(256)
'nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
'sGetINI = left$(sTemp, nLength)
'End Function

Public Sub sWrtIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    'Guarda los datos de configuración
    'Los parámetros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    '
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub
Public Function enviacorreoGmail(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim nError As Long
 With MailObj
    
    .ServerName = Trim(OriginDominio)
    '.ServerPort = CLng(Val(995))
    .UserName = Trim(OriginMail)
    .Password = Trim(OriginPassword)
    .RelayServer = "smtp.gmail.com"
    .RelayPort = 465
    .Secure = True
    .Options = 5

    .timeout = CLng(Val(60))
    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
                                          DestinoMail, _
                                          "betania_bk@outlook.com.pe", _
                                          "jparedes@betania.com.pe", _
                                          DestinoAsunto, _
                                          DestinoCuerpo, _
                                          strMessageHTML, _
                                          2, _
                                          1)
    
    If nError Then
    Exit Function
    End If
    .Priority = 3
'    If nError Then
'        MsgBox ("Unable to create a new message " & .LastErrorString)
'        Exit Function
'    End If
'
'    If .Recipients = 0 Then
'        MsgBox ("There are no recipients for this message")
'        Exit Function
'    End If
    nError = .SendMessage()

    If nError Then
        enviacorreoGmail = False
        MsgBox ("Correo no pudo ser enviado")
    Else
        MsgBox ("Correo enviado con éxito")
        enviacorreoGmail = True
    
    End If
End With
End Function

Public Function enviacorreoPOP(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim strFontName As String
    Dim strFontSize As String
    Dim strMessageHTML As String
    Dim nCharacterSet As Long
    Dim nEncodingType As Long
    Dim nindex As Long
    Dim nError As Long

 With MailObj
    
''    .ServerName = Trim(OriginDominio)
''    .ServerPort = CLng(Val(110))
''    .UserName = Trim(OriginMail)
''    .Password = Trim(OriginPassword)
''    .Timeout = CLng(Val(60))
''
''    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
''                                          DestinoMail, _
''                                          "betania_bk@outlook.com.pe, psalas@betania.com.pe", _
''                                          "responder.britania@gmail.com", _
''                                          DestinoAsunto, _
''                                          DestinoCuerpo, _
''                                          strMessageHTML, _
''                                          2, _
''                                          1)
''
''    If nError Then
''    Exit Function
''    End If
''    .Priority = 3
''    nError = .SendMessage()
''
''    If nError Then
''        enviacorreoPOP = False
''        MsgBox ("Correo no pudo ser enviado")
''    Else
''        MsgBox ("Correo enviado con éxito")
''        enviacorreoPOP = True
''    End If

    
    'nError = CreateMessage()
    nError = .ComposeMessage(OriginMail, _
                             DestinoMail, _
                             "betania_bk@outlook.com.pe", _
                             "responder.britania@gmail.com", _
                             DestinoAsunto, _
                             DestinoCuerpo, _
                             strMessageHTML, _
                             2, _
                             1)
    
    If nError Then
        MsgBox ("Correo no pudo ser creado")
    End If
    
        .RelayServer = ""
        .RelayPort = 0
    nError = .SendMessage()
    If nError Then
        enviacorreoPOP = False
        MsgBox ("Correo no pudo ser enviado")
        Sw_Act = False
    Else
        MsgBox ("Correo enviado con éxito")
        enviacorreoPOP = True
        Sw_Act = True
    End If
End With
End Function

Private Function CreateMessage() As Long
    Dim strFontName As String
    Dim strFontSize As String
    Dim strMessageHTML As String
    Dim nCharacterSet As Long
    Dim nEncodingType As Long
    Dim nindex As Long
    Dim nError As Long
        
    CreateMessage = 0
    
    strMessageHTML = ""
    
    '
    nError = InternetMail1.ComposeMessage(editFrom.Text, _
                                          editTo.Text, _
                                          editCc.Text, _
                                          editBcc.Text, _
                                          editSubject.Text, _
                                          editMessageText.Text, _
                                          strMessageHTML, _
                                          nCharacterSet, _
                                          1)
    
    If nError Then
        CreateMessage = nError
        Exit Function
    End If
    
    
End Function

'Public Function enviacorreoPOPa(ByVal MailObj As Object, _
'ByVal OriginDominio As String, _
'ByVal OriginNombre As String, _
'ByVal OriginMail As String, _
'ByVal OriginPassword As String, _
'ByVal DestinoMail As String, _
'ByVal DestinoAsunto As String, _
'ByVal DestinoCuerpo As String) As Boolean
'
'Dim nError As Long
' With MailObj
'
''    .ServerName = Trim(OriginDominio)
''    '.ServerPort = CLng(Val(995))
''    .UserName = Trim(OriginMail)
''    .Password = Trim(OriginPassword)
''    .RelayServer = "smtp.gmail.com"
''    .RelayPort = 465
''    .Secure = True
''    .Options = 5
'
'    .ServerName = Trim(OriginDominio)
'    .ServerPort = CLng(Val(110))
'    .UserName = Trim(OriginMail)
'    .Password = Trim(OriginPassword)
'    .Timeout = CLng(Val(60))
'
'    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
'                                          DestinoMail, _
'                                          "betania_bk@outlook.com.pe, psalas@betania.com.pe", _
'                                          "responder.britania@gmail.com", _
'                                          DestinoAsunto, _
'                                          DestinoCuerpo, _
'                                          strMessageHTML, _
'                                          2, _
'                                          1)
'
'    If nError Then
'    Exit Function
'    End If
'    .Priority = 3
''    If nError Then
''        MsgBox ("Unable to create a new message " & .LastErrorString)
''        Exit Function
''    End If
''
''    If .Recipients = 0 Then
''        MsgBox ("There are no recipients for this message")
''        Exit Function
''    End If
'    nError = .SendMessage()
'
'    If nError Then
'        enviacorreoPOP = False
'        MsgBox ("Correo no pudo ser enviado")
'    Else
'        MsgBox ("Correo enviado con éxito")
'        enviacorreoPOP = True
'    End If
'End With
'End Function
'
'
'
Public Function enviacorreohotmail(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim nError As Long
 With MailObj
    
    .ServerName = Trim(OriginDominio)
'    .ServerPort = CLng(Val(995))
    .UserName = Trim("OriginMail")
    .Password = Trim("OriginPassword")
    .RelayServer = "smtp.live.com"
    .RelayPort = 587
    '.Secure = True
    .Options = 8


    .timeout = CLng(Val(60))
    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
                                          DestinoMail, _
                                          "responder.britania@gmail.com", _
                                          "jparedes@betania.com.pe", _
                                          DestinoAsunto, _
                                          DestinoCuerpo, _
                                          strMessageHTML, _
                                          2, _
                                          1)
    
    If nError Then
    Exit Function
    End If
    .Priority = 3
'    If nError Then
'        MsgBox ("Unable to create a new message " & .LastErrorString)
'        Exit Function
'    End If
'
'    If .Recipients = 0 Then
'        MsgBox ("There are no recipients for this message")
'        Exit Function
'    End If
    nError = .SendMessage()

    If nError Then
        enviacorreohotmail = False
        MsgBox ("Correo no pudo ser enviado")
    Else
        MsgBox ("Correo enviado con éxito")
        enviacorreohotmail = True
    End If
End With
End Function


Public Function enviacorreo(ByVal MailObj As Object, _
ByVal OriginDominio As String, _
ByVal OriginNombre As String, _
ByVal OriginMail As String, _
ByVal OriginPassword As String, _
ByVal DestinoMail As String, _
ByVal DestinoAsunto As String, _
ByVal DestinoCuerpo As String) As Boolean

Dim nError As Long
 With MailObj
    .ServerName = Trim(OriginDominio)
    .ServerPort = CLng(Val(110))
    .UserName = Trim(OriginMail)
    .Password = Trim(OriginPassword)
    .timeout = CLng(Val(60))
    nError = .ComposeMessage(OriginNombre & " <" & OriginMail & ">", _
                                          DestinoMail, _
                                          "", _
                                          "", _
                                          DestinoAsunto, _
                                          DestinoCuerpo, _
                                          "", _
                                          2, _
                                          1)
    If nError Then
    Exit Function
    End If
    .Priority = 3
    If nError Then
        MsgBox ("Unable to create a new message " & .LastErrorString)
        Exit Function
    End If
    
    If .Recipients = 0 Then
        MsgBox ("There are no recipients for this message")
        Exit Function
    End If
    nError = .SendMessage()
    
    If nError Then
        enviacorreo = False
    Else
        enviacorreo = True
    End If
End With
End Function

''*************************************
''*************************************

Public Function FileExists(strPath As String) As Boolean
    On Error Resume Next

    If Len(strPath) < 4 Then
        FileExists = False

        Exit Function
    End If

    FileExists = IIf(Dir(strPath, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "", True, False)
End Function

Public Function sGetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer

    sTemp = Space$(256)

    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)

    sGetINI = left$(sTemp, nLength)
End Function



Public Function devuelveCodigoMes(ByVal strMes As String) As String
    Select Case UCase(strMes)
        Case "TOTAL", "TOT"
            devuelveCodigoMes = "00"
        Case "ENERO", "ENE", "JAN"
            devuelveCodigoMes = "01"
        Case "FEBRERO", "FEB"
            devuelveCodigoMes = "02"
        Case "MARZO", "MAR"
            devuelveCodigoMes = "03"
        Case "ABRIL", "ABR", "APR"
            devuelveCodigoMes = "04"
        Case "MAYO", "MAY"
            devuelveCodigoMes = "05"
        Case "JUNIO", "JUN"
            devuelveCodigoMes = "06"
        Case "JULIO", "JUL"
            devuelveCodigoMes = "07"
        Case "AGOSTO", "AGO", "AUG"
            devuelveCodigoMes = "08"
        Case "SETIEMBRE", "SET", "SEP"
            devuelveCodigoMes = "09"
        Case "OCTUBRE", "OCT"
            devuelveCodigoMes = "10"
        Case "NOVIEMBRE", "NOV"
            devuelveCodigoMes = "11"
        Case "DICIEMBRE", "DIC", "DEC"
            devuelveCodigoMes = "12"
        Case "1º TRIMESTRE"
            devuelveCodigoMes = "Q1"
        Case "2º TRIMESTRE"
            devuelveCodigoMes = "Q3"
        Case "3º TRIMESTRE"
            devuelveCodigoMes = "Q3"
        Case "4º TRIMESTRE"
            devuelveCodigoMes = "Q4"
    End Select
End Function

Public Function devuelveNombreMes(ByVal strCodMes As String) As String
    Select Case strCodMes
        Case "00", "0"
            devuelveNombreMes = "TOTAL"
        Case "01", "1"
            devuelveNombreMes = "ENERO"
        Case "02", "2"
            devuelveNombreMes = "FEBRERO"
        Case "03", "3"
            devuelveNombreMes = "MARZO"
        Case "04", "4"
            devuelveNombreMes = "ABRIL"
        Case "05", "5"
            devuelveNombreMes = "MAYO"
        Case "06", "6"
            devuelveNombreMes = "JUNIO"
        Case "07", "7"
            devuelveNombreMes = "JULIO"
        Case "08", "8"
            devuelveNombreMes = "AGOSTO"
        Case "09", "9"
            devuelveNombreMes = "SETIEMBRE"
        Case "10", "10"
            devuelveNombreMes = "OCTUBRE"
        Case "11", "11"
            devuelveNombreMes = "NOVIEMBRE"
        Case "12", "12"
            devuelveNombreMes = "DICIEMBRE"
        Case "Q1"
            devuelveNombreMes = "1º TRIMESTRE"
        Case "Q2"
            devuelveNombreMes = "2º TRIMESTRE"
        Case "Q3"
            devuelveNombreMes = "3º TRIMESTRE"
        Case "Q4"
            devuelveNombreMes = "4º TRIMESTRE"
    End Select
End Function

Public Function mesInicioTrimestre(ByVal strCodTri As String) As Integer
    Select Case strCodTri
        Case "Q1"
            mesInicioTrimestre = 1
        Case "Q2"
            mesInicioTrimestre = 4
        Case "Q3"
            mesInicioTrimestre = 7
        Case "Q4"
            mesInicioTrimestre = 10
    End Select
End Function

Public Function mesFinTrimestre(ByVal strCodTri As String) As Integer
    Select Case strCodTri
        Case "Q1"
            mesFinTrimestre = 3
        Case "Q2"
            mesFinTrimestre = 6
        Case "Q3"
            mesFinTrimestre = 9
        Case "Q4"
            mesFinTrimestre = 12
    End Select
End Function

Public Function devuelveCantRegistros(ByVal rsRegistro As ADODB.Recordset) As Double
    Dim i As Double
    
    i = 0
    
    If Not rsRegistro.EOF Then
        rsRegistro.MoveFirst
            
        Do While Not rsRegistro.EOF
            i = i + 1
            
            rsRegistro.MoveNext
        Loop
            rsRegistro.MoveFirst
    End If
    
    devuelveCantRegistros = i
End Function

Public Function obtenerVecesCaracterEnCadena(ByVal strCadena As String, _
                                            ByVal strCaracterBuscado As String) As Integer
    Dim indice, cantVueltas As Integer
    Dim caracter As String
    
    cantVueltas = Len(Trim(strCadena))
    
    For indice = 1 To cantVueltas
        caracter = Mid(Trim(strCadena), indice, 1)
        
        If caracter = strCaracterBuscado Then
             obtenerVecesCaracterEnCadena = obtenerVecesCaracterEnCadena + 1
        End If
    Next indice
End Function

Public Function limpiarCaracteresEnCadena(ByVal strCadena As String, _
                                            Optional ByVal bolObviarValidarEspacios As Boolean, _
                                            Optional ByVal bolConvertirToMayus As Boolean) As String
    Dim indice, cantVueltas As Integer
    Dim caracter As String
    Dim cadenaFinal As String
    
    cantVueltas = Len(Trim(strCadena))
    
    cadenaFinal = vbNullString
    
    For indice = 1 To cantVueltas
        caracter = Mid(Trim(strCadena), indice, 1)
        
        Select Case caracter
            Case "á"
                caracter = "a"
            Case "Á"
                caracter = "A"
            Case "é"
                caracter = "e"
            Case "É"
                caracter = "E"
            Case "í"
                caracter = "i"
            Case "Í"
                caracter = "I"
            Case "ó"
                caracter = "o"
            Case "Ó"
                caracter = "O"
            Case "ú"
                caracter = "u"
            Case "Ú"
                caracter = "U"
            Case ",", ".", "(", ")", "'"
                caracter = ""
            Case "|", "%", "•", "$", "¿", "?", "^", "Ç", "ª", "º", "~", "€", "¬", "ç", "¡"
                caracter = " "
        End Select
        
        cadenaFinal = cadenaFinal & caracter
    Next indice
        If Not bolObviarValidarEspacios Then
            If InStr(1, cadenaFinal, "   ") > 0 Then
                cadenaFinal = Replace(cadenaFinal, "   ", " ")
            End If
            
            If InStr(1, cadenaFinal, "  ") > 0 Then
                cadenaFinal = Replace(cadenaFinal, "  ", " ")
            End If
        End If
            
            
    limpiarCaracteresEnCadena = IIf(bolConvertirToMayus, UCase(cadenaFinal), cadenaFinal)
End Function

Public Function repetirCaracter(ByVal caracter As String, ByVal veces As Long) As String
    Dim i As Integer
    
    repetirCaracter = vbNullString
    
    For i = 1 To veces
        repetirCaracter = repetirCaracter & caracter
    Next i
End Function

Public Function obtenerValorPorcentaje(ByVal valPorcentaje As String) As Double
    If InStr(1, valPorcentaje, "%") > 0 Then
        obtenerValorPorcentaje = Val(Mid(valPorcentaje, 1, InStr(1, valPorcentaje, "%") - 1))
    Else
        obtenerValorPorcentaje = Val(valPorcentaje)
    End If
End Function

Public Function obtenerValorEnMiles(ByVal valEnMiles As String) As Double
    If InStr(1, valEnMiles, ".") = 0 Then
        obtenerValorEnMiles = Val(valEnMiles)
    Else
        Dim indice, cantVueltas, cantChar, cantCeros As Integer
        Dim caracter As String
        Dim numMultiplicador As Long
        
        cantVueltas = Len(Trim(valEnMiles))
        
        For indice = 1 To cantVueltas
            caracter = Mid(Trim(valEnMiles), indice, 1)
            
            If caracter = "." Then
                cantChar = cantChar + 1
            End If
        Next indice
        
        cantCeros = cantChar * 3
        numMultiplicador = Val(1 & repetirCaracter("0", cantCeros))
        
        If cantChar > 0 And Int(Val(valEnMiles)) > 0 Then
            obtenerValorEnMiles = Val(valEnMiles) * numMultiplicador
        Else
            obtenerValorEnMiles = Val(valEnMiles)
        End If
    End If
End Function

Public Function seleccionarItem(ByVal tipoCombo As Object, _
                                ByVal DatoBuscado As String, _
                                Optional ByVal posicionDatoDERorIZQorNULL As String, _
                                Optional ByVal largoDato As Integer) As Integer
    Dim i, indice As Integer
    
    indice = -1
    
    For i = 0 To tipoCombo.ListCount - 1
        If Trim(posicionDatoDERorIZQorNULL) = vbNullString Then
            If UCase(Trim(tipoCombo.List(i))) = UCase(DatoBuscado) Then
                indice = i
                
                Exit For
            End If
        ElseIf UCase(Trim(posicionDatoDERorIZQorNULL)) = "DER" Then
            If right(UCase(Trim(tipoCombo.List(i))), largoDato) = UCase(DatoBuscado) Then
                indice = i
                
                Exit For
            End If
        ElseIf UCase(Trim(posicionDatoDERorIZQorNULL)) = "IZQ" Then
            If left(UCase(Trim(tipoCombo.List(i))), largoDato) = UCase(DatoBuscado) Then
                indice = i
                
                Exit For
            End If
        End If
    Next i
    
    seleccionarItem = indice
End Function

Public Function evaluarDias(ByVal lngValorAbsolutoDeDias As Long) As String
    Select Case lngValorAbsolutoDeDias
        Case Is < 365
            evaluarDias = lngValorAbsolutoDeDias & " día" & IIf(lngValorAbsolutoDeDias > 1, "s", vbNullString)
        Case Is >= 365
            Dim valorEntAnnos, valorDifDias As Long
            
            valorEntAnnos = Int(lngValorAbsolutoDeDias / 365)
            valorDifDias = lngValorAbsolutoDeDias - (valorEntAnnos * 365)
            
            evaluarDias = valorEntAnnos & " año" & IIf(valorEntAnnos > 1, "s", vbNullString) & " y " & _
                            valorDifDias & " día" & IIf(valorDifDias > 1, "s", vbNullString)
                            
            
            valorEntAnnos = 0
            valorDifDias = 0
    End Select
End Function

Public Sub deshabilitarBotonCerrarForm(ByVal El_Formulario As Form)
    Dim Hwnd_Menu As Long

    ' Obtiene el Hwnd del menú para usar con el Api DeleteMenu
    Hwnd_Menu = GetSystemMenu(El_Formulario.HWnd, False)

    ' botón Cerrar
    Call DeleteMenu(Hwnd_Menu, 6, MF_BYPOSITION)
End Sub


'----------------------------------------------------------------------------------------------------
'::::::::::::::::::::::::::::::::::: SK PROCEDIMIENTOS ADICIONALES ::::::::::::::::::::::::::::::::::
'----------------------------------------------------------------------------------------------------

Public Function validarSoloTeclaEnter(ByVal Key As Integer) As Integer
    On Error GoTo errValidarSoloTeclaEnter
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case Else
            validarSoloTeclaEnter = Key
    End Select
    
    Exit Function
errValidarSoloTeclaEnter:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
            Err.Clear
    End Select
End Function

Public Function validarCajaTexto(ByVal Key As Integer) As Integer
    On Error GoTo errValidarCajaTexto
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case 39, 44, 59
            MsgBox ("No acepta caracteres especiales como '|' (') ';'  ',' etc... Ingrese solo Caracteres.")
            validarCajaTexto = 0
        Case 3, 8, 22, 24, 26, 32, 37, 46, 58, 65 To 90, 209 'Espacio & Mayusculas
            validarCajaTexto = Key
        Case 97 To 122 'Minusculas convertidas a Mayusculas
            validarCajaTexto = Key - 32
        Case 193, 201, 205, 211, 218, 209, 225, 233, 237, 243, 250 'Tildados Mayusculas
            MsgBox ("No acepta caracteres tildados como 'á' 'Á'  'ó'  'Ó' etc.")
            validarCajaTexto = 0
        Case 241 'La 'ñ'
            validarCajaTexto = Key - 32
        Case 45, 47 '- ^ / ^
            validarCajaTexto = Key
        Case Else
            MsgBox ("Caracter inválido '" & Chr(Key) & "'.")
            validarCajaTexto = 0
    End Select
    
    Exit Function
errValidarCajaTexto:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName
            Err.Clear
    End Select
End Function

Public Function validarCajaTextoSinMayus(ByVal Key As Integer) As Integer
    On Error GoTo errValidarCajaTextoSinMayus
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case 39, 44
            MsgBox ("No acepta caracteres especiales como '|' (')  ',' etc... Ingrese solo Caracteres.")
            validarCajaTextoSinMayus = 0
        Case 3, 8, 22, 24, 26, 32, 37, 46, 58, 65 To 90, 209 'Espacio & Mayusculas
            validarCajaTextoSinMayus = Key
        Case 193, 201, 205, 211, 218, 209, 225, 233, 237, 243, 250 'Tildados Mayusculas
            MsgBox ("No acepta caracteres tildados como 'á' 'Á'  'ó'  'Ó' etc.")
            validarCajaTextoSinMayus = 0
        Case 45, 47 '- ^ / ^
            validarCajaTextoSinMayus = Key
        Case Else
            'MsgBox ("Caracter inválido '" & Chr(Key) & "'.")
            validarCajaTextoSinMayus = Key
    End Select
    
    Exit Function
errValidarCajaTextoSinMayus:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
            Err.Clear
    End Select
End Function

Public Function validarCajaNumerica(ByVal Key As Integer) As Integer
    On Error GoTo errValidarCajaNumerica
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case 39, 44, 58, 59
            MsgBox ("No acepta caracteres especiales como '|' (') ';'  ':'  ',' etc... Ingrese solo Digitos.")
            validarCajaNumerica = 0
        Case 3, 8, 22, 24, 26, 32, 46, 48 To 57
            validarCajaNumerica = Key
        Case Else
            MsgBox ("Caracter inválido '" & Chr(Key) & "'.")
            validarCajaNumerica = 0
    End Select
    
    Exit Function
errValidarCajaNumerica:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
            Err.Clear
    End Select
End Function

Public Function validarCajaAlfaNumerica(ByVal Key As Integer) As Integer
    On Error GoTo errValidarCajaAlfaNumerica
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case 39, 44, 59
            MsgBox ("No acepta caracteres especiales como '|' (') ';'  ',' etc.")
            validarCajaAlfaNumerica = 0
        Case 3, 8, 22, 24, 26, 32, 37, 45, 47, 58, 65 To 90, 209 'Espacio & Mayusculas
            validarCajaAlfaNumerica = Key
        Case 40, 41, 42, 46, 48 To 57 'Numeros
            validarCajaAlfaNumerica = Key
        Case 97 To 122 'Minusculas convertidas a Mayusculas
            validarCajaAlfaNumerica = Key - 32
        Case 193, 201, 205, 211, 218, 209, 225, 233, 237, 243, 250 'Tildados Mayusculas
            MsgBox ("No acepta caracteres tildados como 'á' 'Á'  'ó'  'Ó' etc.")
            validarCajaAlfaNumerica = 0
        Case 241 'La 'ñ'
            validarCajaAlfaNumerica = Key - 32
        Case Else
            MsgBox ("Caracter inválido '" & Chr(Key) & "'.")
            validarCajaAlfaNumerica = 0
    End Select
    
    Exit Function
errValidarCajaAlfaNumerica:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
            Err.Clear
    End Select
End Function

Public Function validarCajaCodigosConsecutivos(ByVal Key As Integer) As Integer
    On Error GoTo errValidarCajaCodigosConsecutivos
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case 39, 58, 59
            MsgBox ("No acepta caracteres especiales como '|' (') ';'  ':' etc... Ingrese solo Digitos.")
            validarCajaCodigosConsecutivos = 0
        Case 3, 8, 22, 24, 32, 42, 46, 48 To 57, 44
            validarCajaCodigosConsecutivos = Key
        Case Else
            MsgBox ("Caracter inválido '" & Chr(Key) & "'.")
            validarCajaCodigosConsecutivos = 0
    End Select
    
    Exit Function
errValidarCajaCodigosConsecutivos:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
            Err.Clear
    End Select
End Function

Public Function validarCajaTextoNomCarpeta(ByVal Key As Integer) As Integer
    On Error GoTo errValidarCajaTextoNomCarpeta
    
    Select Case Key
        Case 13
            ModUtilitario.pulsarTecla vbKeyTab
        Case 39, 44, 59
            MsgBox ("No acepta caracteres especiales como '|' (') ';'  ',' etc... Ingrese solo Caracteres.")
            validarCajaTextoNomCarpeta = 0
        Case 8, 24, 26, 32, 37, 46, 65 To 90, 209 'Espacio & Mayusculas
            validarCajaTextoNomCarpeta = Key
        Case 3, 22
            MsgBox ("Imposible Copiar o Pegar.")
            validarCajaTextoNomCarpeta = 0
        Case 97 To 122 'Minusculas convertidas a Mayusculas
            validarCajaTextoNomCarpeta = Key - 32
        Case 193, 201, 205, 211, 218, 209, 225, 233, 237, 243, 250 'Tildados Mayusculas
            MsgBox ("No acepta caracteres tildados como 'á' 'Á'  'ó'  'Ó' etc.")
            validarCajaTextoNomCarpeta = 0
        Case Else
            MsgBox ("Caracter inválido '" & Chr(Key) & "'.")
            validarCajaTextoNomCarpeta = 0
    End Select
    
    Exit Function
errValidarCajaTextoNomCarpeta:
    Select Case Err.Number
        Case 70
            Resume Next
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
            Err.Clear
    End Select
End Function

Public Sub seleccionarTextoCaja(ByVal cajaTexto As TextBox)
    cajaTexto.SelStart = 0: cajaTexto.SelLength = Len(cajaTexto.Text)
End Sub

Public Function ObtenerCampoV2(ByVal conConexionDeBaseDeDatos As ADODB.Connection, _
                                ByVal strCampo As String, _
                                ByVal strTabla As String, _
                                ByVal strCampoCom As String, _
                                ByVal strValor As String, _
                                ByVal strTipoDeComparacion As String, _
                                Optional ByVal strCondicion As String) As String
    On Error GoTo errObtenerCampoV2
    
    Dim strCadenaSQL As String
    Dim rst As New ADODB.Recordset
    
    If strTipoDeComparacion = "F" Then
        strCampoCom = "CVDATE(" & strCampoCom & ") = CVDATE('" & strValor & "') "
    ElseIf strTipoDeComparacion = "T" Then
        strCampoCom = strCampoCom & " = '" & strValor & "' "
    ElseIf strTipoDeComparacion = "N" Then
        strCampoCom = strCampoCom & " = " & strValor & " "
    ElseIf strTipoDeComparacion = vbNullString Then
        strCampoCom = vbNullString
    End If
    
    strCadenaSQL = "SELECT " & strCampo & " FROM " & strTabla & " WHERE " & strCampoCom & strCondicion
    
    If rst.State = 1 Then rst.Close
    
    rst.Open strCadenaSQL, conConexionDeBaseDeDatos, adOpenForwardOnly, adLockReadOnly
    
    ObtenerCampoV2 = vbNullString
    
    If Not rst.EOF And Not IsNull(rst.Fields(0)) Then ObtenerCampoV2 = rst.Fields(0)
    
    Exit Function
    Resume
errObtenerCampoV2:
    Select Case Err.Number
        Case 3709
            conConexionDeBaseDeDatos.Open conConexionDeBaseDeDatos.ConnectionString
            
            Resume
        Case Else
            MsgBox "No. Error: " & Err.Number & vbNewLine & _
                    "Descripción: " & Err.Description, _
                    vbCritical, App.ProductName & " - ModUtilitario: ObtenerCampoV2"
    End Select
    
    ObtenerCampoV2 = vbNullString
    
    Err.Clear
End Function

Public Sub listarMeses(ByVal tipoEvaluacion As Integer, _
                        ByVal comboList As Object, _
                        ByVal incluyeTotal As Boolean)
    Select Case tipoEvaluacion
        Case 0
            With comboList
                .Clear
                
                If incluyeTotal Then
                    .AddItem "(*) - Todos" & Space(150) & "00"
                End If
                
                .AddItem "Enero" & Space(150) & "01"
                .AddItem "Febrero" & Space(150) & "02"
                .AddItem "Marzo" & Space(150) & "03"
                .AddItem "Abril" & Space(150) & "04"
                .AddItem "Mayo" & Space(150) & "05"
                .AddItem "Junio" & Space(150) & "06"
                .AddItem "Julio" & Space(150) & "07"
                .AddItem "Agosto" & Space(150) & "08"
                .AddItem "Setiembre" & Space(150) & "09"
                .AddItem "Octubre" & Space(150) & "10"
                .AddItem "Noviembre" & Space(150) & "11"
                .AddItem "Diciembre" & Space(150) & "12"
                
                .ListIndex = 0
            End With
        Case 1
            With comboList
                .Clear
                
                .AddItem "1º Trimestre" & Space(150) & "Q1"
                .AddItem "2º Trimestre" & Space(150) & "Q2"
                .AddItem "3º Trimestre" & Space(150) & "Q3"
                .AddItem "4º Trimestre" & Space(150) & "Q4"
                
                .ListIndex = 0
            End With
    End Select
End Sub

Public Function validarUsoRegistro(ByVal conConexionDeBaseDeDatos As ADODB.Connection, _
                                ByVal strCampoLlave As String, _
                                ByVal strTablaDependiente As String, _
                                ByVal strValorLlave As String, _
                                ByVal strTipoDeComparacion As String, _
                                Optional ByVal strCondicionAdicional As String) As String
    
    On Error GoTo errValidarUsoRegistro
    
    Dim strCadenaSQL, strCondicion As String
    Dim rst As New ADODB.Recordset
    
    If strTipoDeComparacion = "F" Then
        strCondicion = "CVDATE(" & strCampoLlave & ") = CVDATE('" & strValorLlave & "') "
    ElseIf strTipoDeComparacion = "T" Then
        strCondicion = strCampoLlave & " = '" & strValorLlave & "' "
    ElseIf strTipoDeComparacion = "N" Then
        strCondicion = strCampoLlave & " = " & strValorLlave & " "
    End If
    
    strCadenaSQL = "SELECT COUNT(" & strCampoLlave & ") FROM " & strTablaDependiente & " WHERE " & strCondicion & strCondicionAdicional
    
    If rst.State = 1 Then rst.Close
    
    rst.Open strCadenaSQL, conConexionDeBaseDeDatos, adOpenForwardOnly, adLockReadOnly
    
    validarUsoRegistro = vbNullString
    
    If Not rst.EOF And Not IsNull(rst.Fields(0)) Then validarUsoRegistro = Trim(rst.Fields(0) & "")
    
    Exit Function
    Resume
errValidarUsoRegistro:
    validarUsoRegistro = vbNullString
    
    MsgBox "No. Error: " & Err.Number & vbNewLine & _
            "Descripción: " & Err.Description, _
            vbCritical, App.ProductName & " - ModUtilitario: ValidarUsoRegistro"
    
    Err.Clear
End Function

Public Sub validarCodigosConsecutivosTexto(ByVal cajaTextoCodigos As TextBox, _
                                            ByVal cajaTextoDescripciones As TextBox, _
                                            ByVal strCampo As String, _
                                            ByVal strTabla As String, _
                                            ByVal strCampoCom As String, _
                                            Optional ByVal strCondicion As String, _
                                            Optional ByVal intMaxLargoParaDescrip As Integer)
    Dim indCaracter, indice, cantVueltas As Integer
    Dim strCodConsecutivos, strCodigoActual, strDesBuscada, strCadenaFinal As String

    strCodConsecutivos = Trim(cajaTextoCodigos.Text)
    cantVueltas = obtenerVecesCaracterEnCadena(strCodConsecutivos, ",")

    If cantVueltas > 0 Then
        cantVueltas = cantVueltas + 1

        strDesBuscada = vbNullString
        cajaTextoCodigos.Text = vbNullString
        cajaTextoDescripciones.Text = vbNullString

        For indice = 1 To cantVueltas
            indCaracter = InStr(1, Trim(strCodConsecutivos), ",")

            strCodigoActual = Trim(Mid(Trim(strCodConsecutivos), 1, IIf(indCaracter > 1, indCaracter - 1, Len(strCodConsecutivos))))
            
            strDesBuscada = ObtenerCampoV2(cnn_dbbancos, _
                                        strCampo, strTabla, _
                                        strCampoCom, strCodigoActual, _
                                        "T", strCondicion)
            
            If intMaxLargoParaDescrip > 0 Then
                If Len(strDesBuscada) > intMaxLargoParaDescrip Then
                    strDesBuscada = left(strDesBuscada, intMaxLargoParaDescrip)
                End If
            End If

            If strDesBuscada <> vbNullString Then
                cajaTextoCodigos.Text = cajaTextoCodigos.Text & _
                                                IIf(cajaTextoCodigos.Text <> vbNullString, ",", vbNullString) & _
                                                strCodigoActual

                cajaTextoDescripciones.Text = cajaTextoDescripciones.Text & _
                                                IIf(cajaTextoDescripciones.Text <> vbNullString, "," & vbNewLine, vbNullString) & _
                                                strDesBuscada
            End If

            If indice < cantVueltas Then
                strCodConsecutivos = Trim(Mid(Trim(strCodConsecutivos), indCaracter + 1))
            End If
        Next indice
    Else
        strCodigoActual = Trim(strCodConsecutivos)

        strDesBuscada = ObtenerCampoV2(cnn_dbbancos, _
                                    strCampo, strTabla, _
                                    strCampoCom, strCodigoActual, _
                                    "T", strCondicion)

        If intMaxLargoParaDescrip > 0 Then
            strDesBuscada = left(strDesBuscada, intMaxLargoParaDescrip)
        End If

        If strDesBuscada <> vbNullString Then
            cajaTextoCodigos.Text = strCodigoActual
            cajaTextoDescripciones.Text = strDesBuscada
        End If
    End If
End Sub

Public Function validarFormAbierto(ByVal strNomForm As String) As Boolean
    If strNomForm = vbNullString Then
        validarFormAbierto = False
        
        Exit Function
    End If
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = strNomForm Then
            validarFormAbierto = True
            
            Exit For
        End If
    Next frm
    
    Set frm = Nothing
End Function


Public Function generarRutaDestino(ByVal strRutaRaiz As String, _
                                ByVal strRutaAnexaAcrear As String) As String
                                
    On Error GoTo errGenerarRutaDestino
    
    Screen.MousePointer = vbHourglass
    
    If Dir(strRutaRaiz, vbDirectory) = vbNullString Then
        MsgBox "Verifique que la Ruta de Raíz configurada actualmente exista." & _
                "Generación de carpeta de DESTINO fallida.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
        
        generarRutaDestino = vbNullString
        
        Screen.MousePointer = vbDefault
        
        Exit Function
    End If
    
    Dim arraySubDir As Variant
    Dim strRutaFinal As String
    Dim strRutaTmp As String
    Dim strRutaAnt As String
    Dim i As Integer
    
    arraySubDir = Split(strRutaAnexaAcrear, "\")
    
    strRutaFinal = vbNullString
    strRutaFinal = strRutaRaiz
                                                            
    For i = LBound(arraySubDir) To UBound(arraySubDir)
        
        If Len(arraySubDir(i)) > 0 Then
            If InStr(1, Trim(arraySubDir(i)), "-") = 0 Then
                strRutaFinal = strRutaFinal & "\" & Trim(arraySubDir(i))
                
                If Dir(strRutaFinal, vbDirectory) = vbNullString Then
                    MkDir strRutaFinal
                End If
            Else
                strRutaAnt = strRutaFinal
                
                strRutaTmp = strRutaAnt & "\" & Trim(arraySubDir(i)) 'Mid(Trim(arraySubDir(i)), 1, InStr(1, Trim(arraySubDir(i)), "-") - 1) & "*"
                
                strRutaFinal = strRutaFinal & "\" & Trim(arraySubDir(i))
                
                If Dir(strRutaTmp, vbDirectory) = vbNullString Then
                    MkDir strRutaFinal
                Else
                    If strRutaAnt & "\" & Dir(strRutaTmp, vbDirectory) <> strRutaFinal Then
                        Name strRutaAnt & "\" & Dir(strRutaTmp, vbDirectory) As strRutaFinal
                    End If
                End If
            End If
        Else
            strRutaFinal = vbNullString
            
            Exit For
        End If
    Next i
    
    generarRutaDestino = strRutaFinal
    
    Screen.MousePointer = vbDefault
    
    Exit Function
errGenerarRutaDestino:
    If Err.Number = 75 Then
        MsgBox "No se logro renombrar carpeta, verifique que ningun archivo de la misma este abierto.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
        
        Resume Next
    Else
        MsgBox "No Error: " & Err.Number & vbNewLine & _
                "Descripción: " & Err.Description & vbNewLine & _
                "Generación de ruta de destino fallida.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
        
        generarRutaDestino = vbNullString
    End If
    
    Screen.MousePointer = vbDefault
    
    Err.Clear
End Function

Public Function buscarCarpeta(Optional Titulo As String, _
                                Optional Path_Inicial As Variant) As String
    On Local Error GoTo errBuscarCarpeta
    
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
    buscarCarpeta = o_Carpeta.Path
  
    Exit Function
errBuscarCarpeta:
    buscarCarpeta = vbNullString
    
    MsgBox "No.: " & Err.Number & vbNewLine & _
            "Decripción: " & Err.Description, vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
    
    Err.Clear
End Function

Public Function actualizarTCambioWebSunat(ByVal strFechaConsulta As String, _
                                            Optional ByVal bolMostrarMensaje As Boolean) As Boolean
    
    On Error GoTo errActualizarTCambioWebSunat

    actualizarTCambioWebSunat = False

    Dim strDia, strTexto, strCadenaBuscada As String
    Dim dblIntentos, dblPosicionDia As Double
    Dim XML

    'Seleccionamos el Valor Día de la Fecha de Consulta
    strDia = Trim(Day(CDate(strFechaConsulta)) & "")

    strTexto = vbNullString
    dblIntentos = 0

    'Intentar Capturar el Texto obtenido de Consulta Anterior
    'Do While InStr(1, strTexto, "Tipo de cambio publicado al") > 0
    Do While dblIntentos < 21
        'Instanciar un Objeto XML de Microsoft
        Set XML = CreateObject("Microsoft.XMLHTTP")
        'Setear los parametros de Apertura del Objeto, para Consultar la URL
        XML.Open "POST", "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias", False

        'Ejecutamos la Pagina en el Objeto
        XML.send
        'Capturamos el Texto de Respuesta albergado en el Objeto
        strTexto = XML.responseText

        Set XML = Nothing

        dblIntentos = dblIntentos + 1

        If InStr(1, strTexto, "Tipo de cambio publicado al") > 0 Then
            Exit Do
        Else
            Clipboard.SetText strTexto
            strTexto = vbNullString
        End If
    Loop

    If strTexto <> vbNullString Then
        strCadenaBuscada = vbNullString
        'Realizar Mientras Cadena a Buscar (Dia + "Clausula HTML") sea igual a Vacio
        Do While strCadenaBuscada = vbNullString
            'Inicializamos la Cadena a Buscar en el Texto, capturado
            'de la Consulta de la URL de la Sunat
            strCadenaBuscada = strDia & "</strong>"
            'Si no encontramos el Dia Consultado en el Texto de la Pagina Sunat, limpiamos la Cadena a Buscar
            If InStr(1, strTexto, strCadenaBuscada) = 0 Then
                strCadenaBuscada = vbNullString
            Else
                Exit Do
            End If
            'Disminuir el día, de uno en uno hasta encontrar el Ultimo día Registrado en Sunat
            strDia = Trim((Val(strDia) - 1) & "")
            'Salir si Día llega a ser '0' (Cero)
            If Val(strDia) = 0 Then Exit Do
        Loop

        If strCadenaBuscada <> vbNullString Then
            dblPosicionDia = InStr(strTexto, strCadenaBuscada)

            With objAyudaTCambio
                .inicializarEntidades

                .Fecha = strFechaConsulta 'Format(strDia, "00") & "/" & Format(strFechaConsulta, "mm/yyyy")

                'If Not .verificarExistencia Then
                    .Cambio = Val(Mid(strTexto, dblPosicionDia + 192, 10))
                    .TCCompra = Val(Mid(strTexto, dblPosicionDia + 101, 10))
                    .TCVenta = Val(Mid(strTexto, dblPosicionDia + 192, 10))

                    strFechaConsulta = Format(strDia, "00") & "/" & Format(strFechaConsulta, "mm/yyyy")

                    If .guardarTCambio Then
                        If bolMostrarMensaje Then
                            If .Fecha = strFechaConsulta Then
                                .SQLSelectAlter = vbNullString
                                .SQLSelectAlter = "UPDATE CAMBIOS SET CODMONEDA = '*' WHERE CVDATE(FECHA) = CVDATE('" & strFechaConsulta & "')"

                                cnn_dbbancos.Execute .SQLSelectAlter

                                Actualiza_Log .SQLSelectAlter, StrConexDbBancos

                                .SQLSelectAlter = vbNullString

                                MsgBox "Tipo de cambio en SUNAT para hoy " & .Fecha & ":" & vbNewLine & _
                                        "       - COMPRA    : " & .TCCompra & vbNewLine & _
                                        "       - VENTA     : " & .TCVenta & vbNewLine & vbNewLine & _
                                        "Actualizado en el sistema.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
                            Else
                                MsgBox "Ultimo Tipo de cambio registrado en SUNAT al " & strFechaConsulta & ":" & vbNewLine & _
                                        "       - COMPRA    : " & .TCCompra & vbNewLine & _
                                        "       - VENTA     : " & .TCVenta & vbNewLine & vbNewLine & _
                                        "Actualizado en el sistema para hoy.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
                            End If
                        End If
                    End If
                'End If
            End With

            actualizarTCambioWebSunat = True
        Else
'            With objAyudaTCambio
'                .inicializarEntidades
'
'                .Fecha = .obtenerFechaUltimoTCambio
'
'                .obtenerTCambio
'
'                .Fecha = strFechaConsulta
'
'                If Not .verificarExistencia Then
'                    If .guardarTCambio Then
'                        MsgBox "No se encontraron Tipos de Cambio registrados en SUNAT" & vbNewLine & _
'                                "para hoy " & .Fecha & ", por lo cual se copiaron del" & _
'                                "registro con Fecha " & .obtenerFechaUltimoTCambio & " :" & vbNewLine & _
'                                "       - COMPRA    : " & .TCCompra & vbNewLine & _
'                                "       - VENTA     : " & .TCVenta & vbNewLine & vbNewLine & _
'                                "Actualizado en el sistema para hoy.", vbInformation + vbOKOnly, App.ProductName
'                    End If
'                End If
'            End With
            If bolMostrarMensaje Then
                MsgBox "No se encontraron Tipos de Cambio registrados en SUNAT para este mes." & vbNewLine & vbNewLine & _
                        "Tipo de Cambio no Actualizado.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
            End If
        End If
    Else
        If bolMostrarMensaje Then
            MsgBox "Conexión a Web de SUNAT fállida." & vbNewLine & _
                    "Pagina de SUNAT no disponible por el momento o " & vbNewLine & _
                    "problemas de conexión a Internet." & vbNewLine & vbNewLine & _
                    "Tipo de Cambio no Actualizado.", vbInformation + vbOKOnly, App.ProductName & "-" & wnomcia
        End If
    End If

    Set XML = Nothing

    Exit Function
errActualizarTCambioWebSunat:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open StrConexDbBancos

            Resume
        Case Else
            If bolMostrarMensaje Then
                MsgBox "Error No.: " & Err.Number & vbNewLine & _
                        "Descripción: " & Err.Description & vbNewLine & _
                        "Conexión a Web de SUNAT fállida." & vbNewLine & _
                        "Pagina de SUNAT no disponible por el momento o problemas de conexión a Internet." & vbNewLine & vbNewLine & _
                        "Tipo de Cambio no Actualizado.", vbExclamation + vbOKOnly, _
                        App.ProductName & " - ModUtilitario: ActualizarTCambioWebSunat"
            End If
    End Select

    actualizarTCambioWebSunat = False

    Err.Clear
End Function

Public Function actualizarTCambioWebSunatV2(ByVal strFechaConsulta As String, _
                                            Optional ByVal bolMostrarMensaje As Boolean) As Boolean
    On Error GoTo errActualizarTCambioWebSunatV2
    
    actualizarTCambioWebSunatV2 = False
    
    
    Dim strURL As String
    Dim objExplorador As InternetExplorer
    Dim objDocumentoHtml1 As HTMLDocument
    Dim objDocumentoHtml2 As HTMLDocument
    Dim objColeccion1 As IHTMLElementCollection
    Dim objColeccion2 As IHTMLElementCollection
    Dim objCelda As HTMLTableCell
    
    Dim strDia As String, strMes As String, strAnno As String
    
    Dim dblPosicion As Double, strCadenaExtraer As String, dblIndiceCelda As Double
    
    strDia = Trim(Str(Day(CDate(strFechaConsulta))))
    strMes = Trim(Str(Month(CDate(strFechaConsulta))))
    strAnno = Trim(Str(Year(CDate(strFechaConsulta))))
    
    strURL = "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
    
    Set objExplorador = New InternetExplorer
    
    With objExplorador
        .Navigate strURL
        .Visible = False
        
        While .Busy Or .readyState <> READYSTATE_COMPLETE: DoEvents: Wend
        
        Set objDocumentoHtml1 = .document
    End With
    
    With objDocumentoHtml1.selectForm
        .mes.selectedIndex = strMes
        .anho.value = strAnno
        
        .submit
    End With
    
    Set objDocumentoHtml2 = objExplorador.document
    
    While objExplorador.Busy Or objExplorador.readyState <> READYSTATE_COMPLETE: DoEvents: Wend
    
    Set objColeccion2 = objDocumentoHtml2.getElementsByTagName("form")
    
    For Each objCelda In objColeccion2
        If objCelda.sourceIndex = 9 Then
            dblPosicion = InStr(Trim(objCelda.innerText), "-")
            strCadenaExtraer = Mid(Trim(objCelda.innerText), 1, dblPosicion + 5)
        End If
    Next objCelda
    
    Set objColeccion1 = objDocumentoHtml2.getElementsByTagName("Td")
    
    objAyudaTCambio.inicializarEntidades
    
    Do While objAyudaTCambio.Cambio = 0
        For Each objCelda In objColeccion1
            'buscamos el dblIndiceCelda
            If Trim(objCelda.innerText) = strDia Then
                dblIndiceCelda = Val(objCelda.sourceIndex)
            End If
             
            Select Case objCelda.sourceIndex
                Case (dblIndiceCelda + 2)
                    objAyudaTCambio.TCCompra = Val(objCelda.innerText & "")
                Case (dblIndiceCelda + 3)
                    objAyudaTCambio.TCVenta = Val(objCelda.innerText & "")
                    objAyudaTCambio.Cambio = Val(objCelda.innerText & "")
            End Select
        Next objCelda
        
        If objAyudaTCambio.Cambio > 0 Then
            Exit Do
        End If
        
        'Disminuir el día, de uno en uno hasta encontrar el Ultimo día Registrado en Sunat
        strDia = Trim((Val(strDia) - 1) & "")
        'Salir si Día llega a ser '0' (Cero)
        If Val(strDia) = 0 Then Exit Do
    Loop
    
    objExplorador.Quit
    
    If objAyudaTCambio.Cambio > 0 Then
        With objAyudaTCambio
            .Fecha = strFechaConsulta
            
            strFechaConsulta = Format(strDia, "00") & "/" & Format(strFechaConsulta, "mm/yyyy")
            
            If .Fecha = strFechaConsulta Then
                .TcOficial = True
            End If
            
            If .guardarTCambio Then
                If bolMostrarMensaje Then
                    If .Fecha = strFechaConsulta Then
                        MsgBox "Tipo de cambio en SUNAT para hoy " & .Fecha & ":" & vbNewLine & _
                                "       - COMPRA    : " & .TCCompra & vbNewLine & _
                                "       - VENTA     : " & .TCVenta & vbNewLine & vbNewLine & _
                                "Actualizado en el sistema.", vbInformation + vbOKOnly, App.ProductName
                    Else
                        MsgBox "Ultimo Tipo de cambio registrado en SUNAT al " & strFechaConsulta & ":" & vbNewLine & _
                                "       - COMPRA    : " & .TCCompra & vbNewLine & _
                                "       - VENTA     : " & .TCVenta & vbNewLine & vbNewLine & _
                                "Actualizado en el sistema para hoy.", vbInformation + vbOKOnly, App.ProductName
                    End If
                End If
            End If
        End With
        
        actualizarTCambioWebSunatV2 = True
    Else
        If bolMostrarMensaje Then
            MsgBox "No se encontraron Tipos de Cambio registrados en SUNAT para este mes." & vbNewLine & vbNewLine & _
                    "Tipo de Cambio no Actualizado.", vbInformation + vbOKOnly, App.ProductName
        End If
    End If
    
    
    
    Exit Function
errActualizarTCambioWebSunatV2:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos 'StrConexDbBancos
            
            Resume
        Case Else
            If bolMostrarMensaje Then
                MsgBox "Error No.: " & Err.Number & vbNewLine & _
                        "Descripción: " & Err.Description & vbNewLine & _
                        "Conexión a Web de SUNAT fállida." & vbNewLine & _
                        "Pagina de SUNAT no disponible por el momento o problemas de conexión a Internet." & vbNewLine & vbNewLine & _
                        "Tipo de Cambio no Actualizado.", vbExclamation + vbOKOnly, _
                        App.ProductName & " - modUtilitario: ActualizarTCambioWebSunatV2"
            End If
    End Select
    
    actualizarTCambioWebSunatV2 = False
    
    Err.Clear
End Function

Public Function validarRUCenSunat(ByVal strRucConsulta As String, _
                                    Optional ByVal bolMostrarMensaje As Boolean) As Boolean
    
    On Error GoTo errValidarRUCenSunat
    
    validarRUCenSunat = False
    
    Dim strTexto As String
    Dim strPrincipio As String
    Dim strFinal As String
    Dim XML
    Dim dblPosicion1 As Double
    Dim dblPosicion2 As Double
    
    strTexto = vbNullString
    
    'Instanciar un Objeto XML de Microsoft
    Set XML = CreateObject("Microsoft.XMLHTTP")
    'Setear los parametros de Apertura del Objeto, para Consultar la URL
    XML.Open "POST", "http://www.sunat.gob.pe/w/wapS01Alias?ruc=" & strRucConsulta, False
    
    'Ejecutamos la Pagina en el Objeto
    XML.send
    'Capturamos el Texto de Respuesta albergado en el Objeto
    strTexto = XML.responseText
    
    If strTexto <> vbNullString Then
        If InStr(1, strTexto, "El numero Ruc ingresado es invalido") > 0 Then
            If bolMostrarMensaje Then
                MsgBox "El Número RUC ingresado es invalido, verifique.", vbInformation + vbOKOnly, App.ProductName
            End If
            
            validarRUCenSunat = False
        Else
            With objAyudaCliente
                .inicializarEntidades
                
                'Razon Social
                dblPosicion1 = InStr(strTexto, strRucConsulta)
                dblPosicion2 = InStr(strTexto, "<br/></small>")
                
                .NombreCliente = Mid(strTexto, dblPosicion1 + 14, (dblPosicion2 - dblPosicion1) - 14)
                
                'Direccion Fiscal
                dblPosicion1 = InStr(strTexto, "Direcci")
                dblPosicion2 = InStr(strTexto, "Situaci")
                
                .DireccionCliente = Mid(strTexto, dblPosicion1 + 24, (dblPosicion2 - dblPosicion1) - 52)
                
                'Estado
                dblPosicion1 = InStr(strTexto, "Estado.")
                dblPosicion2 = InStr(strTexto, "Agente")
                
                .CodigoVendedor = Mid(strTexto, dblPosicion1 + 11, (dblPosicion2 - dblPosicion1) - 53)
                .CodigoVendedor = IIf(left(.CodigoVendedor, 1) = "A", "ACTIVO", left(.CodigoVendedor, 8))
                
                'Situacion
                dblPosicion1 = InStr(strTexto, "Situaci")
                dblPosicion2 = InStr(strTexto, "Tel")
                
                .CodigoCobrador = Trim(Mid(strTexto, dblPosicion1 + 18, (dblPosicion2 - dblPosicion1) - 53))
                
                'Telefono
                dblPosicion1 = InStr(strTexto, "Tel")
                dblPosicion2 = InStr(strTexto, "Dependencia")
                
                .Telefono = left(Trim(Mid(strTexto, dblPosicion1 + 26, (dblPosicion2 - dblPosicion1) - 53)), 7)
                If Not IsNumeric(.Telefono) Then
                    .Telefono = vbNullString
                End If
            End With
            
            validarRUCenSunat = True
        End If
    Else
        If bolMostrarMensaje Then
            MsgBox "Conexión a Web de SUNAT fállida." & vbNewLine & _
                    "Pagina de SUNAT no disponible por el momento o " & vbNewLine & _
                    "problemas de conexión a Internet." & vbNewLine & vbNewLine & _
                    "Verificación Fallida.", vbInformation + vbOKOnly, App.ProductName
        End If
        
        validarRUCenSunat = False
    End If
    
    Set XML = Nothing
    
    Exit Function
errValidarRUCenSunat:
    Select Case Err.Number
        Case 3704, 3709
            cnn_dbbancos.Open cconex_dbbancos
            
            Resume
        Case Else
            If bolMostrarMensaje Then
                MsgBox "Error No.: " & Err.Number & vbNewLine & _
                        "Descripción: " & Err.Description & vbNewLine & _
                        "Conexión a Web de SUNAT fállida." & vbNewLine & _
                        "Pagina de SUNAT no disponible por el momento o problemas de conexión a Internet." & vbNewLine & vbNewLine & _
                        "Verificación Fallida.", vbExclamation + vbOKOnly, _
                        App.ProductName & " - ModUtilitario: ValidarRUCenSunat"
            End If
    End Select
    
    validarRUCenSunat = False
    
    Err.Clear
End Function

Public Function validarRUCenFactiliza(ByVal strRucConsulta As String, _
                                      Optional ByVal bolMostrarMensaje As Boolean) As Boolean

    On Error GoTo errValidarRUCenFactiliza

    validarRUCenFactiliza = False
    
    Dim strURL As String
    Dim xmlhttp As Object
    Dim jsonResponse As String

    ' Inicializar variables
    srazonSocial = ""
    snombreComercial = ""
    stelefono = ""
    scondicion = ""
    ssistElectronica = ""
    sdireccion = ""

    ' Actualiza la URL de la API para Factiliza
    strURL = "https://api.factiliza.com/pe/v1/ruc/info/" & strRucConsulta

    ' Crear objeto MSXML2.XMLHTTP para la solicitud (igual que el código que funciona)
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

    ' Abrir la solicitud (método GET) y agregar el token en la cabecera
    With xmlhttp
        .Open "GET", strURL, False
        '.setRequestHeader "Authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIyMTQiLCJuYW1lIjoiSnVhbiBDYXJsb3MiLCJlbWFpbCI6Imp1YW5jYXJsb3MuZ2lsYXJkaUBnbWFpbC5jb20iLCJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3dzLzIwMDgvMDYvaWRlbnRpdHkvY2xhaW1zL3JvbGUiOiJjb25zdWx0b3IifQ.QuyNKdJn0L6Nxp7f3pZRkUj7yRytOIqxkDMRIiDg1E0"
        .setRequestHeader "Authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIzNzc2NCIsIm5hbWUiOiJveGFwYW1wYS5wZXJ1LmdyYXRpc0BnbWFpbC5jb20iLCJlbWFpbCI6Im94YXBhbXBhLnBlcnUuZ3JhdGlzQGdtYWlsLmNvbSIsImh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vd3MvMjAwOC8wNi9pZGVudGl0eS9jbGFpbXMvcm9sZSI6ImNvbnN1bHRvciJ9.-jVgTk5XZCrkFvqelXUayRtBwrdB9Q-ZX7huNM_-Tq0"
        .send
    End With
    'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIzNzc2NCIsIm5hbWUiOiJveGFwYW1wYS5wZXJ1LmdyYXRpc0BnbWFpbC5jb20iLCJlbWFpbCI6Im94YXBhbXBhLnBlcnUuZ3JhdGlzQGdtYWlsLmNvbSIsImh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vd3MvMjAwOC8wNi9pZGVudGl0eS9jbGFpbXMvcm9sZSI6ImNvbnN1bHRvciJ9.-jVgTk5XZCrkFvqelXUayRtBwrdB9Q-ZX7huNM_-Tq0

    ' Obtener la respuesta en JSON
    jsonResponse = xmlhttp.responseText

    ' Comprobar el estado de la respuesta
    If xmlhttp.Status = 200 Then
        ' Extraer valores manualmente del JSON usando la función ExtractValue
        srazonSocial = ExtractValue(jsonResponse, "nombre_o_razon_social")
        scondicion = ExtractValue(jsonResponse, "condicion")
        sdireccion = ExtractValue(jsonResponse, "direccion_completa")

        validarRUCenFactiliza = True
    Else
        If bolMostrarMensaje Then
            MsgBox "Error al validar el RUC. Estado: " & xmlhttp.Status, vbExclamation, "Error en Factiliza"
        End If
    End If

    ' Limpiar el objeto xmlhttp
    Set xmlhttp = Nothing

    Exit Function

errValidarRUCenFactiliza:
    If bolMostrarMensaje Then
        MsgBox "Error No.: " & Err.Number & vbNewLine & _
               "Descripción: " & Err.Description & vbNewLine & _
               "Error al conectarse a la API de Factiliza.", vbExclamation + vbOKOnly, _
               App.ProductName & " - validarRUCenFactiliza"
    End If
    validarRUCenFactiliza = False
    Err.Clear
End Function
Private Function ExtractValue(ByVal jsonString As String, ByVal Key As String) As String
    Dim keyPosition As Long
    Dim startPosition As Long
    Dim endPosition As Long
    Dim value As String
    
    ' Localizar la clave en el string JSON
    keyPosition = InStr(jsonString, """" & Key & """")
    
    If keyPosition > 0 Then
        ' Encontrar el inicio del valor (después del ":")
        startPosition = InStr(keyPosition, jsonString, ":") + 1

        ' El valor puede estar entre comillas si es un string, así que encontramos el siguiente carácter de comillas
        If Mid(jsonString, startPosition, 1) = """" Then
            startPosition = startPosition + 1
            endPosition = InStr(startPosition, jsonString, """") ' Buscar el final del valor entre comillas
        Else
            ' Si no está entre comillas, buscar la coma o el cierre de llave que marca el final del valor
            endPosition = InStr(startPosition, jsonString, ",")
            If endPosition = 0 Then
                endPosition = InStr(startPosition, jsonString, "}")
            End If
        End If
        
        ' Extraer el valor completo
        value = Mid(jsonString, startPosition, endPosition - startPosition)
        value = Trim(value) ' Eliminar espacios en blanco
        value = Replace(value, """", "") ' Quitar comillas adicionales si las hubiera
        ExtractValue = value
    Else
        ExtractValue = "No encontrado"
    End If
End Function

Public Sub pulsarTecla(ByVal lngTecla As Long)
    Call keybd_event(lngTecla, 0, 0, 0)
    
                Call keybd_event(lngTecla, 0, KEYEVENTF_KEYUP, 0)
End Sub

Public Sub borrarTablaEnBD(ByVal conConexionDeBaseDeDatos As ADODB.Connection, _
                            ByVal strTabla As String)
    
    Dim objCatalogo As New ADOX.Catalog
    Dim objTabla As New Table
    
    objCatalogo.ActiveConnection = conConexionDeBaseDeDatos.ConnectionString
    
    For Each objTabla In objCatalogo.Tables
        If UCase(objTabla.Name) = UCase(strTabla) Then
            conConexionDeBaseDeDatos.Execute "DROP TABLE " & strTabla
            
            Exit For
        End If
    Next objTabla
End Sub

'***************************************************************************************************************
'***************************************************************************************************************
Public Sub Main()
    If cnRutas.State = 1 Then cnRutas.Close
    
    'abrirCnCpConfig
    
'    abrirCnIntermedio
    
    
    
    cnRutas.Open obtenerRuta(1, App.Path)
    
    Set objLogin = New frmLogin
    
    objLogin.Show
End Sub

Public Function obtenerRuta(ByVal intIndice As Integer, ByVal strRuta As String) As String
    Dim strBase As String
    
    Select Case intIndice
        Case 1 'rutas
            strBase = "\rutas.mdb"
        Case 2 'control
            strBase = "\control.mdb"
        Case 3 'bancos
            strBase = "\db_bancos.mdb"
        Case 4 'tabla
            strBase = "\DB_TABLA.mdb"
        Case 5 'ctrcom
            strBase = "\CTRCOM.MDB"
    End Select
    
    obtenerRuta = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRuta & strBase & ";Persist Security Info=False"
End Function

Public Function cargarParametrosEmpresa(ByVal strEmpresa As String) As Boolean
On Error GoTo errCargarParametros
    'Cargar Rutas de Acceso
    If rsRutas.State = 1 Then rsRutas.Close
    
    rsRutas.Open "Select Temporales, Bancos, Contabilidad, EmpSgte, ContaCnt, Envia, Tablas From Srutas Where Empresa = '" & _
                    strEmpresa & "' order by Empresa", cnRutas, adOpenForwardOnly, adLockReadOnly
    
    If Not rsRutas.EOF Then
        wempresa = strEmpresa
        wrutatemp = Trim(rsRutas!Temporales & "")
        wrutabancos = Trim(rsRutas!Bancos & "")
        wrutaconta = Trim(rsRutas!Contabilidad & "")
        wcontacnt = Trim(rsRutas!ContaCnt & "")
        wRutaEnvia = Trim(rsRutas!envia & "")
        If cnControl.State = 1 Then cnControl.Close
        
        cconex_control = obtenerRuta(2, App.Path)
        cnControl.Open cconex_control
        
        If cnn_control.State = 1 Then cnn_control.Close
        cnn_control.Open cconex_control
        
        If rscontrol.State = 1 Then rscontrol.Close
        
        rscontrol.Open "SELECT * " & _
                        "FROM SF1PARAM " & _
                        "WHERE F1CODEMP = '" & strEmpresa & "'", cnControl
        
        If Not rscontrol.EOF Then
            wIgv = Val(rscontrol!F1IGV & "")
            wprecio = Trim(rscontrol!F1PRECIO & "")
            wcolvalvta = Trim(rscontrol!F1COLVALVTA & "")
            walmagui = Trim(rscontrol!F1PARALMA_GUI & "")
            walmafac = Trim(rscontrol!F1PARALMA_FAC & "")
            walmabol = Trim(rscontrol!F1PARALMA_BOL & "")
            walmadeb = Trim(rscontrol!F1PARALMA_DEB & "")
            walmacre = Trim(rscontrol!F1PARALMA_CRE & "")
            wpartipcam = Trim(rscontrol!F1PARTIPCAM & "")
            wpardecimal = Trim(rscontrol!F1PARDECIMAL & "")
            wbasetemp = Trim(rscontrol!F1RUTABASE & "")
            wparact_stock = Trim(rscontrol!F1PARACT_STOCK & "")
            wvisualiza_act = Trim(rscontrol!F1VISUALIZA_ACT & "")
            wdireccion = Trim(rscontrol!F1DIREMP & "") & " - " & Trim(rscontrol!F1DISTRITO & "")
            wnomcia = Trim(rscontrol!F1NOMEMP & "")
            wtelefono = Trim(rscontrol!F1LOGEMP & "")
            wfax = Trim(rscontrol!F1NUMPED & "")
            wruc = Trim(rscontrol!F1RUCEMP & "")
            wimprimir_ruc = Trim(rscontrol!F1IMPRIMIR_RUC & "")
            wf1codcli_anulado = Trim(rscontrol!F1CODCLI_ANULADO & "")
            wf1codcon_anulado = Trim(rscontrol!F1CODCON_ANULADO & "")
            wf1formapago_contado = Trim(rscontrol!F1FORMAPAGO_CONTADO & "")
            wf1dscto_contado = Val(rscontrol!F1DSCTO_CONTADO & "")
            wf1puntodeventa = Trim(rscontrol!F1PUNTODEVENTA & "")
            wf1grababoleta = Trim(rscontrol!F1GRABABOLETA & "")
            wflete = Trim(rscontrol!F1FLETE & "")
            wpreciofac = Trim(rscontrol!F1PRECIOFAC & "")
            wf1monedalistaprecios = Trim(rscontrol!F1MONEDALISTAPRECIOS & "")
            wf1visualiza_det_lista = Trim(rscontrol!F1VISUALIZA_DET_LISTA & "")
            wf1uupp = Trim(rscontrol!F1UUPP & "")
            wf1evalua_linea_venc = Trim(rscontrol!F1EVALUA_LINEA_VENC & "")
            wf1formato_rv = Trim(rscontrol!F1FORMATO_RV & "")
            wparamcliente = Trim(rscontrol!F1PARAMCLIENTE & "")
            
            wparamprod = Trim(rscontrol!F1PARAMPROD & "")
            wf1anno = Trim(rscontrol!F1ANNO & "")
            wf1distrito = Trim(rscontrol!F1DISTRITO & "")
            wf1impresoras = Trim(rscontrol!F1IMPRESORAS & "")
            wf1elim_item_dcto = Trim(rscontrol!f1elim_item_dcto & "")
            wf1facturar_diario_correla = Trim(rscontrol!F1FACTURAR_DIARIO_CORRELA & "")
            wf1sistema_venta_proyectos = Trim(rscontrol!F1SISTEMA_VENTA_PROYECTOS & "")
            wf1control_menu = Trim(rscontrol!F1CONTROL_MENU & "")
            'wf1factor = Val(rsControl!F1FACTOR & "")
            wf1mant_productos = Trim(rscontrol!F1MANT_PRODUCTOS & "")
            wf1redondeo_dec1 = Trim(rscontrol!F1REDONDEO_DEC1 & "")
            wf1gen_vale_nc = Trim(rscontrol!F1gen_vale_nc & "")
            'wf1codigodscto = Trim(rsControl!F1CODIGODSCTO & "")
            wf1trasladactasxcob = Trim(rscontrol!F1TRASLADACTASXCOB & "")
            wf1evalua_stock = Trim(rscontrol!F1EVALUA_STOCK & "")
            
            wIndEnvia = Trim(rscontrol!F1ENVIA & "")
            wLocal = Trim(rscontrol!F4LOCAL & "")
            almaTrans = Trim(rscontrol!F1ALMATRANS & "")
            comisionVen = Val(rscontrol!F1COMISIONVEN & "")
        End If
        
        rscontrol.Close
        
        Set rscontrol = Nothing
        
        wbasetempND = wrutatemp & "\BASETEMPVENTASND.MDB"
        
        If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
        
        cconex_dbbancos = obtenerRuta(3, wrutabancos)
        cnn_dbbancos.Open cconex_dbbancos
        
        If cnn_dbtabla.State = 1 Then cnn_dbtabla.Close
        
        cnn_dbtabla.Open obtenerRuta(4, wrutabancos)
        
        cargarParametrosEmpresa = True
    Else
        MsgBox "Empresa no registrada.", vbInformation, App.ProductName
        
        cargarParametrosEmpresa = False
    End If
    
    rsRutas.Close
    
    Set rsRutas = Nothing
    
    Exit Function
errCargarParametros:
    MsgBox "Nro.: " & Err.Number & vbNewLine & _
            "Error: " & Err.Description & vbNewLine & _
            "Intente seleccionar nuevamente la Empresa.", vbInformation, App.ProductName & " - ModInfoPlus: F(x) CargarParametrosEmpresa"
    
    Err.Clear
    
    cargarParametrosEmpresa = False
End Function

Public Function verificarTipoCambio(ByVal StrFecha As String, ByVal cajaCompra As TextBox, ByVal cajaVenta As TextBox) As Boolean
    If rsTipoCambio.State = 1 Then rsTipoCambio.Close
    
    rsTipoCambio.Open "SELECT CAMBIOCOMP, CAMBIO_VENTA FROM CAMBIOS WHERE CDATE(FECHA) = CDATE('" & StrFecha & "')", _
                        cnn_dbbancos, adOpenDynamic, adLockBatchOptimistic
    
    If Not rsTipoCambio.EOF Then
        verificarTipoCambio = True
        cajaCompra.Text = Format(Val(rsTipoCambio!CAMBIOCOMP & ""), "#0.000")
        cajaVenta.Text = Format(Val(rsTipoCambio!CAMBIO_VENTA & ""), "#0.000")
    Else
        obtiene_tipodecambio (Date)
        
        If TCVenta > 0 Then
            verificarTipoCambio = True
            cajaCompra.Text = TCCompra
            cajaVenta.Text = TCVenta
        Else
            verificarTipoCambio = False
            cajaCompra.Text = Format(TCCompra, "#0.000")
            cajaVenta.Text = Format(TCVenta, "#0.000")
        End If
    End If
    
    
    rsTipoCambio.Close
    
    Set rsTipoCambio = Nothing
End Function

Public Sub obtiene_tipodecambio(fechax As Date)

Dim http As New XMLHTTP60, itm As String

    With http
        .Open "GET", "https://api.apis.net.pe/v1/tipo-cambio-sunat?fecha=" & Format(fechax, "yyyy-MM-dd"), False
        .send
        itm = .responseText
    End With
    
Dim p As Object
Set p = JSON.parse(itm)

TCCompra = p.ITEM("compra")
TCVenta = p.ITEM("venta")

Dim dia As String
Dim Af As New ADOFunctions
Dim intento As Integer
Dim Amov(0 To 10) As a_grabacion
Dim Rs As New ADODB.Recordset
Dim TC, Principio, Final, Texto As String
Dim objXMLHTTP, XML
Dim Posicion1, Posicion2 As Integer

'''''''On Error Resume Next
''''''''TC = "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
'''''''TC = "http://e-consulta.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
'''''''dia = Day(fechax)
'''''''
'''''''Final = "<br/></small>"
'''''''Set XML = CreateObject("Microsoft.XMLHTTP")
'''''''XML.Open "POST", TC, False
'''''''XML.send
'''''''Texto = XML.responseText
'''''''
''''''''Open App.Path & "\TC.txt" For Append As #1
''''''''Write #1, texto
''''''''Close #1
'''''''
'''''''Do While Not Posicion1 > 0
'''''''    Principio = Trim(str(dia) & "</strong>")
'''''''    Posicion1 = InStr(XML.responseText, Principio)
'''''''    dia = dia - 1
'''''''    If dia = 0 Then Exit Do
'''''''Loop
'''''''Posicion2 = InStr(XML.responseText, Final)
'''''''TCCompra = Val(Mid(XML.responseText, Posicion1 + 101, 10))
'''''''TCVenta = Val(Mid(XML.responseText, Posicion1 + 192, 10))

'If Err <> 0 Then
'    MsgBox "Verificar conexión a Internet"
'End If
''''''''''''''''''''Set XML = Nothing
''''''''''''''''''''
''''''''''''''''''''If Val(TCCompra) > 2 And Val(TCCompra) And Val(TCVenta) > 2 And Val(TCVenta) < 3 Then
''''''''''''''''''''    MsgBox "Tipo de cambio importado de Sunat para hoy " & fechax & ".          COMPRA: " & TCCompra & ". VENTA: " & TCVenta, vbInformation, "CONTROL Plus!"
''''''''''''''''''''Else
''''''''''''''''''''    intento = 0
''''''''''''''''''''    Do While TCCompra = 0 And intento < 10
''''''''''''''''''''        Set XML = CreateObject("Microsoft.XMLHTTP")
''''''''''''''''''''        XML.Open "POST", TC, False
''''''''''''''''''''        XML.send
'''''''''''''''''''' Texto = XML.responseText
''''''''''''''''''''        Principio = Trim(str(dia) & "</strong>")
''''''''''''''''''''        Posicion1 = InStr(XML.responseText, Principio)
''''''''''''''''''''        Posicion2 = InStr(XML.responseText, Final)
''''''''''''''''''''        TCCompra = Val(Mid(XML.responseText, Posicion1 + 101, 10))
''''''''''''''''''''        TCVenta = Val(Mid(XML.responseText, Posicion1 + 192, 10))
''''''''''''''''''''        Set XML = Nothing
''''''''''''''''''''        If Val(TCCompra) > 2 And Val(TCCompra) And Val(TCVenta) > 2 And Val(TCVenta) < 3 Then
''''''''''''''''''''            MsgBox "Tipo de cambio importado de Sunat para hoy " & fechax & ".          COMPRA: " & TCCompra & ". VENTA: " & TCVenta, vbInformation, "CONTROL Plus!"
''''''''''''''''''''        Else
''''''''''''''''''''            TCVenta = 0
''''''''''''''''''''        End If
''''''''''''''''''''        intento = intento + 1
''''''''''''''''''''    Loop
''''''''''''''''''''End If
    If TCVenta > 0 Then
        Amov(0).campo = "CAMBIO": Amov(0).valor = Val(TCVenta): Amov(0).Tipo = "N"
        Amov(1).campo = "FECHA": Amov(1).valor = CDate(fechax): Amov(1).Tipo = "F"
        Amov(2).campo = "CAMBIOCOMP": Amov(2).valor = Val(TCCompra): Amov(2).Tipo = "N"
        Amov(3).campo = "CAMBIO_VENTA": Amov(3).valor = Val(TCVenta): Amov(3).Tipo = "N"

        GRABA_REGISTRO Amov, "CAMBIOS", "A", 3, cnn_dbbancos.ConnectionString, ""
        End If
End Sub

Public Sub abrirPaginaWeb(ByVal URL As String)
    Call ShellExecute(0&, vbNullString, URL, vbNullString, _
                              vbNullString, vbNormalFocus)
End Sub
Public Function validarKeyPress(ByVal Key As Integer) As Integer
    Select Case Key
        Case 13
            SendKeys "{TAB}"
        Case 22, 39, 44, 58, 59
            MsgBox ("No acepta caracteres especiales como '|' ';'  ':'  ',' etc... Ingrese solo Caracteres o Digitos...")
            validarKeyPress = 0
        Case 8, 32, 45, 46, 47, 49 To 57, 65 To 90, 193, 201, 205, 211, 218, 209
            validarKeyPress = Key
        Case 97 To 122, 225, 233, 237, 243, 250, 241
            validarKeyPress = Key - 32
    End Select
End Function
