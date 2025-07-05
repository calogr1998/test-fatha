Attribute VB_Name = "mod_variables"
Option Explicit
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

Public X                            As Integer

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
Function Calcula_Cantidad(pcodigo As Variant, pmoneda As String, palmacen As String) As Double
Dim recset As ADODB.Recordset

Dim sql     As String
    
    sql = ""
    If pmoneda = "S" Then
        If ctipoadm_bd = "M" Then
            sql = sql & "SELECT IF3VALES.F5CODPRO,IF3VALES.F2CODALM, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD "
        Else
            sql = sql & "SELECT IF3VALES.F5CODPRO,IF3VALES.F2CODALM, Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD "
        End If
    Else
        If ctipoadm_bd = "M" Then
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD "
        Else
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD "
        End If
    End If
    sql = sql & "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F2CODALM = IF3VALES.F2CODALM) AND (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) "
    sql = sql & "Where IF3VALES.F5CODPRO = '" & pcodigo & "' and IF3VALES.f2codalm = '" & palmacen & "' "
    sql = sql & "GROUP BY IF3VALES.F5CODPRO, IF3VALES.F2CODALM;"
    
    Set recset = New ADODB.Recordset
    recset.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not recset.EOF Then
        Calcula_Cantidad = IIf(IsNull(recset.Fields("CANTIDAD")), "", recset.Fields("CANTIDAD"))
    End If
    recset.Close
End Function

Function Costo_Unitario(pcodigo As String, pfecha As Date, pmoneda As String) As Double
    Dim CosUni  As ADODB.Recordset
    Dim sql     As String
    
    sql = ""
    If ctipoadm_bd <> "M" Then
        If pmoneda = "S" Then
            sql = sql & "SELECT "
            sql = sql & "IF3VALES.F5CODPRO, "
            sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, "
            sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO, (IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
            sql = sql & "[VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
        Else
            sql = sql & "SELECT "
            sql = sql & "IF3VALES.F5CODPRO, "
            sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, "
            sql = sql & "Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALDOL*IF3VALES.F3CANPRO, (IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
            sql = sql & "[VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
        End If
        sql = sql & "FROM "
        sql = sql & "IF4VALES "
        sql = sql & "INNER JOIN IF3VALES ON (IF4VALES.F2CODALM = IF3VALES.F2CODALM) AND (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) "
        sql = sql & "Where "
        sql = sql & "IF3VALES.F4FECVAL <= CVDATE('" & Format(pfecha, "DD/MM/YYYY") & "') And "
        sql = sql & "IF3VALES.F5CODPRO = '" & pcodigo & "' "
        sql = sql & "GROUP BY "
        sql = sql & "IF3VALES.F5CODPRO"
    Else
        If pmoneda = "S" Then
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
            sql = sql & "SUM(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1))/SUM(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS COSTO_UNITARIO "
        Else
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALDOL*IF3VALES.F3CANPRO,(IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
            sql = sql & "SUM(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1))/SUM(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS COSTO_UNITARIO "
        End If
        sql = sql & "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F2CODALM = IF3VALES.F2CODALM) AND (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) "
        sql = sql & "Where IF3VALES.F4FECVAL <= '" & Format(pfecha, "DD/MM/YYYY") & "' And IF3VALES.F5CODPRO = '" & pcodigo & "' "
        sql = sql & "GROUP BY IF3VALES.F5CODPRO;"
    End If
    
    Set CosUni = New ADODB.Recordset
    CosUni.Open UCase(sql), cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not CosUni.EOF Then '
        '' ----SE TIENE UNA DUDA
       Costo_Unitario = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), 0, CosUni.Fields("COSTO_UNITARIO"))
       
       'Costo_Unitario = IIf(IsNull(CosUni.Fields("VALOR_VENTA")), 0, CosUni.Fields("VALOR_VENTA"))
    End If
    wCost = Costo_Unitario
    CosUni.Close
    
End Function
Function Costo_Unitario2(pcodigo As Variant, pmoneda As String) As Double()
Dim arreglo(1) As Double
Dim CosUni  As ADODB.Recordset
Dim sql     As String
    
    sql = ""
    If pmoneda = "S" Then
        If ctipoadm_bd = "M" Then
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
            sql = sql & " Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1))/Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS COSTO_UNITARIO "
        Else
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALVTA*IF3VALES.F3CANPRO,(IF3VALES.F3VALVTA*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, [VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
        End If
    Else
        If ctipoadm_bd = "M" Then
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALDOL*IF3VALES.F3CANPRO,(IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, "
            sql = sql & "Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALDOL*IF3VALES.F3CANPRO,(IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1))/Sum(IF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS COSTO_UNITARIO "
        Else
            sql = sql & "SELECT IF3VALES.F5CODPRO, Sum (IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3CANPRO,IF3VALES.F3CANPRO*-1)) AS CANTIDAD, Sum(IIF(LEFT(IF3VALES.F4NUMVAL,1) = 'I',IF3VALES.F3VALDOL*IF3VALES.F3CANPRO,(IF3VALES.F3VALDOL*IF3VALES.F3CANPRO)*-1)) AS VALOR_VENTA, [VALOR_VENTA]/[CANTIDAD] AS COSTO_UNITARIO "
        End If
    End If
    sql = sql & "FROM IF4VALES INNER JOIN IF3VALES ON (IF4VALES.F2CODALM = IF3VALES.F2CODALM) AND (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) "
    sql = sql & "Where IF3VALES.F5CODPRO = '" & pcodigo & "' "
    sql = sql & "GROUP BY IF3VALES.F5CODPRO;"
    
    Set CosUni = New ADODB.Recordset
    CosUni.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not CosUni.EOF Then
       arreglo(0) = IIf(IsNull(CosUni.Fields("COSTO_UNITARIO")), "", CosUni.Fields("COSTO_UNITARIO"))
       arreglo(1) = IIf(IsNull(CosUni.Fields("CANTIDAD")), "", CosUni.Fields("CANTIDAD"))
    End If
    Costo_Unitario2 = arreglo
    CosUni.Close
End Function

Sub ImprimeXY(pdata As Variant, ptipo As String, ptama As Integer, PFILA As Integer, pcolu As Integer)
    
    Dim wtemp   As String
    
    Select Case ptipo
        Case "0"     'texto
             Printer.CurrentY = PFILA
             Printer.CurrentX = pcolu
             Printer.Print left(pdata, ptama)
        Case "7"     'texto
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            Printer.Print left(Val(pdata), ptama)

        Case "1"      'entero
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            If Val(Format(pdata, "#0")) <> 0 Then
                Printer.Print Format$(Format(pdata, "####,##0"), "@@@@@@")
            End If
        Case "2"      'doble
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            If Val(Format(pdata, "#0.00")) <> 0# Then
                If Val(Format(pdata, "#0.00")) > 0# Then
                    Printer.Print Format$(Format(pdata, "#,###,##0.00"), "@@@@@@@@@")
                Else
                    Printer.CurrentX = pcolu - 2
                    Printer.Print "(" & Format$(Format(pdata, "#,###,##0.00"), "@@@@@@@@@") & ")"
                End If
            End If
        Case "3"      'fecha
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            Select Case ptama
                Case 1: Printer.Print right(Trim(pdata), 1)
                Case 2: Printer.Print right(Trim(pdata), 2)
                Case 4: Printer.Print left(Trim(pdata), 4)
                Case 10: Printer.Print cmes(Val(Trim(pdata)))
            End Select
        Case "4"     'memo
            Do While Len(Trim(pdata)) > 0
                Printer.CurrentY = PFILA
                Printer.CurrentX = pcolu
                If right(left(pdata, ptama), 1) = " " Or Len(pdata) < ptama Then
                    Printer.Print LTrim(left(pdata, ptama))
                    pdata = right(pdata, Len(pdata) - Len(LTrim(left(pdata, ptama))))
                Else
                    wtemp = left(pdata, ptama)
                    Do While right(wtemp, 1) <> " "
                        wtemp = left(wtemp, Len(wtemp) - 1)
                    Loop
                    Printer.Print LTrim(wtemp)
                    pdata = right(pdata, Len(pdata) - Len(wtemp))
                End If
                PFILA = PFILA + IIf(Rs.Fields("F1scale") = 6, 4, 1)
                If Len(Trim(pdata)) > 0 Then
                    gfilas = gfilas + IIf(Rs.Fields("F1scale") = 6, 4, 1)
                End If
            Loop
        Case "5"      'doble cuadruple para  el valvta...
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            If Val(Format(pdata, "#0.0000")) <> 0# Then
                If Val(Format(pdata, "#0.0000")) > 0# Then
                    Printer.Print Format$(Format(pdata, "###,##0.0000"), "@@@@@@@@@@@")
                Else
                    Printer.CurrentX = pcolu - 2
                    Printer.Print "(" & Format$(Format(pdata, "###,##0.0000"), "@@@@@@@@@@@") & ")"
                End If
            End If
        Case "6"      'triple
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            If Val(Format(pdata, "#0.000")) <> 0# Then
                If Val(Format(pdata, "#0.000")) > 0# Then
                    Printer.Print Format$(Format(pdata, "###,##0.000"), "@@@@@@@@@@")
                Else
                    Printer.CurrentX = pcolu - 2
                    Printer.Print "(" & Format$(Format(pdata, "###,##0.000"), "@@@@@@@@@@@") & ")"
                End If
            End If


        End Select

End Sub

Public Function VALIDA_UUPP(puupp As String)
Dim sw      As Boolean
Dim rst     As New ADODB.Recordset

    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from IF6DOCUM where F4LOCALIDAD='" & Trim(puupp) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        wcodlocalidad = Trim(rst!F4LOCALIDAD & "")
        wdeslocalidad = Trim(rst!F4DESCLOC & "")
        wdireccion = Trim(rst!F4DIRECCION & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_UUPP = sw

End Function

Public Function VALIDA_ALMACEN(pcodalmacen As String)
Dim sw_e    As Boolean

    If rsalmacen.State = adStateOpen Then rsalmacen.Close
    rsalmacen.Open "SELECT F2NOMALM FROM EF2ALMACENES WHERE F2CODALM='" & pcodalmacen & "'", cnn_dbbancos
    If Not rsalmacen.EOF Then
        wnomalmacen = Trim(rsalmacen.Fields("F2NOMALM") & "")
        sw_e = True
    Else
        sw_e = False
    End If
    rsalmacen.Close
    VALIDA_ALMACEN = sw_e

End Function

Public Sub Saldo_Inicial(pcodprod, pdesde, Optional palma)
    Dim ncanti      As Double
    Dim npesoi      As Double
    Dim ncant       As Double
    Dim npeso       As Double
    Dim dfecha      As Date
    Dim csql        As String
    Dim RSCONSULTA  As New ADODB.Recordset
    
    If ctipoadm_bd = "M" Then
        csql = "SELECT B.F4NUMVAL,A.F5CODPRO,B.F4FECVAL, " _
        & "IF(LEFT(C.F4NUMVAL,1)= 'I',C.F3CANPRO,0) AS ENTRADAK, " _
        & "IF (LEFT(C.F4NUMVAL,1) = 'S',C.F3CANPRO,0) AS SALIDAK " _
        & "FROM IF5PLA AS A,IF4VALES AS B, IF3VALES AS C " _
        & "WHERE A.F5CODPRO=C.F5CODPRO and A.F5CODPRO='" & pcodprod & "' AND B.F4NUMVAL=C.F4NUMVAL " _
        & "AND B.F4FECVAL< '" & pdesde & "'AND B.F2CODALM=C.F2CODALM AND B.F2CODALM='" & palma & "' ORDER BY B.F4FECVAL"
    Else
        csql = "SELECT B.F4NUMVAL,A.F5CODPRO,B.F4FECVAL, " _
        & "IIF(LEFT(C.F4NUMVAL,1)= 'I',C.F3CANPRO,0) AS ENTRADAK, " _
        & "IIF (LEFT(C.F4NUMVAL,1) = 'S',C.F3CANPRO,0) AS SALIDAK " _
        & "FROM IF5PLA AS A,IF4VALES AS B, IF3VALES AS C " _
        & "WHERE a.f5codpro=c.f5codpro and A.F5CODPRO='" & pcodprod & "' AND B.F4NUMVAL=C.F4NUMVAL " _
        & "AND B.F4FECVAL< CVDATE('" & pdesde & "')AND B.F2CODALM=C.F2CODALM AND B.F2CODALM='" & palma & "' ORDER BY B.F4FECVAL"
    End If
    
    ncant = 0#
    If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
    RSCONSULTA.Open csql, cnn_dbbancos
    If Not RSCONSULTA.EOF Then
        RSCONSULTA.MoveFirst
        Do While Not RSCONSULTA.EOF
            If "" & left(RSCONSULTA.Fields("F4NUMVAL"), 1) = "I" Then
                ncant = ncant + CDbl("" & RSCONSULTA.Fields("entradak"))
            Else
                ncant = ncant - CDbl("" & RSCONSULTA.Fields("salidak"))
            End If
            RSCONSULTA.MoveNext
        Loop
    End If
    RSCONSULTA.Close
    
    'wcant = ncanti + ncant
    wcant = ncant
    
End Sub

Public Function Devuelve_Path(quebusco As String)
    
    Select Case quebusco
        Case ""
            Devuelve_Path = IIf(right(App.Path, 1) = "\", App.Path, App.Path & "\")
        Case "BD"
            Devuelve_Path = IIf(right(App.Path, 1) = "\", App.Path & wempresa & "\", App.Path & "\" & wempresa & "\")
        Case "BMP"
            Devuelve_Path = IIf(right(App.Path, 1) = "\", App.Path, App.Path & "\Bmps_Iconos\")
    End Select
    
End Function

Public Function VALIDA_CC(pcentro As String)
Dim sw      As Boolean
Dim rst     As New ADODB.Recordset

    sw = False
    If rst.State = adStateOpen Then rst.Close
    rst.Open "Select * from CENTROS where F3COSTO='" & Trim(pcentro) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        wcodcosto = pcentro
        wdescosto = Trim(rst!F3DESCRIP & "")
        sw = True
    Else
        sw = False
    End If
    rst.Close
    VALIDA_CC = sw

End Function

Private Function getCantidad(prod As String, umVen As String, Cantidad As String)
    Dim um      As String
    Dim factor  As Double
    Dim rptFun  As Double
    um = traerCampo("IF5PLA", "F7CODMED", "F5CODPRO", prod)
    If UCase(um) <> UCase(umVen) Then
        factor = traerCampo("MEDIVENTAS", "F5FACTOR", "F5CODPRO", prod, "and F7CODMED = '" & umVen & "'")
        If Val(factor) > 0 Then
            rptFun = Cantidad * factor
        Else
            rptFun = Cantidad
        End If
    Else
        rptFun = Cantidad
    End If
    getCantidad = rptFun
End Function

