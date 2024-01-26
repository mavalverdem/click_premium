Attribute VB_Name = "modMain"
Option Explicit

Public gsCodEmpCompass As String

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public ps_WinSystem As String ' Directorio de sistema
Public gsIdioma As String * 1 ' Idioma del sistema (Español o Ingles)
Public gbEnPcc As Boolean     'En uso en la empresa cliente (producción).
Public gbEsUsr As Boolean     'En uso por un usuario.

Public ps_Plataforma As String    ' Plataforma del servidor de base de datos
Public ps_Provider As String      ' Proveedor de base de datos
Public CONNSTRG As String         ' Cadena de conexion
Public ps_Servidor As String      ' Servidor de base de datos
Public ps_UserId As String        ' Usuario de conexion de base de datos
Public ps_Password As String      ' Paswword de conexion de base de datos
Public ps_Licencia As String      ' Licencencia de usos del sistema
Public ps_Puerto As String        ' Puerto
Public gsNomBDC As String         ' Base de datos de configuración.

'ini 2015-01-14 conver a sql
'2013-05-22 config..
Public gsNomBDC_Exte As String    ' Base de datos de configuración simbolo externo.
'fin 2015-01-14 conver a sql

Public gsNomBDS As String         ' Base de datos del sistema
Public ps_Prefijo As String       ' Prefijo de tabla temporal

Public gsRutSis As String     'Ruta Base del Sistema.
Public gsRutBDC As String     'Ruta de la BD de Configuración.
Public gsRutBDS As String     'Ruta de la Base de Datos.
Public gsRutRpt As String     'Ruta de los Reportes.
Public gsCodSis As String     'Código del Sistema.
Public gsNomSis As String     'Nombre del Sistema.
Public gsCodUsr As String     'Código del Usuario Activo.
Public gsAbvUsr As String     'Abreviación del Usuario Activo.
Public gsCodEmp As String     'Código de la Empresa Activa.
Public gsRazEmp As String     'Razón Social de la Empresa Activa.
Public gsRUCEmp As String     'Número de RUC de la Empresa Activa.
Public gsDirEmp As String     'Direccion de la Empresa Activa.
Public gsLocEmp As String     'Localidad de la empresa activa
Public gsGirEmp As String     'Giro o Actividad de la Empresa Activa.
Public gsRepEmp As String     'Representante Legal de la Empresa Activa.
Public gsRepDNIEmp As String  ' DNI Representante Legal de la Empresa Activa.
Public gsConEmp As String     'Contador de la Empresa Activa.
Public gsConDNIEmp As String  'DNI Contador de la Empresa Activa.
Public gsBuenContriEmp As String  'Buen contribuyente de la Empresa Activa.  '2015-08-27 ctr obligac sunat

Public gsAnoAct As String     'Año Activo.
Public gsMesAct As String     'Mes Activo.
Public gsMesApe As String     'Mes Apertura
Public gsMesCie As String     'Mes Cierre
Public gnFrances As String    'Proceso ejercicio frances
Public gnIndPedido As String  'filtrar pedido de compra
Public gnProDestino As Byte   ' Proceso detallado cuenta destino
Public gsAcceso As String     'Variable para verificar se selecciono empresa
Public aLabel() As String     ' Array de las etiquetas de los formularios

Public gsNvlUsr As String
Public Const NvlUsr_Adm As Byte = 0, _
             NvlUsr_Sup As Byte = 1, _
             NvlUsr_Asi As Byte = 2

' Variables de texto
Public NvlUsr_Adm_Txt As String, _
       NvlUsr_Sup_Txt As String, _
       NvlUsr_Asi_Txt As String

Public gbPms01 As Boolean    'Permiso 01 Nuevo/Varios.
Public gbPms02 As Boolean    'Permiso 02 Corregir/Varios.
Public gbPms03 As Boolean    'Permiso 03 Eliminar/Varios.
Public gbPms04 As Boolean    'Permiso 04 Vista Previa/Varios.
Public gbPms05 As Boolean    'Permiso 05 Imprimir/Varios.
Public gbPms06 As Boolean    'Permiso 06 Exportar/Varios.
Public gbPms07 As Boolean    'Permiso 07 Varios.
Public gbPms08 As Boolean    'Permiso 08 Varios.
Public gbPms09 As Boolean    'Permiso 09 Varios.
Public gbPms10 As Boolean    'Permiso 10 Varios.

Public a_Sufijo(31) As String

'ini 2015-01-14 conver a sql
'2012-12-19
'configura el tipo de cliente segun catalogo de pTCli001
'esto para perfilar segun configuracion Interback o tipo normal
Public gbTCli As String
Public Const pTCli001 As String = "001" 'Segun Interbank
Public Const pTCli002 As String = "002" 'Segun Compass y otros


Public Const pLenSerDoc As Integer = 4        ' Modificar longitud de serie
'fin 2015-01-14 conver a sql
Public Const pSrvSql As String = "Sql"
Public Const pSrvMySql As String = "MySql"

'             pFileCfg As String = "OwlOContnvo.ini", _
'             pFileCfg As String = "OwlOContsql.ini", _
'fin 2015-01-14 conver a sql

Public Const AYULLA As String = "L", _
             AYUDAT As String = "D", _
             FORMATOLONGTIME = "hh:mm:ss AMPM", _
             PRN_DEST_GRAF As String = "G", _
             PRN_DEST_MATR As String = "M", _
             PRN_ORIE_VERT As String = "V", _
             PRN_ORIE_HORI As String = "H"
             
Public Const pFileCfg As String = "owloContnvo.ini"
Public Const pFileCfg_sql As String = "owloContsql.ini" '2015-01-14 conver a sql

Public TEXT_NUEVO As String, _
       TEXT_MODIF As String, _
       TEXT_BUSCA As String, _
       TEXT_1021 As String, _
       TEXT_1022 As String, _
       TEXT_1031 As String, _
       TEXT_3001 As String, _
       TEXT_3101 As String, _
       TEXT_6001 As String, _
       TEXT_6002 As String, _
       TEXT_8001 As String, _
       TEXT_8002 As String, _
       TEXT_8003 As String, _
       TEXT_8004 As String, _
       TEXT_8005 As String, _
       TEXT_8006 As String, _
       TEXT_8007 As String, _
       TEXT_8008 As String, _
       TEXT_8009 As String, _
       TEXT_8010 As String

Public CtaAuxiliar As String
Public DesAuxiliar As String

Public proceso As Boolean
Public Activar As Boolean
Public ayudaban As Boolean
Public xqmes As String
Public rTitulo As String
Public xIndicador As String
    
'esta variable controla salir de form
'desde el load, con Unload Me
'asi bloqueo el show me de funcion
'ppCapturaTitulo que llama a todos los form del menu
'el problema es que llamada dos vecea al load
'Public Const pExitForm As Integer = 0        '2015-05-18 validacion
Public pExitForm  As Integer       '2015-05-18 validacion
'pExitForm=0 no hace show
'pExitForm=1 si hace show
    
'2016-01-29 version sistema contable, caso Esta.RiveNavar error acceso
'Public Const pSisVer300 As Boolean = True 'solo para expresas antiguas error acceso al entrar
Public Const pSisVer300 As Boolean = False

Sub Main()

'ini 2015-01-14 conver a sql
    '2012-12-19 configura modalidad de uso
    'gbTCli = pTCli001
    gbTCli = pTCli002
'fin 2015-01-14 conver a sql

gsCodEmpCompass = CODEMP_COMPASS ' 001 si es estado compass CODEMP_NORMAL=000
    
  Dim dsBuffer As String, dnSize As Long
  Dim dofsoFileCfg As New FileSystemObject, dotstFileCfg As TextStream
  Dim dsRutDrv As String, dsFileCfg As String, psLinea As String, ps_Driver As String
  
  Dim porstOpciones As ADODB.Recordset
  Dim n_Index As Long
  
  ' Seleciono el idioma del sistema
  frmIdioma.Show vbModal
  ' Cargo las variable de Textos
  Mensajes
  '2014-04-04 codigos de detraccion
  MatrizDetraccion
  Mensajes2 '2014-07-18
  If gsIdioma = NvlUsr_Asi Then
    For dnSize = 1 To 31
      a_Sufijo(dnSize) = Choose(dnSize, "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th", _
                                "th", "th", "th", "th", "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "st")
    Next dnSize
  End If

  '[Al compilar para el usuario colocar ambos en: True
  gbEnPcc = True: gbEsUsr = True
  ']
  
  ' Reconoce el directorio de windows.
  dsBuffer = Space$(255)
  dnSize = Len(dsBuffer)
  GetSystemDirectory dsBuffer, dnSize
  dsBuffer = Left(dsBuffer, Len(Trim(dsBuffer)) - 1)
  ps_WinSystem = dsBuffer
  ps_Prefijo = ""
  
  'Captura unidad de drive.
  If gbEnPcc And gbEsUsr Then
    ' Verifico que exista el Archivo de Configuracion
    dsFileCfg = ps_WinSystem & "\" & pFileCfg
    If StrConv(Dir$(dsFileCfg, vbHidden), vbLowerCase) <> LCase(pFileCfg) Then
      MsgBox TEXT_8002, vbCritical
      End
    End If
    
    ' Abro Archivo de Configuracion
    Set dotstFileCfg = dofsoFileCfg.OpenTextFile(dsFileCfg, ForReading)
    Do While Not dotstFileCfg.AtEndOfStream
      psLinea = dotstFileCfg.ReadLine
      If Left$(psLinea, 14) = "[Contabilidad]" Then dsRutDrv = Trim(Mid$(psLinea, InStr(psLinea, "=") + 1))
      If Left$(psLinea, 11) = "[Proveedor]" Then ps_Provider = Mid$(psLinea, InStr(psLinea, "=") + 1)
      If Left$(psLinea, 8) = "[Server]" Then ps_Servidor = Mid$(psLinea, InStr(psLinea, "=") + 1)
      If Left$(psLinea, 8) = "[UserId]" Then ps_UserId = Mid$(psLinea, InStr(psLinea, "=") + 1)
      If Left$(psLinea, 10) = "[Password]" Then ps_Password = Mid$(psLinea, InStr(psLinea, "=") + 1)
      If Left$(psLinea, 11) = "[BaseDatos]" Then gsNomBDC = Mid$(psLinea, InStr(psLinea, "=") + 1)
      If Left$(psLinea, 10) = "[Licencia]" Then ps_Licencia = Mid$(psLinea, InStr(psLinea, "=") + 1)
      If Left$(psLinea, 8) = "[Puerto]" Then ps_Puerto = Mid$(psLinea, InStr(psLinea, "=") + 1)
      
      If pFileCfg = pFileCfg_sql Then
      'sql2008 2012-09-25 gsNomBDS
      If Left$(psLinea, 15) = "[BaseDatoConta]" Then gsNomBDS = Mid$(psLinea, InStr(psLinea, "=") + 1)
      End If
    Loop

    dotstFileCfg.Close
    ' Verifico los parametrso de conexion y genero cadena
    
    If ps_Provider = "" Then MsgBox TEXT_8003, vbCritical: End
    If ps_Servidor = "" Then MsgBox TEXT_8003, vbCritical: End
    If ps_UserId = "" Then MsgBox TEXT_8003, vbCritical: End
    If gsNomBDC = "" Then MsgBox TEXT_8003, vbCritical: End
    If Mid(dsRutDrv, 2, 1) <> ":" Or Right(dsRutDrv, 1) <> "\" Then MsgBox TEXT_8003, vbCritical: End
    
    ps_Driver = "{MySQL ODBC 3.51 Driver}"
    
    ' Genero la cadena de conexion segun plataforma
    Select Case UCase$(ps_Provider)
     Case "SQLOLEDB.1", "MSDAORA", "MICROSOFT OLE DB PROVIDER FOR SQL SERVER"
      CONNSTRG = "Provider=" & ps_Provider & ";Data Source=" & ps_Servidor & ";User Id=" & ps_UserId & ";Password=" & ps_Password & ";Persist Security Info=False;Initial Catalog="
      ps_Plataforma = pSrvSql
      ps_Prefijo = "#"
'ini 2015-01-14 conver a sql
      'ini sql8 03/07/12
     Case "SQLNCLI10.1"
     'sql8 16/08/12Case "SQLNCLI10"
      'sql8 16/08/12CONNSTRG = "Provider=" & ps_Provider & ";Server=" & ps_Servidor & ";Uid=" & ps_UserId & ";pwd=" & ps_Password & ";Persist Security Info=False;Database="
      
      ''''''no valeCONNSTRG = "Provider=SQLNCLI10;Server=" & ps_Servidor & ";Uid=" & ps_UserId & ";pwd=" & ps_Password & ";Persist Security Info=False;Database="
      'SQLNCLI10.1
      'CONNSTRG = "Provider=" & ps_Provider & ";Data Source=" & ps_Servidor & ";User Id=" & ps_UserId & ";Integrated Security='';Persist Security Info=False;Initial File Name='';Server SPN='';Initial Catalog="
      '2012-08-25 CONNSTRG = "Provider=" & ps_Provider & ";Data Source=" & ps_Servidor & ";User Id=" & ps_UserId & ";Password=" & ps_Password & ";Integrated Security='';Persist Security Info=False;Initial File Name='';Server SPN='';Initial Catalog="
      
      'ori 2013-04-30 original
      'CONNSTRG = "Provider=" & ps_Provider & ";Data Source=" & ps_Servidor & ";User Id=" & ps_UserId & ";Password=" & ps_Password & ";Persist Security Info=False;Initial Catalog="
      '2013-12-19 (fuentes server) se quita comentario para conextarse como compass, conexion origen
      If gbTCli = pTCli002 Then
        CONNSTRG = "Provider=" & ps_Provider & ";Data Source=" & ps_Servidor & ";User Id=" & ps_UserId & ";Password=" & ps_Password & ";Persist Security Info=False;Initial Catalog="
      End If
      
      '2013-04-30 cambiar a seguridad integrada para banco interbak
      'segun henry
      'CONNSTRG = "Provider=" & ps_Provider & ";Data Source=" & ps_Servidor & ";Persist Security Info=False;Integrated Security=SSPI;Initial Catalog="
      
      '2013-05-05 el proveedor que quiere el banco es:
      'asi es como lo dio el bco Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & argBaseDato & " ;Data Source=" & argServidor
      'CONNSTRG = "Provider=" & ps_Provider & ";Integrated Security=SSPI;Persist Security Info=False" & " ;Data Source=" & ps_Servidor & ";Initial Catalog=" ' & argBaseDato
      '2013-12-19 (fuentes server) configuro para que funcion segun tipo de cliente
      If gbTCli = pTCli001 Then
       CONNSTRG = "Provider=" & ps_Provider & ";Integrated Security=SSPI;Persist Security Info=False" & " ;Data Source=" & ps_Servidor & ";Initial Catalog=" ' & argBaseDato
      End If
      'ps_Plataforma = pSrvSql8
      ps_Plataforma = pSrvSql
      ps_Prefijo = "#"
      'fin sql8 03/07/12
'fin 2015-01-14 conver a sql
     Case "MSDASQL.1"
     If ps_Puerto = "" Then
      CONNSTRG = "driver=" & ps_Driver & ";server=" & ps_Servidor & ";user=" & ps_UserId & ";password=" & ps_Password & ";option=3;persist security info=False;database="
     Else
      CONNSTRG = "driver=" & ps_Driver & ";server=" & ps_Servidor & ";user=" & ps_UserId & ";password=" & ps_Password & ";Port=" & ps_Puerto & ";option=3;persist security info=False;database="
     End If
      ps_Plataforma = pSrvMySql
     Case Else
      MsgBox TEXT_8003, vbCritical: End
    End Select
  End If
  Set dofsoFileCfg = Nothing
  Set dotstFileCfg = Nothing
  
  '[ARREGLAR***
  gsAnoAct = "2003"
  gsTpoMon_Sgn_MN = "S/."
  gsTpoMon_Sgn_ME = "US$"
  ']ARREGLAR***
  Dim xx_version
 ' xx_version = "v3.151w8" '28-09-2015 teo opc tra/honorario tc valida por f.oper por que se inst en exter, para  f.emis
  'xx_version = "v3.152w8" 'corre utilitario/transf exporta-ventas incluy detracc,consta. depos
  'xx_version = "v3.153w8" 'rpt utili/nueva empres- año/exporta excel consoli empresa
  'xx_version = "v3.155w7" 'rpt utili/nueva empres- año/exporta excel balance del mes
  'xx_version = "v3.156w7" 'rpt libros y reg ple compra arch plano
  'xx_version = "v3.157w7" 'rpt libros y reg ple venta arch plano
  'xx_version = "v3.158w7" 'rpt libros y reg ple venta arch plano
  'xx_version = "v3.159w7" 'rpt libros y reg ple venta adicion codmon
  'xx_version = "v3.160w7" 'rpt libros y reg ple diario y mayor adicion codmon en diario
  'xx_version = "v3.161w7" 'rpt libros y reg ple diario y mayor adicion codmon en diario
  'xx_version = "v3.162w7" 'rpt saco combos ples del compras y pongo ayuda con text y boton
  'xx_version = "v3.163w7" 'rpt correccion vta,cpr, ple
  'xx_version = "v3.164w7" 'rpt correccion vta,cpr, ple
  'xx_version = "v3.165w7" 'rpt correccion vta,cpr, ple
  'xx_version = "v3.165w7" 'rpt modifica combo ventas, cambio por tex y ayuda
  'xx_version = "v3.166w7" 'modicacion ple por teo, me envio la fuentes nuevas 206-02-10
  'xx_version = "v3.167w7" 'modifica transac.caja y banco codmon con ayuda tex y btn ayu
  'xx_version = "v3.168w7" 'modifica transac.diario codmon con ayuda tex y btn ayu y arre ple cpr no domi interfondo
  'xx_version = "v3.169w7" 'corrigio nombre archivos ple
  'xx_version = "v3.170w7" 'corrige honorario tpo cambio por f.emision
  'xx_version = "v3.171w7" 'adicion control de glosa ventas con limite max ingreso
  'xx_version = "v3.172w7" 'correcc archivo plano ple, revision bakup
  'xx_version = "v3.173w7" 'revision diario, mayor ple. hornorio esta otros empresas
  'xx_version = "v3.174w7" 'revis reporte tras-diario y reg compras duplica
  'xx_version = "v3.175w7" 'revis reporte diario y mayor ple error
  'xx_version = "v3.176w7" 'erro rpt bitacora, ccosto, adicion detracc vta, cpr ayuda salio combo
  '  xx_version = "v3.178w7" 'solucion ple compra tip.doc ref
  'xx_version = "v3.177w7" 'correc.teo del ple dia.y may., me paso el form frmLibros se chanco en fuente
  'xx_version = "v3.178w7" 'correc.interfase ventas
  'xx_version = "v3.179w7" 'correc.ple dia, may y Daort
  'xx_version = "v3.180w7" 'correc.ple dia lili encontro error colum 12
  'xx_version = "v3.181w7" 'corrreccion glosa cuenta interna singno mas a 75 carac
  'xx_version = "v3.182w7" 'corrreccion ple compra codtdc=91,97,98
  'xx_version = "v3.183w7" 'formato mashall imprime venta
  'xx_version = "v3.184w7" 'cambio valida indajd=1 tpoanl=2 trasc diario mone cta error
  'xx_version = "v3.184w7" 'cambio compras ple no domicilia y util-tran-plan cta
  'xx_version = "v3.185w7" 'cambio compras ple no domicilia y util-tran-plan cta
  'xx_version = "v3.187w7" 'correccion util/tranfere macro compras y ventas
  'xx_version = "v3.188w7" 'correccion PLE compra adici campo imptcb_inv,feedoc_inv
  'xx_version = "v3.189w7" 'correccion tamaño det cta ventas y reporte quitar mashall
  'xx_version = "v3.190w7" 'seguridad asistente no permite borrar
  'xx_version = "v3.191w7" 'correccion ple mayor-nueva empresa-campos estemp, filtros de respor estaemp=A
  'xx_version = "v3.192w7" 'error mayor ple corregidof
  'xx_version = "v3.193w7" 'error mayor ple corregido
  'xx_version = "v3.194w7" 'reporte prespuesto para interfondo y cambio ple teo
  'xx_version = "v3.195w7" 'correcion codmon asto destino y dif cam
  'xx_version = "v3.196w7" 'correcion ple 8.1 y 8.2
  'xx_version = "v3.197w7" 'correcion ple diario y mayoe sacar caracatere especiales a col 23
  'xx_version = "v3.198w7" 'correccion error campos numerico ple
  'xx_version = "v3.199w7" 'corre ppto interfondo error
  'xx_version = "v3.200w7" 'corre ppto interfondo error
  'xx_version = "v3.201w7" 'corre ppto interfondo error
  'xx_version = "v3.202w7" 'correccion ple compras no domiciliado
  'xx_version = "v3.203w7" 'correccion ple compras no domiciliado y ppto ayuda cco
  'xx_version = "v3.204w7" 'correccion ple compras no domiciliado compra
  'xx_version = "v3.205w7" 'correccion ple compras no domiciliado compra
  'xx_version = "v3.206w7" 'correccion reg compr exporta excel error duplicidad doc y ple 8.2
  'xx_version = "v3.207w7" 'correccion exporta excel auxiliar tabla
  'xx_version = "v3.208w7" 'errro apertura recordset auxilia tabla
  'xx_version = "v3.209w7" 'expota caja y bancos excel
  'xx_version = "v3.210w7" 'correccion error compra , vtas, hono transacc. error cabecera -detalle sdo
  'xx_version = "v3.211w7" 'corregir ple comp, dia. mayor
  xx_version = "v3.212w7" 'corregir ple comp, dia. mayor
  gsNomSis = Choose(gsIdioma, "Contabilidad - Click Premium " & xx_version, "Accounting - Click Premium " & xx_version)
  gsCodSis = "CO"
  '[ARREGLAR***
  gsNomBDS = "sysmacon"
  ']ARREGLAR***
  gsRutSis = IIf(gbEnPcc, dsRutDrv, dsRutDrv & "OWL-Paqu\OCont\") ' & gsCodSis & "\")
  gsRutRpt = IIf(gbEnPcc, gsRutSis & "repo\", dsRutDrv & "OWL-Paqu\OCont\Repo\")
  gsRutBDC = IIf(gbEnPcc, gsRutSis, dsRutDrv & "OWL-Paqu\OCont\") & "Data\a" & gsAnoAct & "\"
  gsRutBDS = IIf(gbEnPcc, gsRutSis, dsRutDrv & "OWL-Paqu\OCont\") & "Data\a" & gsAnoAct & "\"
'ini 2015-01-14 conver a sql
'2013-05-22 ini asignacion variable de base dato config externo
  If ps_Plataforma = pSrvMySql Then
    gsNomBDC_Exte = gsNomBDC & IIf(Len(Trim(gsNomBDC)) <> 0, ".", "")
  Else
    gsNomBDC_Exte = gsNomBDC & IIf(Len(Trim(gsNomBDC)) <> 0, "..", "")
  End If
'2013-05-22 fin asignacion variable de base dato config externo
'fin 2015-01-14 conver a sql

  ' Pantalla de usuario
  frmLogin.Show vbModal
  If Not frmLogin.ubCorrecto Then
    'Falla al iniciar la sesión, se sale de la aplicación
    End
  End If
  Unload frmLogin
    
  '[   Miguel Angel
  ' Cargo la pantalla de seleccion de empresa y periodos
  gsAcceso = "N"
  frmOSelEmp.Show vbModal
   
  If gsAcceso = "N" Then
    ' No seleccionada, se sale de la aplicación
    End
  End If
  ']
  
'ini 2015-08-27 ctr obligac sunat
If pSisVer300 = False Then
gfctrl_obliga_sunat
End If
'fin 2015-08-27 ctr obligac sunat
  
  '[Propio del Proyecto.
  gpCamposSaldos
  gpCieMes
  ']
  frmMain.Caption = gsNomSis
  Load frmMain
  '[ Cargo los datos del menu
  Set porstOpciones = New ADODB.Recordset
  porstOpciones.ActiveConnection = CONNSTRG & gsNomBDC
  porstOpciones.Source = "SELECT opcion, orden, codmdl, nommdl, "
  porstOpciones.Source = porstOpciones.Source & Choose(gsIdioma, "detmdl", "detmdlx") & " AS  descripcion "
  porstOpciones.Source = porstOpciones.Source & "FROM sgmdl "
  porstOpciones.Source = porstOpciones.Source & "WHERE codsis='" & gsCodSis & "' "
  porstOpciones.Source = porstOpciones.Source & "ORDER BY opcion, orden, codmdl"
  porstOpciones.CursorType = adOpenStatic
  porstOpciones.LockType = adLockReadOnly
  porstOpciones.Open
  
  frmMain.LblTitu(0).Caption = Choose(gsIdioma, "Periodo :", "Period :")
  frmMain.LblTitu(1).Caption = Choose(gsIdioma, "Mes :", "Month :")
  ' Datos menus principales
  frmMain.mnuTransacciones.Caption = Choose(gsIdioma, "&Transacciones", "&Transactions")
  frmMain.mnuReportes.Caption = Choose(gsIdioma, "&Reportes", "&Reports")
  frmMain.mnuProcesos.Caption = Choose(gsIdioma, "&Procesos", "&Processes")
  frmMain.mnuMaestros.Caption = Choose(gsIdioma, "Ta&blas", "Ta&bles")
  frmMain.mnuUtil.Caption = Choose(gsIdioma, "&Utilitarios", "T&ools")
  ' Datos sub menus
  frmMain.mnuDro.Caption = Choose(gsIdioma, "&Diarios", "&Journals")
  frmMain.mnuRMay.Caption = Choose(gsIdioma, "&Mayores", "&Ledgers")
  frmMain.mnuRCja.Caption = Choose(gsIdioma, "&Caja Bancos", "&Cash and Banks")
  frmMain.mnuRCCo.Caption = Choose(gsIdioma, "&Centro de Costos", "&Cost Center")
  frmMain.mnuRSdo.Caption = Choose(gsIdioma, "&Saldos", "&Balances")
  frmMain.mnuRReg.Caption = Choose(gsIdioma, "Reg&istros", "Reg&isters")
  frmMain.mnuRCtaCte.Caption = Choose(gsIdioma, "C&uentas Corrientes", "C&urrent Accounts")
  frmMain.mnuRTipo54.Caption = Choose(gsIdioma, "&Reportes Tipo 54", "&Reports Type 54")
  frmMain.mnuRTipo56.Caption = Choose(gsIdioma, "&Reportes Tipo 56", "&Reports Type 56")
  frmMain.opcPTrfPDT.Caption = Choose(gsIdioma, "&Transferencia al PDT", "&Transfer to PDT")
  frmMain.mnuUtilCfg.Caption = Choose(gsIdioma, "C&onfiguración", "&Configuration")
  frmMain.mnuUtilSeg.Caption = Choose(gsIdioma, "&Seguridad", "&Security")
  frmMain.opcSalir.Caption = Choose(gsIdioma, "S&alir", "&Exit")
  'Carga todas la opciones del sistema en arreglos
  While Not porstOpciones.EOF
    For n_Index = 0 To (frmMain.Controls.Count - 1)
      If frmMain.Controls(n_Index).Name = porstOpciones!NomMdl Then
        frmMain.Controls(n_Index).Caption = Mid(porstOpciones!descripcion, 6)
      End If
    Next n_Index
    porstOpciones.MoveNext
  Wend
  ' Cierro el recordset y saco del entorno
  porstOpciones.Close: Set porstOpciones = Nothing
  ']
  frmMain.Show

End Sub
Private Sub Mensajes()
  NvlUsr_Adm_Txt = Choose(gsIdioma, "Administrador ", "Administrator ")
  NvlUsr_Sup_Txt = Choose(gsIdioma, "Supervisor ", "Supervisor ")
  NvlUsr_Asi_Txt = Choose(gsIdioma, "Asistente  ", "Assistant  ")

  ESTMDL_ACT_TXT = Choose(gsIdioma, "Activo  ", "Active  ")
  ESTMDL_INA_TXT = Choose(gsIdioma, "Inactivo", "Inactive")
  ESTUSR_ACT_TXT = Choose(gsIdioma, "Activo  ", "Active  ")
  ESTUSR_INA_TXT = Choose(gsIdioma, "Inactivo", "Inactive")

  ESTAUX_ACT_TXT = Choose(gsIdioma, "Activo  ", "Active  ")
  ESTAUX_INA_TXT = Choose(gsIdioma, "Inactivo", "Inactive")
  ESTCCO_ACT_TXT = Choose(gsIdioma, "Activo  ", "Active  ")
  ESTCCO_INA_TXT = Choose(gsIdioma, "Inactivo", "Inactive")
  ESTCTA_ACT_TXT = Choose(gsIdioma, "Activa  ", "Active  ")
  ESTCTA_INA_TXT = Choose(gsIdioma, "Inactiva", "Inactive")

  INDSDO_POS_TXT = Choose(gsIdioma, "Positivo", "Positive")
  INDSDO_NEG_TXT = Choose(gsIdioma, "Negativo", "Negative")
  
  NATCTA_DEU_TXT = Choose(gsIdioma, "Deudora  ", "Debtor   ")
  NATCTA_ACR_TXT = Choose(gsIdioma, "Acreedora", "Creditor ")
  SGNTDC_POS_TXT = Choose(gsIdioma, "Positivo", "Positive")
  SGNTDC_NEG_TXT = Choose(gsIdioma, "Negativo", "Negative")

  TPOANL_CTA_TXT = Choose(gsIdioma, "Cuenta Contable", "Accountable Account")
  TPOANL_AUX_TXT = Choose(gsIdioma, "Cuenta Auxiliar", "Auxiliary Account")
  TPOANL_DOC_TXT = Choose(gsIdioma, "Documento", "Document")
  TPOCTA_TIT_TXT = Choose(gsIdioma, "Título ", "Title  ")
  TPOCTA_TRA_TXT = Choose(gsIdioma, "Detalle", "Detail ")

  TPOGNR_DRO_TXT = Choose(gsIdioma, "Diario", "Journal")
  TPOGNR_CPR_TXT = Choose(gsIdioma, "Compra", "Purchase")
  TPOGNR_VTA_TXT = Choose(gsIdioma, "Venta", "Sale")
  TPOGNR_HPR_TXT = Choose(gsIdioma, "Honor.", "Fees")
  TPOGNR_DST_TXT = Choose(gsIdioma, "Destino", "Destination")
  TPOGNR_DCA_TXT = Choose(gsIdioma, "D.Camb.", "D.Exchan.")
  TPOGNR_APE_TXT = Choose(gsIdioma, "Apert.", "Opening")
  TPOGNR_CIE_TXT = Choose(gsIdioma, "Cierre", "Closing")
  TPOGNR_DRP_TXT = Choose(gsIdioma, "Rtn/Pcn", "Whn/Pcn")
  TPOGNR_BAN_TXT = Choose(gsIdioma, "CajaBanco", "CashBank")
  TPOHT1_SAL_TXT = Choose(gsIdioma, "Saldo", "Rest")
  TPOHT1_DEP_TXT = Choose(gsIdioma, "Depreciación", "Depreciation")

  TPOBAN_ING_TXT = Choose(gsIdioma, "Ingreso", "Income")
  TPOBAN_EGR_TXT = Choose(gsIdioma, "Egreso", "Expense")

  TPODOC_DPS_TXT = Choose(gsIdioma, "Dep. cuenta", "Depos. account")
  TPODOC_GRO_TXT = Choose(gsIdioma, "Giro bancario", "Bank draft")
  TPODOC_TRA_TXT = Choose(gsIdioma, "Transf. fondos", "Transf. bottoms")
  TPODOC_ORD_TXT = Choose(gsIdioma, "Orden pago", "Payment order")
  TPODOC_DEB_TXT = Choose(gsIdioma, "Tarjeta Debito", "Card debit")
  TPODOC_CRE_TXT = Choose(gsIdioma, "Tarjeta Credito", "Card credit")
  TPODOC_CHQ_TXT = Choose(gsIdioma, "Cheque", "Check")
  TPODOC_OTR_TXT = Choose(gsIdioma, "Otro", "Other")
  TPODOC_EFE_TXT = Choose(gsIdioma, "Efectivo", "Cash")
  TPODOC_PEX_TXT = Choose(gsIdioma, "Medio com. exterior", "Average foreign trade")
  TPODOC_LTR_TXT = Choose(gsIdioma, "Letra cambio", "Change letter")
  TPODOC_CGE_TXT = Choose(gsIdioma, "Cheque Gerencia", "Check Management")
  
  TPOLIN_CTA_TXT = Choose(gsIdioma, "Cuentas", "Accounts")
  TPOLIN_TIT_TXT = Choose(gsIdioma, "Título", "Title")
  TPOLIN_STO_TXT = Choose(gsIdioma, "Subtotal", "Subtotal")
  TPOLIN_TOT_TXT = Choose(gsIdioma, "Total", "Total")
  TPOLIN_OPE_TXT = Choose(gsIdioma, "Operación", "Operation")
  TPOLIN_MAS_TXT = Choose(gsIdioma, "Mascara", "Mask")
  TPOMON_NAC_TXT_0 = Choose(gsIdioma, "MN", "NC")
  TPOMON_EXT_TXT_0 = Choose(gsIdioma, "ME", "FC")
  TPOMON_NAC_TXT_1 = Choose(gsIdioma, "Mon.Nac.", "Nat.Cur.")
  TPOMON_EXT_TXT_1 = Choose(gsIdioma, "Mon.Ext.", "For.Cur.")
  TPOMON_NAC_TXT_2 = Choose(gsIdioma, "Moneda Nacional", "National Currency")
  TPOMON_EXT_TXT_2 = Choose(gsIdioma, "Moneda Extranjera", "Foreign Currency")
  
  TPOCTA_COR_TXT_2 = Choose(gsIdioma, "Cuenta Corriente", "Current Account")
  TPOCTA_AHO_TXT_2 = Choose(gsIdioma, "Cuenta de Ahorros", "Savings Account")
  TPOCTA_MAE_TXT_2 = Choose(gsIdioma, "Cuenta Maestra", "Master Account")
  
  TPOSDO_INV_TXT = Choose(gsIdioma, "Inventario", "Inventory")
  TPOSDO_RES_TXT = Choose(gsIdioma, "Resultados", "Results")
  TPOSDO_FUN_TXT = Choose(gsIdioma, "Función", "Function")
  TPOSDO_NAT_TXT = Choose(gsIdioma, "Naturaleza", "Nature")
  TPOSDO_AMB_TXT = Choose(gsIdioma, "Función y Naturaleza", "Function and Nature")
  TPOTCB_CPR_TXT = Choose(gsIdioma, "Compra", "Purchase")
  TPOTCB_VTA_TXT = Choose(gsIdioma, "Venta", "Sale")

  TPOGRU1_TXT_1 = Choose(gsIdioma, "INGRESOS", "INCOME")
  TPOGRU2_TXT_1 = Choose(gsIdioma, "GASTOS", "EXPENSES")
  TPOGRU3_TXT_1 = Choose(gsIdioma, "TOTAL", "TOTAL")
  TPOGRU4_TXT_1 = Choose(gsIdioma, "OTROS", "OTHERS")

  TPOFJO_ING_TXT = Choose(gsIdioma, "Ingreso ", "Income  ")
  TPOFJO_EGR_TXT = Choose(gsIdioma, "Egreso  ", "Disbursements")

  TPOEFE_OPE_TXT = Choose(gsIdioma, "Operación      ", "Operation      ")
  TPOEFE_INV_TXT = Choose(gsIdioma, "Inversión      ", "Investment     ")
  TPOEFE_FIN_TXT = Choose(gsIdioma, "Financiamiento ", "Financing      ")
  
  TEXT_Ninguno = Choose(gsIdioma, "Ninguno", "None")
  TEXT_ResponsableInscrito = Choose(gsIdioma, "[RI] Responsable Inscrito", "Enrolled Person in charge")
  TEXT_ResponsableMonotributo = Choose(gsIdioma, "[RC] Responsable Monotributo", "Monotributo Person in charge")
  TEXT_Exepto = Choose(gsIdioma, "[Ex] Exepto", "Exepto")
  TEXT_NoAlcanzado = Choose(gsIdioma, "[NA] No Alcanzado", "Not Reached")
  TEXT_ConsumidosFinal = Choose(gsIdioma, "[CF] Consumidos Final", "Consumed Final")
  
  TEXT_ImpuestoDetallado = Choose(gsIdioma, "Factura A Detallado", "Detailed Tax")
  TEXT_FacturaPublica = Choose(gsIdioma, "Factura B (no muestra impuesto)", "Invoice (it does not show tax)")
  TEXT_FacturaContador = Choose(gsIdioma, "Factura C Pública (no muestra impuesto) contadores y abogados", "Public invoice (it does not show tax) accountants and lawyers")
  TEXT_RetencionIva = Choose(gsIdioma, "Ret Iva", "Ret Iva")
  TEXT_RetencionIB = Choose(gsIdioma, "Ret IB", "Ret IB")
  TEXT_RetencionIG = Choose(gsIdioma, "Ret IG", "Ret IG")
  TEXT_RetencionSuss = Choose(gsIdioma, "Ret Suss", "Ret Suss")
  TEXT_RetencionOtro = Choose(gsIdioma, "Ret Otros", "Ret Otros")
  
  TEXT_NUEVO = Choose(gsIdioma, "Nuevo", "New")
  TEXT_MODIF = Choose(gsIdioma, "Corregir los datos de", "Correct the data of")
  TEXT_BUSCA = Choose(gsIdioma, "&Buscar por ", "Search for ")
  TEXT_1021 = Choose(gsIdioma, "¿Realmente desea eliminar el registro", "Do you really want to eliminate the register")
  TEXT_1022 = Choose(gsIdioma, "¿Desea revisar este documento?", "Do you want to review this document?")
  TEXT_1031 = Choose(gsIdioma, "Listado de", "Listing of")
  TEXT_3001 = Choose(gsIdioma, "Registro(s)", "Register(s)")
  TEXT_3101 = Choose(gsIdioma, "<Vacío>", "<Empty>")
  TEXT_6001 = Choose(gsIdioma, "Error:", "Error:")
  TEXT_6002 = Choose(gsIdioma, "Faltan datos.", "Lack information")
  TEXT_8001 = Choose(gsIdioma, "No hay datos.", "There are not data")
  TEXT_8002 = Choose(gsIdioma, "No existe el archivo de configuración.", "The configuration file does not exist.")
  TEXT_8003 = Choose(gsIdioma, "Corriga el archivo de configuración.", "Correct the configuration file.")
  TEXT_8004 = Choose(gsIdioma, "No existe la base de datos.", "The data base does not exist.")
  TEXT_8005 = Choose(gsIdioma, "Este dato no debe quedar en blanco.", "This data must not be empty.")
  TEXT_8006 = Choose(gsIdioma, "El dato no existe.", "The data does not exist.")
  TEXT_8007 = Choose(gsIdioma, "La llave ya existe.", "The primary key already exists.")
  TEXT_8008 = Choose(gsIdioma, "El proceso ha terminado.", "The process has finished.")
  TEXT_8009 = Choose(gsIdioma, "El documento está anulado.", "The document is annulled.")
  TEXT_8010 = Choose(gsIdioma, "El dato no es valido", "The data is not been worth.")

  TEXT_9011 = Choose(gsIdioma, "El importe total (", "The total amount (") & TPOMON_NAC_TXT_1 & Choose(gsIdioma, ") del documento no cuadra con los parciales.", ") of document does not tally with the parcial ones.")
  TEXT_9012 = Choose(gsIdioma, "El importe total (", "The total amount (") & TPOMON_EXT_TXT_1 & Choose(gsIdioma, ") del documento no cuadra con los parciales.", ") of document does not tally with the parcial ones.")
  TEXT_9013 = Choose(gsIdioma, "No han sido registradas todas las cuentas para los importes ingresados o estos no cuadran.", "All accounts for the entered amounts have not been registered or they do not tally with")
  TEXT_9015 = Choose(gsIdioma, "El Tipo de Cambio no ha sido registrado para la fecha indicada.", "The rate of exchange has not been registered for the indicated date")
  TEXT_9016 = Choose(gsIdioma, "El mes está cerrado para este tipo de transacción.", "The month is closed for this type of transaction")
  
    'ini 2014-07-10 validacion T.Doc=05
    TEXT_9017 = Choose(gsIdioma, _
    "Error, en la SERIE debe poner los siguientes digitos: 1=Boleto Manual, 2=Boleto Automatico, 3=Boleto Electronico, 4=Otros." _
    , "Error in SERIES digits must add the following: 1 = Manual Ticket, Ticket 2 = Auto, 3 = Electronic Ticket, 4 = Other.")
    'ini 2014-07-10 validacion T.Doc=05
    
'ini 2015-05-18 validacion frm
    TEXT_9018 = Choose(gsIdioma, "No aplica diferencia de cambio", "Not applicable exchange difference")
    TEXT_9019 = Choose(gsIdioma, "Registrar Tipo de Cambio de Cierre.", "Register Exchange Closing")
'fin 2015-05-18 validacion frm
   
'ini 2015-06-06 Si Mayorizo o no . Estado Mayorizacion
  TEXT_9020 = Choose(gsIdioma, "Vuelva a Mayorizar. Existen Transacciones modificadas", "Vuelva a Mayorizar. Existen Transacciones modificadas")
'fin 2015-06-06 Si Mayorizo o no . Estado Mayorizacion
    
''ini 2015-06-30 correccion tipo mon cta
'    TEXT_9021 = Choose(gsIdioma, _
'    "Error, tipo de cambio diferente entre cuenta y transaccion." _
'    , "Error, different exchange rate between account and transaction.")
''fin 2015-06-30 correccion tipo mon cta
'ini 2015-08-19 correccion tipo mon cta
    TEXT_9021 = Choose(gsIdioma, _
    "Error, moneda diferente entre cuenta y documento" _
    , "Error,different currency between account and document.")
'fin 2015-08-19 correccion tipo mon cta

'ini '2015-07-17 error eliminar detraccion
    TEXT_9022 = Choose(gsIdioma, _
    "No puede eliminar, existen trasacciones relacionadas." _
    , "You can not delete, there are related trasacciones.")
'fin '2015-07-17 error eliminar detraccion
'ini 2015-07-21 t.cambio sunat
    TEXT_9023 = Choose(gsIdioma, _
    "No existen informacion de SUNAT." _
    , "No  information de SUNAT.")
'fin 2015-07-21 t.cambio sunat
'ini 2015-08-27 ctr obligac sunat
    TEXT_9024 = Choose(gsIdioma, _
    "Su impuesto vence: " _
    , "Your tax expires")
    TEXT_9025 = Choose(gsIdioma, _
    " ¿ya lo presento? " _
    , " Already I present? ")
'fin 2015-08-27 ctr obligac sunat

'ini 2016-05-27/28 nivel=asisten no elimin datos
    TEXT_9026 = Choose(gsIdioma, _
    "No puede Eliminar. Su nivel de acceso no lo permite" _
    , "You can not delete . Your access level does not allow")
'fin 2016-05-27/28 nivel=asisten no elimin datos
  
    
'ini 2014-08-05 RR.HH afecto afp/onp
  TPOCOMI_MIXTA_TXT = Choose(gsIdioma, "Mixta", "Joint")
  TPOCOMI_FLUJO_TXT = Choose(gsIdioma, "Flujo", "Flow ")
'fin 2014-08-05 RR.HH afecto afp/onp
  
  
  'No puede eliminar, existen trasacciones relacionadas
  'You can not delete, there are related trasacciones
End Sub

Private Sub MatrizDetraccion()
'ini 2015-07-02 adic tabla detrac

''''n = n + 1: aDtraccCod(n) = "1"
'''Dim n As Integer
''''n = 0
''''1
'''n = 1 + 0: aDtraccDet(n) = "00101-Azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00102-Azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00103-Azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00104-Azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00105-Azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00201-Arroz Pilado 3.85%"
'''n = 1 + n: aDtraccDet(n) = "00202-Arroz Pilado 3.85%"
'''n = 1 + n: aDtraccDet(n) = "00203-Arroz Pilado 3.85%"
'''n = 1 + n: aDtraccDet(n) = "00204-Arroz Pilado 3.85%"
'''n = 1 + n: aDtraccDet(n) = "00205-Arroz Pilado 3.85%"
''''10
'''n = 1 + n: aDtraccDet(n) = "00301-Alcohol etílico 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00302-Alcohol etílico 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00303-Alcohol etílico 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00304-Alcohol etílico 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00305-Alcohol etílico 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00401-Recursos hidrobiológicos 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00402-Recursos hidrobiológicos 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00403-Recursos hidrobiológicos 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00404-Recursos hidrobiológicos 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00405-Recursos hidrobiológicos 4%" '9%
''''20
'''n = 1 + n: aDtraccDet(n) = "00501-Maíz amarillo duro 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00502-Maíz amarillo duro 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00503-Maíz amarillo duro 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00504-Maíz amarillo duro 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00505-Maíz amarillo duro 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00601-Algodón 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00602-Algodón 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00603-Algodón 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00604-Algodón 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00605-Algodón 0%" '9%
''''30
'''n = 1 + n: aDtraccDet(n) = "00701-Caña de azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00702-Caña de azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00703-Caña de azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00704-Caña de azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00705-Caña de azúcar 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "00801-Madera 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00802-Madera 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00803-Madera 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00804-Madera 4%" '9%
'''n = 1 + n: aDtraccDet(n) = "00805-Madera 4%" '9%
''''40
'''n = 1 + n: aDtraccDet(n) = "00901-Arena y piedra. 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "00902-Arena y piedra. 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "00903-Arena y piedra. 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "00904-Arena y piedra. 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "00905-Arena y piedra. 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01001-Residuos, subprod,desech, recor y desperdicio 15%"
'''n = 1 + n: aDtraccDet(n) = "01002-Residuos, subprod,desech, recor y desperdicio 15%"
'''n = 1 + n: aDtraccDet(n) = "01003-Residuos, subprod,desech, recor y desperdicio 15%"
'''n = 1 + n: aDtraccDet(n) = "01004-Residuos, subprod,desech, recor y desperdicio 15%"
'''n = 1 + n: aDtraccDet(n) = "01005-Residuos, subprod,desech, recor y desperdicio 15%"
''''50
'''n = 1 + n: aDtraccDet(n) = "01101-Bienes grava. con el IGV, x renunci.exone (2) 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01102-Bienes grava. con el IGV, x renunci.exone (2) 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01103-Bienes grava. con el IGV, x renunci.exone (2) 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01104-Bienes grava. con el IGV, x renunci.exone (2) 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01105-Bienes grava. con el IGV, x renunci.exone (2) 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01201-Intermediacion laboral y tercerización 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01202-Intermediacion laboral y tercerización 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01203-Intermediacion laboral y tercerización 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01204-Intermediacion laboral y tercerización 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01205-Intermediacion laboral y tercerización 10%" '12%
''''60
''''2015-01-13 no existe codigo 013 en la ultima lista monto2
'''n = 1 + n: aDtraccDet(n) = "01301-Animales vivos 10%"
'''n = 1 + n: aDtraccDet(n) = "01302-Animales vivos 10%"
'''n = 1 + n: aDtraccDet(n) = "01303-Animales vivos 10%"
'''n = 1 + n: aDtraccDet(n) = "01304-Animales vivos 10%"
'''n = 1 + n: aDtraccDet(n) = "01305-Animales vivos 10%"
'''n = 1 + n: aDtraccDet(n) = "01401-Carnes y despojos comestibles 4%"
'''n = 1 + n: aDtraccDet(n) = "01402-Carnes y despojos comestibles 4%"
'''n = 1 + n: aDtraccDet(n) = "01403-Carnes y despojos comestibles 4%"
'''n = 1 + n: aDtraccDet(n) = "01404-Carnes y despojos comestibles 4%"
'''n = 1 + n: aDtraccDet(n) = "01405-Carnes y despojos comestibles 4%"
''''70
''''2015-01-13 no existe codigo 015 en la ultima lista monto2
'''n = 1 + n: aDtraccDet(n) = "01501-Abonos, cueros y pieles de origen animal 10%"
'''n = 1 + n: aDtraccDet(n) = "01502-Abonos, cueros y pieles de origen animal 10%"
'''n = 1 + n: aDtraccDet(n) = "01503-Abonos, cueros y pieles de origen animal 10%"
'''n = 1 + n: aDtraccDet(n) = "01504-Abonos, cueros y pieles de origen animal 10%"
'''n = 1 + n: aDtraccDet(n) = "01505-Abonos, cueros y pieles de origen animal 10%"
'''n = 1 + n: aDtraccDet(n) = "01601-Aceite de pescado 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01602-Aceite de pescado 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01603-Aceite de pescado 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01604-Aceite de pescado 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01605-Aceite de pescado 0%" '9%
''''80
'''n = 1 + n: aDtraccDet(n) = "01701-Harina, polvo y pellets de pesca, crustáce.,  4%" '9%
'''n = 1 + n: aDtraccDet(n) = "01702-Harina, polvo y pellets de pesca, crustáce.,  4%" '9%
'''n = 1 + n: aDtraccDet(n) = "01703-Harina, polvo y pellets de pesca, crustáce.,  4%" '9%
'''n = 1 + n: aDtraccDet(n) = "01704-Harina, polvo y pellets de pesca, crustáce.,  4%" '9%
'''n = 1 + n: aDtraccDet(n) = "01705-Harina, polvo y pellets de pesca, crustáce.,  4%" '9%
'''n = 1 + n: aDtraccDet(n) = "01801-Embarcaciones pesqueras 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01802-Embarcaciones pesqueras 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01803-Embarcaciones pesqueras 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01804-Embarcaciones pesqueras 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "01805-Embarcaciones pesqueras 0%" '9%
''''90
'''n = 1 + n: aDtraccDet(n) = "01901-Arrendamiento de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01902-Arrendamiento de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01903-Arrendamiento de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01904-Arrendamiento de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "01905-Arrendamiento de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02001-Mantenimiento y reparación de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02002-Mantenimiento y reparación de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02003-Mantenimiento y reparación de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02004-Mantenimiento y reparación de bienes muebles 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02005-Mantenimiento y reparación de bienes muebles 10%" '12%
''''100
'''n = 1 + n: aDtraccDet(n) = "02101-Movimiento de carga 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02102-Movimiento de carga 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02103-Movimiento de carga 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02104-Movimiento de carga 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02105-Movimiento de carga 10%" '12%
''''ini 2014-07-07 cambio segun req sandra.
''''n = 1 + n: aDtraccDet(n) = "02201-Otros servicios empresariales 12%"
''''n = 1 + n: aDtraccDet(n) = "02202-Otros servicios empresariales 12%"
''''n = 1 + n: aDtraccDet(n) = "02203-Otros servicios empresariales 12%"
''''n = 1 + n: aDtraccDet(n) = "02204-Otros servicios empresariales 12%"
''''n = 1 + n: aDtraccDet(n) = "02205-Otros servicios empresariales 12%"
'''n = 1 + n: aDtraccDet(n) = "02201-Otros servicios empresariales 10%"
'''n = 1 + n: aDtraccDet(n) = "02202-Otros servicios empresariales 10%"
'''n = 1 + n: aDtraccDet(n) = "02203-Otros servicios empresariales 10%"
'''n = 1 + n: aDtraccDet(n) = "02204-Otros servicios empresariales 10%"
'''n = 1 + n: aDtraccDet(n) = "02205-Otros servicios empresariales 10%"
''''110
''''fin 2014-07-07 cambio segun req sandra.
'''n = 1 + n: aDtraccDet(n) = "02301-Leche 0%" '4%
'''n = 1 + n: aDtraccDet(n) = "02302-Leche 0%" '4%
'''n = 1 + n: aDtraccDet(n) = "02303-Leche 0%" '4%
'''n = 1 + n: aDtraccDet(n) = "02304-Leche 0%" '4%
'''n = 1 + n: aDtraccDet(n) = "02305-Leche 0%" '4%
'''n = 1 + n: aDtraccDet(n) = "02401-Comisión mercantil 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02402-Comisión mercantil 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02403-Comisión mercantil 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02404-Comisión mercantil 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02405-Comisión mercantil 10%" '12%
''''120
'''n = 1 + n: aDtraccDet(n) = "02501-Fabricación de bienes por encargo 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02502-Fabricación de bienes por encargo 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02503-Fabricación de bienes por encargo 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02504-Fabricación de bienes por encargo 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02505-Fabricación de bienes por encargo 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02601-Servicio de transporte de personas 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02602-Servicio de transporte de personas 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02603-Servicio de transporte de personas 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02604-Servicio de transporte de personas 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "02605-Servicio de transporte de personas 10%" '12%
''''130
'''n = 1 + n: aDtraccDet(n) = "02701-Servic. transpo. bienes realiz. x vía terrest 4%"
'''n = 1 + n: aDtraccDet(n) = "02702-Servic. transpo. bienes realiz. x vía terrest 4%"
'''n = 1 + n: aDtraccDet(n) = "02703-Servic. transpo. bienes realiz. x vía terrest 4%"
'''n = 1 + n: aDtraccDet(n) = "02704-Servic. transpo. bienes realiz. x vía terrest 4%"
'''n = 1 + n: aDtraccDet(n) = "02705-Servic. transpo. bienes realiz. x vía terrest 4%"
'''n = 1 + n: aDtraccDet(n) = "02801-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'''n = 1 + n: aDtraccDet(n) = "02802-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'''n = 1 + n: aDtraccDet(n) = "02803-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'''n = 1 + n: aDtraccDet(n) = "02804-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'''n = 1 + n: aDtraccDet(n) = "02805-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
''''140
'''n = 1 + n: aDtraccDet(n) = "02901-Algodón en rama sin desmotar (artículo 3° de  0%" '9%
'''n = 1 + n: aDtraccDet(n) = "02902-Algodón en rama sin desmotar (artículo 3° de  0%" '9%
'''n = 1 + n: aDtraccDet(n) = "02903-Algodón en rama sin desmotar (artículo 3° de  0%" '9%
'''n = 1 + n: aDtraccDet(n) = "02904-Algodón en rama sin desmotar (artículo 3° de  0%" '9%
'''n = 1 + n: aDtraccDet(n) = "02905-Algodón en rama sin desmotar (artículo 3° de  0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03001-Contratos de construcción 4%"
'''n = 1 + n: aDtraccDet(n) = "03002-Contratos de construcción 4%"
'''n = 1 + n: aDtraccDet(n) = "03003-Contratos de construcción 4%"
'''n = 1 + n: aDtraccDet(n) = "03004-Contratos de construcción 4%"
'''n = 1 + n: aDtraccDet(n) = "03005-Contratos de construcción 4%"
''''150
'''n = 1 + n: aDtraccDet(n) = "03101-Oro gravado con el IGV (2) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03102-Oro gravado con el IGV (2) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03103-Oro gravado con el IGV (2) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03104-Oro gravado con el IGV (2) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03105-Oro gravado con el IGV (2) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03201-Páprika y otros fruto. Género.capsicum o pimi 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03202-Páprika y otros fruto. Género.capsicum o pimi 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03203-Páprika y otros fruto. Género.capsicum o pimi 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03204-Páprika y otros fruto. Género.capsicum o pimi 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03205-Páprika y otros fruto. Género.capsicum o pimi 0%" '9%
''''160
'''n = 1 + n: aDtraccDet(n) = "03301-Espárragos 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03302-Espárragos 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03303-Espárragos 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03304-Espárragos 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03305-Espárragos 0%" '9%
'''n = 1 + n: aDtraccDet(n) = "03401-Minerales metálicos no auriferos 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03402-Minerales metálicos no auriferos 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03403-Minerales metálicos no auriferos 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03404-Minerales metálicos no auriferos 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03405-Minerales metálicos no auriferos 10%" '12%
''''170
'''n = 1 + n: aDtraccDet(n) = "03501-Bienes exonerados del IGV (3) 1.5%"
'''n = 1 + n: aDtraccDet(n) = "03502-Bienes exonerados del IGV (3) 1.5%"
'''n = 1 + n: aDtraccDet(n) = "03503-Bienes exonerados del IGV (3) 1.5%"
'''n = 1 + n: aDtraccDet(n) = "03504-Bienes exonerados del IGV (3) 1.5%"
'''n = 1 + n: aDtraccDet(n) = "03505-Bienes exonerados del IGV (3) 1.5%"
'''n = 1 + n: aDtraccDet(n) = "03601-Oro y demás minerales metálicos exonerados de 1.5%" '4%
'''n = 1 + n: aDtraccDet(n) = "03602-Oro y demás minerales metálicos exonerados de 1.5%" '4%
'''n = 1 + n: aDtraccDet(n) = "03603-Oro y demás minerales metálicos exonerados de 1.5%" '4%
'''n = 1 + n: aDtraccDet(n) = "03604-Oro y demás minerales metálicos exonerados de 1.5%" '4%
'''n = 1 + n: aDtraccDet(n) = "03605-Oro y demás minerales metálicos exonerados de 1.5%" '4%
''''180
''''ini 2014-07-07 cambio segun req sandra.
''''n = 1 + n: aDtraccDet(n) = "03701-Demás servicios gravados con el IGV  12%"
''''n = 1 + n: aDtraccDet(n) = "03702-Demás servicios gravados con el IGV  12%"
''''n = 1 + n: aDtraccDet(n) = "03703-Demás servicios gravados con el IGV  12%"
''''n = 1 + n: aDtraccDet(n) = "03704-Demás servicios gravados con el IGV  12%"
''''n = 1 + n: aDtraccDet(n) = "03705-Demás servicios gravados con el IGV  12%"
'''n = 1 + n: aDtraccDet(n) = "03701-Demás servicios gravados con el IGV  10%"
'''n = 1 + n: aDtraccDet(n) = "03702-Demás servicios gravados con el IGV  10%"
'''n = 1 + n: aDtraccDet(n) = "03703-Demás servicios gravados con el IGV  10%"
'''n = 1 + n: aDtraccDet(n) = "03704-Demás servicios gravados con el IGV  10%"
'''n = 1 + n: aDtraccDet(n) = "03705-Demás servicios gravados con el IGV  10%"
''''fin 2014-07-07 cambio segun req sandra.
'''
'''n = 1 + n: aDtraccDet(n) = "03801-Espectáculos públicos no culturales (4) 0%" '7%
'''n = 1 + n: aDtraccDet(n) = "03802-Espectáculos públicos no culturales (4) 0%" '7%
'''n = 1 + n: aDtraccDet(n) = "03803-Espectáculos públicos no culturales (4) 0%" '7%
'''n = 1 + n: aDtraccDet(n) = "03804-Espectáculos públicos no culturales (4) 0%" '7%
'''n = 1 + n: aDtraccDet(n) = "03805-Espectáculos públicos no culturales (4) 0%" '7%
''''190
'''n = 1 + n: aDtraccDet(n) = "03901-Minerales no metálicos (3) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03902-Minerales no metálicos (3) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03903-Minerales no metálicos (3) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03904-Minerales no metálicos (3) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "03905-Minerales no metálicos (3) 10%" '12%
'''n = 1 + n: aDtraccDet(n) = "04001-Bien inmueble gravado con el IGV (5) 4%"
'''n = 1 + n: aDtraccDet(n) = "04002-Bien inmueble gravado con el IGV (5) 4%"
'''n = 1 + n: aDtraccDet(n) = "04003-Bien inmueble gravado con el IGV (5) 4%"
'''n = 1 + n: aDtraccDet(n) = "04004-Bien inmueble gravado con el IGV (5) 4%"
'''n = 1 + n: aDtraccDet(n) = "04005-Bien inmueble gravado con el IGV (5) 4%"
''''200
'''n = 1 + n: aDtraccDet(n) = "04101-Plomo (6) 0%" '15%
'''n = 1 + n: aDtraccDet(n) = "04102-Plomo (6) 0%" '15%
'''n = 1 + n: aDtraccDet(n) = "04103-Plomo (6) 0%" '15%
'''n = 1 + n: aDtraccDet(n) = "04104-Plomo (6) 0%" '15%
'''n = 1 + n: aDtraccDet(n) = "04105-Plomo (6) 0%" '15%
''''******************************************************************************************************
''''1
'''n = 1 + 0: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.0385
'''n = 1 + n: aDtraccPor(n) = 0.0385
'''n = 1 + n: aDtraccPor(n) = 0.0385
'''n = 1 + n: aDtraccPor(n) = 0.0385
'''n = 1 + n: aDtraccPor(n) = 0.0385
''''10
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
''''20
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
''''30
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
''''40
'''n = 1 + n: aDtraccPor(n) = 0.1 ' 0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 ' 0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 ' 0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 ' 0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 ' 0.12
'''n = 1 + n: aDtraccPor(n) = 0.15
'''n = 1 + n: aDtraccPor(n) = 0.15
'''n = 1 + n: aDtraccPor(n) = 0.15
'''n = 1 + n: aDtraccPor(n) = 0.15
'''n = 1 + n: aDtraccPor(n) = 0.15
''''50
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
''''60
''''2015-01-13 no existe codigo 013 en la ultima lista monto2
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
''''70
''''2015-01-13 no existe codigo 015 en la ultima lista monto2
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
''''80
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
''''90
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
''''100
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
''''ini 2014-07-07 cambio segun req sandra.
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
''''fin 2014-07-07 cambio segun req sandra.
''''110
'''n = 1 + n: aDtraccPor(n) = 0 '0.04
'''n = 1 + n: aDtraccPor(n) = 0 '0.04
'''n = 1 + n: aDtraccPor(n) = 0 '0.04
'''n = 1 + n: aDtraccPor(n) = 0 '0.04
'''n = 1 + n: aDtraccPor(n) = 0 '0.04
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
''''120
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
''''130
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0
'''n = 1 + n: aDtraccPor(n) = 0
'''n = 1 + n: aDtraccPor(n) = 0
'''n = 1 + n: aDtraccPor(n) = 0
'''n = 1 + n: aDtraccPor(n) = 0
''''140
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
''''150
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
''''160
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0 '0.09
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
''''170
'''n = 1 + n: aDtraccPor(n) = 0.015
'''n = 1 + n: aDtraccPor(n) = 0.015
'''n = 1 + n: aDtraccPor(n) = 0.015
'''n = 1 + n: aDtraccPor(n) = 0.015
'''n = 1 + n: aDtraccPor(n) = 0.015
'''n = 1 + n: aDtraccPor(n) = 0.015 '0.04
'''n = 1 + n: aDtraccPor(n) = 0.015 '0.04
'''n = 1 + n: aDtraccPor(n) = 0.015 '0.04
'''n = 1 + n: aDtraccPor(n) = 0.015 '0.04
'''n = 1 + n: aDtraccPor(n) = 0.015 '0.04
''''180
''''ini 2014-07-07 cambio segun req sandra.
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
''''n = 1 + n: aDtraccPor(n) = 0.12
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
'''n = 1 + n: aDtraccPor(n) = 0.1
''''fin 2014-07-07 cambio segun req sandra.
'''n = 1 + n: aDtraccPor(n) = 0 '0.07
'''n = 1 + n: aDtraccPor(n) = 0 '0.07
'''n = 1 + n: aDtraccPor(n) = 0 '0.07
'''n = 1 + n: aDtraccPor(n) = 0 '0.07
'''n = 1 + n: aDtraccPor(n) = 0 '0.07
''''190
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.1 '0.12
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
'''n = 1 + n: aDtraccPor(n) = 0.04
''''200
'''n = 1 + n: aDtraccPor(n) = 0 '0.15
'''n = 1 + n: aDtraccPor(n) = 0 '0.15
'''n = 1 + n: aDtraccPor(n) = 0 '0.15
'''n = 1 + n: aDtraccPor(n) = 0 '0.15
'''n = 1 + n: aDtraccPor(n) = 0 '0.15
'''
''''******************************************************************************************************
''''activacion segun teo
'''n = 1 + 0: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''n = 1 + n: aDtraccEst(n) = 1
'''
'''
''''activacion segun angi
''''n = 51 '11
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 56 '12
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 91 '19
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 96 '20
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 101 '21
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 106 '22
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 116 '24
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 121 '25
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 126 '26
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 131 '27
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 146 '30
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 171 '35
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 181 '37
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 186 '38
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
''''n = 196 '40
''''aDtraccEst(n) = 1
''''aDtraccEst(n + 1) = 1
''''aDtraccEst(n + 2) = 1
''''aDtraccEst(n + 3) = 1
''''aDtraccEst(n + 4) = 1
'''
''''2014-05-29
''''Código del Plan de Cuentas utilizado por el deudor tributario
'fin 2015-07-02 adic tabla detrac

Dim n As Integer

n = 0
n = 1 + 0: aCodPlCta(n) = "00"
n = 1 + n: aCodPlCta(n) = "01"
n = 1 + n: aCodPlCta(n) = "02"
n = 1 + n: aCodPlCta(n) = "03"
n = 1 + n: aCodPlCta(n) = "04"
n = 1 + n: aCodPlCta(n) = "05"
n = 1 + n: aCodPlCta(n) = "06"
n = 1 + n: aCodPlCta(n) = "07"
'n = 1 + n: aCodPlCta(n) = "99"
n = 1 + n: aCodPlCta(n) = "08" 'reemplazar por "99" al grabar en campo

n = 0
n = 1 + 0: aDetPlCta(n) = "Elegir opcion ...."
n = 1 + n: aDetPlCta(n) = "01-PLAN CONTABLE GENERAL EMPRESARIAL"
n = 1 + n: aDetPlCta(n) = "02-PLAN CONTABLE GENERAL REVISADO"
n = 1 + n: aDetPlCta(n) = "03-PLAN DE CUENTAS PARA EMPRESAS DEL SISTEMA FINANCIERO, SUPERVISADAS POR SBS"
n = 1 + n: aDetPlCta(n) = "04-PLAN DE CUENTAS PARA ENTIDADES PRESTADORAS DE SALUD, SUPERVISADAS POR SBS"
n = 1 + n: aDetPlCta(n) = "05-PLAN DE CUENTAS PARA EMPRESAS DEL SISTEMA ASEGURADOR, SUPERVISADAS POR SBS"
n = 1 + n: aDetPlCta(n) = "06-PLAN DE CUENTAS DE LAS ADMINISTRADORAS PRIVADAS DE FONDOS DE PENSIONES, SUPERVISADAS POR SBS"
n = 1 + n: aDetPlCta(n) = "07-PLAN CONTABLE GUBERNAMENTAL"
n = 1 + n: aDetPlCta(n) = "99-OTROS"

End Sub

Private Sub MatrizDetraccion_2015_01_13()

''n = n + 1: aDtraccCod(n) = "1"
'Dim n As Integer
''n = 0
''1
'n = 1 + 0: aDtraccDet(n) = "00101-Azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00102-Azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00103-Azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00104-Azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00105-Azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00201-Arroz Pilado 3.85%"
'n = 1 + n: aDtraccDet(n) = "00202-Arroz Pilado 3.85%"
'n = 1 + n: aDtraccDet(n) = "00203-Arroz Pilado 3.85%"
'n = 1 + n: aDtraccDet(n) = "00204-Arroz Pilado 3.85%"
'n = 1 + n: aDtraccDet(n) = "00205-Arroz Pilado 3.85%"
''10
'n = 1 + n: aDtraccDet(n) = "00301-Alcohol etílico 9%"
'n = 1 + n: aDtraccDet(n) = "00302-Alcohol etílico 9%"
'n = 1 + n: aDtraccDet(n) = "00303-Alcohol etílico 9%"
'n = 1 + n: aDtraccDet(n) = "00304-Alcohol etílico 9%"
'n = 1 + n: aDtraccDet(n) = "00305-Alcohol etílico 9%"
'n = 1 + n: aDtraccDet(n) = "00401-Recursos hidrobiológicos 9%"
'n = 1 + n: aDtraccDet(n) = "00402-Recursos hidrobiológicos 9%"
'n = 1 + n: aDtraccDet(n) = "00403-Recursos hidrobiológicos 9%"
'n = 1 + n: aDtraccDet(n) = "00404-Recursos hidrobiológicos 9%"
'n = 1 + n: aDtraccDet(n) = "00405-Recursos hidrobiológicos 9%"
''20
'n = 1 + n: aDtraccDet(n) = "00501-Maíz amarillo duro 9%"
'n = 1 + n: aDtraccDet(n) = "00502-Maíz amarillo duro 9%"
'n = 1 + n: aDtraccDet(n) = "00503-Maíz amarillo duro 9%"
'n = 1 + n: aDtraccDet(n) = "00504-Maíz amarillo duro 9%"
'n = 1 + n: aDtraccDet(n) = "00505-Maíz amarillo duro 9%"
'n = 1 + n: aDtraccDet(n) = "00601-Algodón 9%"
'n = 1 + n: aDtraccDet(n) = "00602-Algodón 9%"
'n = 1 + n: aDtraccDet(n) = "00603-Algodón 9%"
'n = 1 + n: aDtraccDet(n) = "00604-Algodón 9%"
'n = 1 + n: aDtraccDet(n) = "00605-Algodón 9%"
''30
'n = 1 + n: aDtraccDet(n) = "00701-Caña de azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00702-Caña de azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00703-Caña de azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00704-Caña de azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00705-Caña de azúcar 9%"
'n = 1 + n: aDtraccDet(n) = "00801-Madera 9%"
'n = 1 + n: aDtraccDet(n) = "00802-Madera 9%"
'n = 1 + n: aDtraccDet(n) = "00803-Madera 9%"
'n = 1 + n: aDtraccDet(n) = "00804-Madera 9%"
'n = 1 + n: aDtraccDet(n) = "00805-Madera 9%"
''40
'n = 1 + n: aDtraccDet(n) = "00901-Arena y piedra. 12%"
'n = 1 + n: aDtraccDet(n) = "00902-Arena y piedra. 12%"
'n = 1 + n: aDtraccDet(n) = "00903-Arena y piedra. 12%"
'n = 1 + n: aDtraccDet(n) = "00904-Arena y piedra. 12%"
'n = 1 + n: aDtraccDet(n) = "00905-Arena y piedra. 12%"
'n = 1 + n: aDtraccDet(n) = "01001-Residuos, subprod,desech, recor y desperdicio 15%"
'n = 1 + n: aDtraccDet(n) = "01002-Residuos, subprod,desech, recor y desperdicio 15%"
'n = 1 + n: aDtraccDet(n) = "01003-Residuos, subprod,desech, recor y desperdicio 15%"
'n = 1 + n: aDtraccDet(n) = "01004-Residuos, subprod,desech, recor y desperdicio 15%"
'n = 1 + n: aDtraccDet(n) = "01005-Residuos, subprod,desech, recor y desperdicio 15%"
''50
'n = 1 + n: aDtraccDet(n) = "01101-Bienes grava. con el IGV, x renunci.exone (2) 9%"
'n = 1 + n: aDtraccDet(n) = "01102-Bienes grava. con el IGV, x renunci.exone (2) 9%"
'n = 1 + n: aDtraccDet(n) = "01103-Bienes grava. con el IGV, x renunci.exone (2) 9%"
'n = 1 + n: aDtraccDet(n) = "01104-Bienes grava. con el IGV, x renunci.exone (2) 9%"
'n = 1 + n: aDtraccDet(n) = "01105-Bienes grava. con el IGV, x renunci.exone (2) 9%"
'n = 1 + n: aDtraccDet(n) = "01201-Intermediacion laboral y tercerización 12%"
'n = 1 + n: aDtraccDet(n) = "01202-Intermediacion laboral y tercerización 12%"
'n = 1 + n: aDtraccDet(n) = "01203-Intermediacion laboral y tercerización 12%"
'n = 1 + n: aDtraccDet(n) = "01204-Intermediacion laboral y tercerización 12%"
'n = 1 + n: aDtraccDet(n) = "01205-Intermediacion laboral y tercerización 12%"
''60
'n = 1 + n: aDtraccDet(n) = "01301-Animales vivos 10%"
'n = 1 + n: aDtraccDet(n) = "01302-Animales vivos 10%"
'n = 1 + n: aDtraccDet(n) = "01303-Animales vivos 10%"
'n = 1 + n: aDtraccDet(n) = "01304-Animales vivos 10%"
'n = 1 + n: aDtraccDet(n) = "01305-Animales vivos 10%"
'n = 1 + n: aDtraccDet(n) = "01401-Carnes y despojos comestibles 4%"
'n = 1 + n: aDtraccDet(n) = "01402-Carnes y despojos comestibles 4%"
'n = 1 + n: aDtraccDet(n) = "01403-Carnes y despojos comestibles 4%"
'n = 1 + n: aDtraccDet(n) = "01404-Carnes y despojos comestibles 4%"
'n = 1 + n: aDtraccDet(n) = "01405-Carnes y despojos comestibles 4%"
''70
'n = 1 + n: aDtraccDet(n) = "01501-Abonos, cueros y pieles de origen animal 10%"
'n = 1 + n: aDtraccDet(n) = "01502-Abonos, cueros y pieles de origen animal 10%"
'n = 1 + n: aDtraccDet(n) = "01503-Abonos, cueros y pieles de origen animal 10%"
'n = 1 + n: aDtraccDet(n) = "01504-Abonos, cueros y pieles de origen animal 10%"
'n = 1 + n: aDtraccDet(n) = "01505-Abonos, cueros y pieles de origen animal 10%"
'n = 1 + n: aDtraccDet(n) = "01601-Aceite de pescado 9%"
'n = 1 + n: aDtraccDet(n) = "01602-Aceite de pescado 9%"
'n = 1 + n: aDtraccDet(n) = "01603-Aceite de pescado 9%"
'n = 1 + n: aDtraccDet(n) = "01604-Aceite de pescado 9%"
'n = 1 + n: aDtraccDet(n) = "01605-Aceite de pescado 9%"
''80
'n = 1 + n: aDtraccDet(n) = "01701-Harina, polvo y pellets de pesca, crustáce.,  9%"
'n = 1 + n: aDtraccDet(n) = "01702-Harina, polvo y pellets de pesca, crustáce.,  9%"
'n = 1 + n: aDtraccDet(n) = "01703-Harina, polvo y pellets de pesca, crustáce.,  9%"
'n = 1 + n: aDtraccDet(n) = "01704-Harina, polvo y pellets de pesca, crustáce.,  9%"
'n = 1 + n: aDtraccDet(n) = "01705-Harina, polvo y pellets de pesca, crustáce.,  9%"
'n = 1 + n: aDtraccDet(n) = "01801-Embarcaciones pesqueras 9%"
'n = 1 + n: aDtraccDet(n) = "01802-Embarcaciones pesqueras 9%"
'n = 1 + n: aDtraccDet(n) = "01803-Embarcaciones pesqueras 9%"
'n = 1 + n: aDtraccDet(n) = "01804-Embarcaciones pesqueras 9%"
'n = 1 + n: aDtraccDet(n) = "01805-Embarcaciones pesqueras 9%"
''90
'n = 1 + n: aDtraccDet(n) = "01901-Arrendamiento de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "01902-Arrendamiento de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "01903-Arrendamiento de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "01904-Arrendamiento de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "01905-Arrendamiento de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "02001-Mantenimiento y reparación de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "02002-Mantenimiento y reparación de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "02003-Mantenimiento y reparación de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "02004-Mantenimiento y reparación de bienes muebles 12%"
'n = 1 + n: aDtraccDet(n) = "02005-Mantenimiento y reparación de bienes muebles 12%"
''100
'n = 1 + n: aDtraccDet(n) = "02101-Movimiento de carga 12%"
'n = 1 + n: aDtraccDet(n) = "02102-Movimiento de carga 12%"
'n = 1 + n: aDtraccDet(n) = "02103-Movimiento de carga 12%"
'n = 1 + n: aDtraccDet(n) = "02104-Movimiento de carga 12%"
'n = 1 + n: aDtraccDet(n) = "02105-Movimiento de carga 12%"
''ini 2014-07-07 cambio segun req sandra.
''n = 1 + n: aDtraccDet(n) = "02201-Otros servicios empresariales 12%"
''n = 1 + n: aDtraccDet(n) = "02202-Otros servicios empresariales 12%"
''n = 1 + n: aDtraccDet(n) = "02203-Otros servicios empresariales 12%"
''n = 1 + n: aDtraccDet(n) = "02204-Otros servicios empresariales 12%"
''n = 1 + n: aDtraccDet(n) = "02205-Otros servicios empresariales 12%"
'n = 1 + n: aDtraccDet(n) = "02201-Otros servicios empresariales 10%"
'n = 1 + n: aDtraccDet(n) = "02202-Otros servicios empresariales 10%"
'n = 1 + n: aDtraccDet(n) = "02203-Otros servicios empresariales 10%"
'n = 1 + n: aDtraccDet(n) = "02204-Otros servicios empresariales 10%"
'n = 1 + n: aDtraccDet(n) = "02205-Otros servicios empresariales 10%"
''110
''fin 2014-07-07 cambio segun req sandra.
'n = 1 + n: aDtraccDet(n) = "02301-Leche 4%"
'n = 1 + n: aDtraccDet(n) = "02302-Leche 4%"
'n = 1 + n: aDtraccDet(n) = "02303-Leche 4%"
'n = 1 + n: aDtraccDet(n) = "02304-Leche 4%"
'n = 1 + n: aDtraccDet(n) = "02305-Leche 4%"
'n = 1 + n: aDtraccDet(n) = "02401-Comisión mercantil 12%"
'n = 1 + n: aDtraccDet(n) = "02402-Comisión mercantil 12%"
'n = 1 + n: aDtraccDet(n) = "02403-Comisión mercantil 12%"
'n = 1 + n: aDtraccDet(n) = "02404-Comisión mercantil 12%"
'n = 1 + n: aDtraccDet(n) = "02405-Comisión mercantil 12%"
''120
'n = 1 + n: aDtraccDet(n) = "02501-Fabricación de bienes por encargo 12%"
'n = 1 + n: aDtraccDet(n) = "02502-Fabricación de bienes por encargo 12%"
'n = 1 + n: aDtraccDet(n) = "02503-Fabricación de bienes por encargo 12%"
'n = 1 + n: aDtraccDet(n) = "02504-Fabricación de bienes por encargo 12%"
'n = 1 + n: aDtraccDet(n) = "02505-Fabricación de bienes por encargo 12%"
'n = 1 + n: aDtraccDet(n) = "02601-Servicio de transporte de personas 12%"
'n = 1 + n: aDtraccDet(n) = "02602-Servicio de transporte de personas 12%"
'n = 1 + n: aDtraccDet(n) = "02603-Servicio de transporte de personas 12%"
'n = 1 + n: aDtraccDet(n) = "02604-Servicio de transporte de personas 12%"
'n = 1 + n: aDtraccDet(n) = "02605-Servicio de transporte de personas 12%"
''130
'n = 1 + n: aDtraccDet(n) = "02701-Servic. transpo. bienes realiz. x vía terrest 4%"
'n = 1 + n: aDtraccDet(n) = "02702-Servic. transpo. bienes realiz. x vía terrest 4%"
'n = 1 + n: aDtraccDet(n) = "02703-Servic. transpo. bienes realiz. x vía terrest 4%"
'n = 1 + n: aDtraccDet(n) = "02704-Servic. transpo. bienes realiz. x vía terrest 4%"
'n = 1 + n: aDtraccDet(n) = "02705-Servic. transpo. bienes realiz. x vía terrest 4%"
'n = 1 + n: aDtraccDet(n) = "02801-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'n = 1 + n: aDtraccDet(n) = "02802-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'n = 1 + n: aDtraccDet(n) = "02803-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'n = 1 + n: aDtraccDet(n) = "02804-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
'n = 1 + n: aDtraccDet(n) = "02805-Servi.Transp.públic.pasaje.realiza.x vía terr 0%"
''140
'n = 1 + n: aDtraccDet(n) = "02901-Algodón en rama sin desmotar (artículo 3° de  9%"
'n = 1 + n: aDtraccDet(n) = "02902-Algodón en rama sin desmotar (artículo 3° de  9%"
'n = 1 + n: aDtraccDet(n) = "02903-Algodón en rama sin desmotar (artículo 3° de  9%"
'n = 1 + n: aDtraccDet(n) = "02904-Algodón en rama sin desmotar (artículo 3° de  9%"
'n = 1 + n: aDtraccDet(n) = "02905-Algodón en rama sin desmotar (artículo 3° de  9%"
'n = 1 + n: aDtraccDet(n) = "03001-Contratos de construcción 4%"
'n = 1 + n: aDtraccDet(n) = "03002-Contratos de construcción 4%"
'n = 1 + n: aDtraccDet(n) = "03003-Contratos de construcción 4%"
'n = 1 + n: aDtraccDet(n) = "03004-Contratos de construcción 4%"
'n = 1 + n: aDtraccDet(n) = "03005-Contratos de construcción 4%"
''150
'n = 1 + n: aDtraccDet(n) = "03101-Oro gravado con el IGV (2) 12%"
'n = 1 + n: aDtraccDet(n) = "03102-Oro gravado con el IGV (2) 12%"
'n = 1 + n: aDtraccDet(n) = "03103-Oro gravado con el IGV (2) 12%"
'n = 1 + n: aDtraccDet(n) = "03104-Oro gravado con el IGV (2) 12%"
'n = 1 + n: aDtraccDet(n) = "03105-Oro gravado con el IGV (2) 12%"
'n = 1 + n: aDtraccDet(n) = "03201-Páprika y otros fruto. Género.capsicum o pimi 9%"
'n = 1 + n: aDtraccDet(n) = "03202-Páprika y otros fruto. Género.capsicum o pimi 9%"
'n = 1 + n: aDtraccDet(n) = "03203-Páprika y otros fruto. Género.capsicum o pimi 9%"
'n = 1 + n: aDtraccDet(n) = "03204-Páprika y otros fruto. Género.capsicum o pimi 9%"
'n = 1 + n: aDtraccDet(n) = "03205-Páprika y otros fruto. Género.capsicum o pimi 9%"
''160
'n = 1 + n: aDtraccDet(n) = "03301-Espárragos 9%"
'n = 1 + n: aDtraccDet(n) = "03302-Espárragos 9%"
'n = 1 + n: aDtraccDet(n) = "03303-Espárragos 9%"
'n = 1 + n: aDtraccDet(n) = "03304-Espárragos 9%"
'n = 1 + n: aDtraccDet(n) = "03305-Espárragos 9%"
'n = 1 + n: aDtraccDet(n) = "03401-Minerales metálicos no auriferos 12%"
'n = 1 + n: aDtraccDet(n) = "03402-Minerales metálicos no auriferos 12%"
'n = 1 + n: aDtraccDet(n) = "03403-Minerales metálicos no auriferos 12%"
'n = 1 + n: aDtraccDet(n) = "03404-Minerales metálicos no auriferos 12%"
'n = 1 + n: aDtraccDet(n) = "03405-Minerales metálicos no auriferos 12%"
''170
'n = 1 + n: aDtraccDet(n) = "03501-Bienes exonerados del IGV (3) 1.5%"
'n = 1 + n: aDtraccDet(n) = "03502-Bienes exonerados del IGV (3) 1.5%"
'n = 1 + n: aDtraccDet(n) = "03503-Bienes exonerados del IGV (3) 1.5%"
'n = 1 + n: aDtraccDet(n) = "03504-Bienes exonerados del IGV (3) 1.5%"
'n = 1 + n: aDtraccDet(n) = "03505-Bienes exonerados del IGV (3) 1.5%"
'n = 1 + n: aDtraccDet(n) = "03601-Oro y demás minerales metálicos exonerados de 4%"
'n = 1 + n: aDtraccDet(n) = "03602-Oro y demás minerales metálicos exonerados de 4%"
'n = 1 + n: aDtraccDet(n) = "03603-Oro y demás minerales metálicos exonerados de 4%"
'n = 1 + n: aDtraccDet(n) = "03604-Oro y demás minerales metálicos exonerados de 4%"
'n = 1 + n: aDtraccDet(n) = "03605-Oro y demás minerales metálicos exonerados de 4%"
''180
''ini 2014-07-07 cambio segun req sandra.
''n = 1 + n: aDtraccDet(n) = "03701-Demás servicios gravados con el IGV  12%"
''n = 1 + n: aDtraccDet(n) = "03702-Demás servicios gravados con el IGV  12%"
''n = 1 + n: aDtraccDet(n) = "03703-Demás servicios gravados con el IGV  12%"
''n = 1 + n: aDtraccDet(n) = "03704-Demás servicios gravados con el IGV  12%"
''n = 1 + n: aDtraccDet(n) = "03705-Demás servicios gravados con el IGV  12%"
'n = 1 + n: aDtraccDet(n) = "03701-Demás servicios gravados con el IGV  10%"
'n = 1 + n: aDtraccDet(n) = "03702-Demás servicios gravados con el IGV  10%"
'n = 1 + n: aDtraccDet(n) = "03703-Demás servicios gravados con el IGV  10%"
'n = 1 + n: aDtraccDet(n) = "03704-Demás servicios gravados con el IGV  10%"
'n = 1 + n: aDtraccDet(n) = "03705-Demás servicios gravados con el IGV  10%"
''fin 2014-07-07 cambio segun req sandra.
'
'n = 1 + n: aDtraccDet(n) = "03801-Espectáculos públicos no culturales (4) 7%"
'n = 1 + n: aDtraccDet(n) = "03802-Espectáculos públicos no culturales (4) 7%"
'n = 1 + n: aDtraccDet(n) = "03803-Espectáculos públicos no culturales (4) 7%"
''190
'n = 1 + n: aDtraccDet(n) = "03804-Espectáculos públicos no culturales (4) 7%"
'n = 1 + n: aDtraccDet(n) = "03805-Espectáculos públicos no culturales (4) 7%"
'n = 1 + n: aDtraccDet(n) = "03901-Minerales no metálicos (3) 12%"
'n = 1 + n: aDtraccDet(n) = "03902-Minerales no metálicos (3) 12%"
'n = 1 + n: aDtraccDet(n) = "03903-Minerales no metálicos (3) 12%"
'n = 1 + n: aDtraccDet(n) = "03904-Minerales no metálicos (3) 12%"
'n = 1 + n: aDtraccDet(n) = "03905-Minerales no metálicos (3) 12%"
'n = 1 + n: aDtraccDet(n) = "04001-Bien inmueble gravado con el IGV (5) 4%"
'n = 1 + n: aDtraccDet(n) = "04002-Bien inmueble gravado con el IGV (5) 4%"
'n = 1 + n: aDtraccDet(n) = "04003-Bien inmueble gravado con el IGV (5) 4%"
''200
'n = 1 + n: aDtraccDet(n) = "04004-Bien inmueble gravado con el IGV (5) 4%"
'n = 1 + n: aDtraccDet(n) = "04005-Bien inmueble gravado con el IGV (5) 4%"
'n = 1 + n: aDtraccDet(n) = "04101-Plomo (6) 15%"
'n = 1 + n: aDtraccDet(n) = "04102-Plomo (6) 15%"
'n = 1 + n: aDtraccDet(n) = "04103-Plomo (6) 15%"
'n = 1 + n: aDtraccDet(n) = "04104-Plomo (6) 15%"
'n = 1 + n: aDtraccDet(n) = "04105-Plomo (6) 15%"
''******************************************************************************************************
''1
'n = 1 + 0: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.0385
'n = 1 + n: aDtraccPor(n) = 0.0385
'n = 1 + n: aDtraccPor(n) = 0.0385
'n = 1 + n: aDtraccPor(n) = 0.0385
'n = 1 + n: aDtraccPor(n) = 0.0385
''10
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
''20
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
''30
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
''40
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
''50
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
''60
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
''70
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
''80
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
''90
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
''100
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
''ini 2014-07-07 cambio segun req sandra.
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
''fin 2014-07-07 cambio segun req sandra.
''110
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
''120
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
''130
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0
'n = 1 + n: aDtraccPor(n) = 0
'n = 1 + n: aDtraccPor(n) = 0
'n = 1 + n: aDtraccPor(n) = 0
'n = 1 + n: aDtraccPor(n) = 0
''140
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
''150
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
''160
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.09
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
''170
'n = 1 + n: aDtraccPor(n) = 0.015
'n = 1 + n: aDtraccPor(n) = 0.015
'n = 1 + n: aDtraccPor(n) = 0.015
'n = 1 + n: aDtraccPor(n) = 0.015
'n = 1 + n: aDtraccPor(n) = 0.015
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
''180
''ini 2014-07-07 cambio segun req sandra.
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
''n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
'n = 1 + n: aDtraccPor(n) = 0.1
''fin 2014-07-07 cambio segun req sandra.
'n = 1 + n: aDtraccPor(n) = 0.07
'n = 1 + n: aDtraccPor(n) = 0.07
'n = 1 + n: aDtraccPor(n) = 0.07
'n = 1 + n: aDtraccPor(n) = 0.07
'n = 1 + n: aDtraccPor(n) = 0.07
''190
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.12
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
'n = 1 + n: aDtraccPor(n) = 0.04
''200
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'n = 1 + n: aDtraccPor(n) = 0.15
'
''******************************************************************************************************
''activacion segun teo
'n = 1 + 0: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'n = 1 + n: aDtraccEst(n) = 1
'
'
'
''2014-05-29
''Código del Plan de Cuentas utilizado por el deudor tributario
'n = 0
'n = 1 + 0: aCodPlCta(n) = "00"
'n = 1 + n: aCodPlCta(n) = "01"
'n = 1 + n: aCodPlCta(n) = "02"
'n = 1 + n: aCodPlCta(n) = "03"
'n = 1 + n: aCodPlCta(n) = "04"
'n = 1 + n: aCodPlCta(n) = "05"
'n = 1 + n: aCodPlCta(n) = "06"
'n = 1 + n: aCodPlCta(n) = "07"
''n = 1 + n: aCodPlCta(n) = "99"
'n = 1 + n: aCodPlCta(n) = "08" 'reemplazar por "99" al grabar en campo
'
'n = 0
'n = 1 + 0: aDetPlCta(n) = "Elegir opcion ...."
'n = 1 + n: aDetPlCta(n) = "01-PLAN CONTABLE GENERAL EMPRESARIAL"
'n = 1 + n: aDetPlCta(n) = "02-PLAN CONTABLE GENERAL REVISADO"
'n = 1 + n: aDetPlCta(n) = "03-PLAN DE CUENTAS PARA EMPRESAS DEL SISTEMA FINANCIERO, SUPERVISADAS POR SBS"
'n = 1 + n: aDetPlCta(n) = "04-PLAN DE CUENTAS PARA ENTIDADES PRESTADORAS DE SALUD, SUPERVISADAS POR SBS"
'n = 1 + n: aDetPlCta(n) = "05-PLAN DE CUENTAS PARA EMPRESAS DEL SISTEMA ASEGURADOR, SUPERVISADAS POR SBS"
'n = 1 + n: aDetPlCta(n) = "06-PLAN DE CUENTAS DE LAS ADMINISTRADORAS PRIVADAS DE FONDOS DE PENSIONES, SUPERVISADAS POR SBS"
'n = 1 + n: aDetPlCta(n) = "07-PLAN CONTABLE GUBERNAMENTAL"
'n = 1 + n: aDetPlCta(n) = "99-OTROS"

End Sub

