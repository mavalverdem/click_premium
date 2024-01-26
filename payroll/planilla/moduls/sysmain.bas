Attribute VB_Name = "mdlMain"
Option Explicit

' Controles o modulos de clases
'Public gdl_Conexion As Conexionbd.conexion          ' Control o clase de conexión de la base de datos
Public gdl_Conexion As Object                      ' Control o clase de conexión de la base de datos
Public porstRecordset As ADODB.Recordset            ' Recordset de Resultado(WithEvents)

Public ps_BDSystems As String, ps_Licencia As String
Public ps_Provider As String, ps_StrgConnec As String
Public ps_Servidor As String, ps_ServidorCon As String
Public ps_UserId As String, ps_Password As String
Public ps_MysqlExe As String, ps_ConxionSql As String
Public ps_WinSystem As String                       ' Directorio de windows

Public go_tdbBusqueda As TDBGrid                    ' Tabla de búsqueda
Public go_dcaBusqueda As Adodc                      ' Recordset de búsqueda
Public gn_ColBusqueda As Byte                       ' Numero campos de busqueda

Public n_SwConfigura As Byte                        ' Swits de seguridad

Public Const ps_CopyRight As String = "sysma"
Public Const ps_CodSistema As String = "PL"
Public Const ps_NombreDsn As String = "mavm"
Public Const pFileSystem As String = "sysmavm.ini"
'Public Const pFileSystem As String = "plamavm.ini"
Public Const pFileStruc1 As String = "plbdqryt.mac"
Public Const pFileStruc2 As String = "plbdqryf.mac"
Public Const pFileStruc3 As String = "plbdqryd.mai"
Public Const pFileStruc4 As String = "cobdqryt.mac"
Public Const pFileStruc5 As String = "cobdqryd.mai"
Public Const ps_DriverCnn As String = "{MySQL ODBC 3.51 Driver}"

Public formulario As String
Public labelprogreso As String
Public IntervalodeTiempo As Integer
Public TipodeProgreso As Integer

Public queopcion As String
Public quecodpvs As String
Public quecodpsn As String
Public quepdopvs As String
Public quepdoano As String

Public quepdomes As String
Public quefechaini As String
Public quefechafin As String
Public quenumerodias As String
Public quecodmon As String
Public queremunera_mn As String
Public queremunera_me As String
Public queimporpvsacu_mn As String
Public queimporpvsacu_me As String
Public queimportepvs_mn As String
Public queimportepvs_me As String
Public quefechacan As String
Public queestadodet As String
Public quecodcta_debmn As String
Public quecodcta_habmn As String
Public quecodcta_debme As String
Public quecodcta_habme As String

' 2015 restricciones de uso del sistema campo en tabla plcfgempresa
Public Flag_RestringeSistema As String

Global Quickref As tQuickRef

'[
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const Text_8002 As String = "Instalación de sistema no autorizada"
'INI File Functions...
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Type tQuickRef
    sServidor As String
    sActualizador As String
    cRuta2 As String
    DBFileName As String
    VERSION As Double
    INIFileName As String
End Type

']
Sub Main()
  Dim pnSize As Long
  Dim psArchivo As String, psLinea As String
  Dim pofsoFileCfg As New FileSystemObject, potstFileCfg As TextStream
  
  ' Configuro los objetos de clases
  '  Set gdl_Funcion = CreateObject("syslink.Funciones")
  '  Set gdl_Procedure = CreateObject("syslink.Procedure")
  
  ' Reconoce el directorio de windows.
  ps_WinSystem = Space$(255)
  pnSize = Len(ps_WinSystem)
  GetSystemDirectory ps_WinSystem, pnSize
  ps_WinSystem = Left(ps_WinSystem, Len(Trim(ps_WinSystem)) - 1)
  ps_NomSistema = "Click Premium Personal y Planilla"
  gdl_Procedure.ps_PathImagen = gdl_Funcion.PathApp(App.path) & "bmp\"
  gs_FechaHora = Now

  psArchivo = ps_WinSystem & "\" & pFileSystem

  ' Actualización de versiones
  If dir$(psArchivo) <> "" Then
    Quickref.INIFileName = psArchivo
  Else
    MsgBox "No se encuentra el archivo 'sysmavm.ini'. este archivo es importante para que el programa pueda correr correctamente. Si tu lo encuentras copialo y guardalo en: " & UCase$(App.path) & "." & _
    Chr(13) & Chr(10) & Chr(13) & Chr(10) & "si no funciona tienes que reinstalar el programa.", vbCritical, "Archivo: sysmavm.ini no encontrado..."
    End
  End If
  If Trim(ReadIni("sServidor", "sServidor")) <> "" Then Quickref.sServidor = ReadIni("sServidor", "sServidor") Else Quickref.sServidor = "192.168.1.189"
  If Trim(ReadIni("sActualizador", "sActualizador")) <> "" Then Quickref.sActualizador = ReadIni("sActualizador", "sActualizador") Else Quickref.sActualizador = "192.168.1.31"
  If Trim(ReadIni("cRuta2", "cRuta2")) <> "" Then Quickref.cRuta2 = ReadIni("cRuta2", "cRuta2") Else Quickref.cRuta2 = "planilla"
  If Trim(ReadIni("VERSION", "VERSION")) <> "" Then Quickref.VERSION = ReadIni("VERSION", "VERSION") Else Quickref.VERSION = 1
  
  ' Muestro la Pantalla de Presentacion del Sistema
  fInicio.Show vbModal

  ' Verifico que exista el Archivo de Configuracion
  If StrConv(dir$(psArchivo, vbHidden), vbLowerCase) <> LCase(pFileSystem) Then
    MsgBox Text_8002, vbCritical
    n_SwConfigura = 1
    fPassword.Show vbModal
  End If
  
  n_SwConfigura = 0
  
  ' Abro Archivo de Configuracion
  Set potstFileCfg = pofsoFileCfg.OpenTextFile(psArchivo, ForReading)
  
  Do While Not potstFileCfg.AtEndOfStream
    psLinea = potstFileCfg.ReadLine
    If Left$(psLinea, 10) = "[Planilla]" Then ps_PathSystem = Trim(Mid$(psLinea, InStr(psLinea, "=") + 1)): gdl_Procedure.ps_PathReport = ps_PathSystem & "reports\"
    If Left$(psLinea, 11) = "[Proveedor]" Then ps_Provider = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 8) = "[Server]" Then ps_Servidor = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 11) = "[Servercon]" Then ps_ServidorCon = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 8) = "[UserId]" Then ps_UserId = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 10) = "[Password]" Then ps_Password = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 11) = "[BaseDatos]" Then ps_BDSystems = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 10) = "[Licencia]" Then ps_Licencia = Mid$(psLinea, InStr(psLinea, "=") + 1)
    If Left$(psLinea, 10) = "[MysqlExe]" Then ps_MysqlExe = Mid$(psLinea, InStr(psLinea, "=") + 1)
  Loop
  potstFileCfg.Close
  Set pofsoFileCfg = Nothing
  Set potstFileCfg = Nothing
  
  ' Inicializo la cadena de Conexion Nativa de SQL
  ps_ConxionSql = "Provider=MSDASQL.1;Extended Properties=DNS=" & ps_CopyRight & ";desc=;SERVER=" & ps_Servidor & ";UID=" & ps_UserId & ";PASSWORD=" & ps_Password & ";PORT=3306;OPTION=3;STMT=;" & "DATABASE="
  ' Muestro Pantalla de Seguridad y Pantalla de Elección de Empresas
  fPassword.Show vbModal
  If Not pl_Salir Then End
  
  ' Cargo el menu y empresas
  fSelEmpresa.Show vbModal
  If Not pl_Salir Then End
  
  ' Cargo el menu del sistema
  fMenu.Show

End Sub

Function ReadIni(sSection As String, sKeyName As String) As String
  On Local Error Resume Next
  Dim sRet As String
  sRet = String(255, Chr(0))
  ReadIni = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), Quickref.INIFileName))
End Function

Function ReadServerIni(sSection As String, sKeyName As String) As String
  On Local Error Resume Next
  Dim sRet As String
  sRet = String(255, Chr(0))
  ReadServerIni = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), "\\" & Quickref.sActualizador & "\usuarios\planilla\sysmavm.ini"))
End Function

