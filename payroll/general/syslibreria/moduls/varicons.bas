Attribute VB_Name = "Definicion"
' Definici�n de variables
Public sSQL As String                         ' Cadena de sentencia sql para ejecuci�n
Public gdl_Funcion As Object                  ' Control o clase de libreria de funciones

' definici�n de constantes
Public Const s_Sigla As String = "sysma"                    ' Siglas de los sistemas
Public Const n_BackColorActive As Double = &HFFFFFF         ' Color de Texto en Modo Edici�n
Public Const n_BackColorInactive As Double = &HC7D8E0       ' Color de Texto en Modo Locked
Public Const n_ForeColorActive As Double = &HC00000         ' Color de Texto en Modo Edici�n
Public Const n_ForeColorInactive As Double = &HC00000       ' Color de Texto en Modo Locked
Public Const n_BackColorHelp As Double = &H80000018         ' Color de Grilla de Ayuda

Public Const s_Estado_Ina As String = "0"                   ' Estado de registro inactivo
Public Const s_Estado_Act As String = "1"                   ' Estado de registro activo
Public Const s_Estado_Blq As String = "2"                   ' Estado de registro activo

Public Const s_Codmon_mn As String = "N"                    ' Tipo de moneda nacional
Public Const s_Codmon_me As String = "E"                    ' Tipo de moneda extranjera
Public Const s_Codmon_mn_Txt As String = "S/."              ' Signo de moneda nacional
Public Const s_Codmon_me_Txt As String = "US$"              ' Signo de moneda extranjera

Public Const s_PeriodoRemAper As String = "00000000"        ' Periodo Remuneraciones Anteriores
Public Const s_ProcesoRemAper As String = "00"              ' Proceso de Calculo de Remuneraciones Anteriores
