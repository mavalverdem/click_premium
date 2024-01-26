VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "conexion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                             ' Declarar variable antes de usarla
Public ps_CadenaConexion As String          ' cadena de conexión

Private s_Driver As String                  ' Driver de conexión
Private s_Proveedor As String               ' proveedor de conexión
Private s_NameDsn As String                 ' Nombre de dsn de conexion
Private s_Servidor As String                ' servidor de base de datos de conexión
Private s_BaseDatos As String               ' nombre de base de datos a conectar
Private s_Usuario As String                 ' user id para la conexión
Private s_Password As String                ' password para la conexión
Public cn_Conexion As ADODB.Connection      ' conexión de base de datos
Private porstRecord As ADODB.Recordset      ' recordset de selección
Public cm_Comando As ADODB.Command          ' comando de ejecucion procedimientos

Public Enum TipoExecution                   ' Tipo de acciona realizar
  Inserta = 0: Modifica = 1: Elimina = 2: Seleccion = 3
End Enum
'[
Public Property Let BaseDatos(ByVal sBaseDatos As String)
  s_BaseDatos = sBaseDatos
End Property
Public Property Let CadenaConexion(ByVal sCadenaConexion As String)
  ps_CadenaConexion = sCadenaConexion
End Property
Public Property Get CadenaConexion() As String
  CadenaConexion = ps_CadenaConexion
End Property
Public Property Let Driver(ByVal sDriver As String)
  s_Driver = sDriver
End Property
Public Property Let NameDsn(ByVal sNombreDsn As String)
  s_NameDsn = sNombreDsn
End Property
Public Property Let Password(ByVal sPassword As String)
  s_Password = sPassword
End Property
Public Property Let Proveedor(ByVal sProveedor As String)
  s_Proveedor = UCase(sProveedor)
End Property
Public Property Let Servidor(ByVal sServidor As String)
  s_Servidor = sServidor
End Property
Public Property Let Usuario(ByVal sUsuario As String)
  s_Usuario = sUsuario
End Property
']
Function ActivaConexion() As Boolean
  Dim s_Connection As String
  
  ActivaConexion = False
      
  ' Verifico que Exista el Servidor y el User Id
  If s_Proveedor = "" Then Beep: MsgBox "No se Encuentro el Nombre del Provedor de Servicio", vbCritical: Exit Function
  If s_Servidor = "" Then Beep: MsgBox "No se Encuentro el Nombre del Servidor", vbCritical: Exit Function
  If s_Usuario = "" Then Beep: MsgBox "No se Encuentro el Nombre del Usuario", vbCritical: Exit Function
  If s_BaseDatos = "" Then Beep: MsgBox "No se Encuentro el Nombre de la Base de Datos", vbCritical: Exit Function
  
  Set cn_Conexion = New ADODB.Connection
  Select Case UCase$(s_Proveedor)
   Case "MICROSOFT.JET.OLEDB.3.5", "MICROSOFT.JET.OLEDB.4.0"
    s_Connection = "data source=" & s_BaseDatos & ".mdb" & ";user id=" & s_Usuario & ";password=" & s_Password & ";Persist Security Info=False;"
   Case "SQLOLEDB", "MSDAORA", "MICROSOFT OLE DB PROVIDER FOR SQL SERVER"
    s_Connection = "data source=" & s_Servidor & ";initial catalog=" & s_BaseDatos & ";user id=" & s_Usuario & ";password=" & s_Password & ";persist security info=False;"
   Case "MSDASQL.1"
    s_Connection = "driver=" & s_Driver & ";server=" & s_Servidor & ";port=3306;USER=" & s_Usuario & ";PWD=" & s_Password & ";DATABASE=" & s_BaseDatos & ";option=3;persist security info=False;"
   Case Else
    Exit Function
  End Select
  
  With cn_Conexion
    .Provider = s_Proveedor
    .ConnectionString = s_Connection
    .ConnectionTimeout = 30
    .CommandTimeout = 300
    .CursorLocation = adUseClient
  End With
  
  On Error GoTo ErrorConexion
  
  If cn_Conexion.State = adStateOpen Then cn_Conexion.Close
  cn_Conexion.Open
  ps_CadenaConexion = cn_Conexion
  ' modifico w7
  ps_CadenaConexion = s_Connection
  ActivaConexion = True
  Exit Function
  
ErrorConexion:
  ErrorClass "Los Datos de la Conexión no son Correctos"

End Function
Sub CancelaTransaccion()
  cn_Conexion.RollbackTrans
End Sub
Sub ConfirmaTransaccion()
  cn_Conexion.CommitTrans
End Sub
Function CommandParameter(ByVal s_Procedure As String, ByVal a_Parameter, ByVal a_Type, ByVal a_Direction, ByVal a_Size) As Boolean
  Static n_Index As Integer, n_Limite As Integer
  Static s_NameParameter As String
  
  On Error GoTo ErrorParameters
    
  CommandParameter = False
  Set cm_Comando = New ADODB.Command
  
  ' Elimina los parameters
  Do While cm_Comando.Parameters.Count <> 0: cm_Comando.Parameters.Delete 0: Loop
  ' Crea los parameters de acuerdo a los parametros
  With cm_Comando
    .ActiveConnection = cn_Conexion
    .CommandText = s_Procedure
    .CommandType = adCmdStoredProc
    .Prepared = False
    '  .Parameters.Refresh
    '  .Parameters.Item(1).Size = ""
    ' Creo los parameters con sus respectivos valores
    n_Limite = UBound(a_Parameter, 1)
    For n_Index = 0 To n_Limite
      s_NameParameter = "parame" & Format(n_Index, "00")
      .Parameters.Append .CreateParameter(s_NameParameter, a_Type(n_Index), a_Direction(n_Index), a_Size(n_Index), a_Parameter(n_Index))
    Next n_Index
    .Execute
  End With
  
  CommandParameter = True
  GoTo Salir

ErrorParameters:
  ErrorClass "Ocurrio un error al Ejecutar " & s_Procedure
  Set cm_Comando = Nothing
  GoTo Salir
Salir:
  On Error GoTo 0
  
End Function
Private Sub ErrorClass(ByVal s_Menssage As String)
  ' Muestro el mensaje de error
  MsgBox s_Menssage & ", Verifique . . ." & vbCrLf & vbCrLf & _
  "Fuente: " & Err.Source & vbCrLf & _
  "Numero: " & Err.Number & vbCrLf & _
  "Descripcion: " & Err.Description & vbCrLf, vbCritical, "ERROR"
End Sub
Function Execucion(ByVal s_Sentencia As String, Optional ByVal b_TipoAccion As TipoExecution) As Boolean
  Static n_RecordAffected As Long
  Static s_Mensaje As String

  Execucion = False

  On Error GoTo ErrorExecution

  ' Ejecuto la sentencia y devuelvo el numero de registros afectados
  cn_Conexion.Execute s_Sentencia, n_RecordAffected
  Execucion = True
  GoTo Salir

ErrorExecution:
  s_Mensaje = "Ocurrio un error al "
  Select Case b_TipoAccion
   Case TipoExecution.Inserta: s_Mensaje = s_Mensaje & "Actualizar"
   Case TipoExecution.Modifica: s_Mensaje = s_Mensaje & "Actualizar"
   Case TipoExecution.Elimina: s_Mensaje = s_Mensaje & "Eliminar"
   Case TipoExecution.Seleccion: s_Mensaje = s_Mensaje & "Seleccionar"
  End Select
  ' Ejecuto el error
  ErrorClass s_Mensaje
  GoTo Salir
Salir:
  On Error GoTo 0

End Function
Sub IniciaTransaccion()
  cn_Conexion.BeginTrans
End Sub
Private Sub Class_Terminate()
  If cn_Conexion.State = adStateOpen Then cn_Conexion.Close
  Set cn_Conexion = Nothing
End Sub
Function Recordset(ByVal n_CursorType As CursorTypeEnum, ByVal n_LockType As LockTypeEnum, ByVal n_Location As CursorLocationEnum, ByVal sSentencia As String) As ADODB.Recordset

  On Error GoTo ErrorRecordset

  ' iguala el nuevo cursor para la selección
  Set porstRecord = New ADODB.Recordset
  With porstRecord
    If .State = adStateOpen Then .Close
    .CursorType = n_CursorType
    .LockType = n_LockType
    .CursorLocation = n_Location
    .Open sSentencia, cn_Conexion
  End With
  Set Recordset = porstRecord
  Set porstRecord = Nothing
  GoTo Salir

ErrorRecordset:
  ErrorClass "Ocurrio un error al Generar Recordset "
  Set porstRecord = Nothing
  GoTo Salir
Salir:
  On Error GoTo 0

End Function
