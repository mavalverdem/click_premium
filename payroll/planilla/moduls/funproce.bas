Attribute VB_Name = "mdlFunProce"
Option Explicit

Public sys_lst_func()
Public sys_num_func As Integer
Public sys_lst_const()
Public sys_num_const As Integer
Public sys_lst_concpt()
Public sys_num_concpt As Integer
Public sys_lst_valores()
Public sys_num_valores As Integer

'[
Private porstRecord As ADODB.Recordset      ' Recordset de Resultado(WithEvents)
Private cm_Comando As ADODB.Command         ' Comando de ejecución de sentecias
Private s_Sentencia As String               ' Cadena de Sentencia Sql

' Declaraciones del Api
Private Declare Function SendMessageLongRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub SendKeysEvent Lib "user32" Alias "keybd_event" (ByVal bKey As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Function Valida_CierrePeriodo(ClasePlanilla As String, CodPeridoc As String, TipoPeriodo) As String
  'Retorna si el valor del periodo que se recibe como parametro es abierto, cerrado
  '0 = NO PROCESADO
  '1 = PROCESADO
  '2 = CERRADO
  
  Dim Ssql_cp As String
  Dim Est_CierrePeriodo, Des_CierrePeriodo As String
  
  Ssql_cp = "SELECT codcls, codpdo, anopdo, mespdo, estadopdo FROM plperiodo WHERE codcls='" & ClasePlanilla & "' AND codpdo='" & CodPeridoc & "' AND tpopdo='" & TipoPeriodo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, Ssql_cp)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    Est_CierrePeriodo = Val(porstRecordset!estadopdo)
    If (Est_CierrePeriodo = 0 Or Est_CierrePeriodo = 1) Then
      Valida_CierrePeriodo = "P.XTRABAJAR"
    ElseIf Est_CierrePeriodo = 2 Then
      Valida_CierrePeriodo = "P.CERRADO"
    End If
  Else
    Valida_CierrePeriodo = "P.NO_REGISTRADO"
  End If
End Function
Function Valida_LicenciaUso(ByVal S_AnioProcesosSIS As String, ByVal ps_Fecha_LimiteProc As String, Optional ByVal s_Pmes As String, Optional ByVal s_Panio As String) As Boolean
  Dim s_Dia As String
  Dim s_FechaProcesoreg As String
  
  ' Retorno Verdader significa que el parametro de fecha que limita procesos es menor que la fecha que se usara en el registro.
  ' Valido Si la fecha que da el sistemas es mayor a la que Fecha limite proceso segun valor registrado en la BD.
  If CDate(Format(ps_Fecha_LimiteProc, "dd/mm/yyyy")) > CDate(Format(Date, "dd/mm/yyyy")) Then
    Valida_LicenciaUso = True
  Else
    Valida_LicenciaUso = False
  End If
  
  ' Si la fecha del Windows es menor a la fecha del añor de trabajo actual en el sistemas
  'If S_AnioProcesosSIS <= Year(ps_Fecha_LimiteProc) Then
  '   Valida_LicenciaUso = True
  'End If
  ' Valido que los campos que contines el periodo a registar sean menores a la fecha del imite proceso segun valor registrado en la BD.
  If (Len(s_Pmes) > 0) And (Len(s_Panio) > 0 And s_Pmes <> "00") Then
    s_Dia = Trim$("01")
    s_Pmes = Trim$(s_Pmes)
    s_Panio = Trim$(s_Panio)
  
    s_FechaProcesoreg = s_Dia & "/" & s_Pmes & "/" & s_Panio
    If CDate(ps_Fecha_LimiteProc) >= CDate(s_FechaProcesoreg) Then
      Valida_LicenciaUso = True
    Else
      Valida_LicenciaUso = False
    End If
  End If

End Function
']
Function EnviaCorreoElectronico(ByVal o_MapiSesion As Object, ByVal o_MapiMensaje As Object, ByVal a_Destinatario, ByVal s_Asunto As String, ByVal s_Adjunto As String, ByVal s_Mensaje As String) As Boolean
  Dim nIntervalo As Integer
  
  With o_MapiSesion
    .NewSession = False
    .SignOn
  End With
  
  With o_MapiMensaje
    .SessionID = o_MapiSesion.SessionID
    ' Creamos el mensaje
    .Compose
    ' Asunto del mensaje
    .MsgSubject = s_Asunto
    ' Mensaje
    .MsgNoteText = s_Mensaje
    For nIntervalo = 0 To UBound(a_Destinatario, 2)
      ' Nombre del Mail del destinatario
      .RecipIndex = o_MapiMensaje.RecipCount
      .RecipDisplayName = a_Destinatario(nIntervalo)
    Next nIntervalo
    ' Archivo Adjunto
    If s_Adjunto <> "" Then
      .AttachmentPathName = s_Adjunto
    End If
    ' Enviamos el correo
    .Send False
  End With
  ' Cerramos la sesión abierta del Mapi
  o_MapiSesion.SignOff

End Function
Function EnviaCorreoCDOWeb(ByVal n_Servidor As Integer, ByVal s_Usuario As String, ByVal s_Password As String, ByVal s_Remitente As String, ByVal s_Destinatario As String, ByVal s_Asunto As String, ByVal s_Mensaje As String, Optional ByVal s_Copia As String, Optional ByVal s_CopiaOculta As String, Optional s_Adjunto As String, Optional n_Puerto As Integer = 25, Optional b_Importancia As Byte = 1) As Boolean
  Dim oMensaje As Object
  Dim aMatriz() As String, sServidor As String
  Dim nSecuencia As Long

  On Error GoTo EnviaCorreoCDOWeb_TratamientoErrores

  sServidor = Choose(n_Servidor, "smtp.gmail.com", "smtp.live.com")
  Set oMensaje = CreateObject("CDO.Message")
  
  ' configuración de CDO
  oMensaje.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = sServidor
  oMensaje.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  
  ' Configuración SMTP
  With oMensaje.Configuration.Fields
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = n_Puerto
    ' Tipo autentificación con el servidor de correo 0:no requiere autentificacion; 1:con autentificación
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    ' Tiempo máximo espera en segundos para la conexión
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
    ' Usuario servidor Smtp
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = s_Usuario
    ' Password cuenta
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = s_Password
    ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    ' Prioridad -1=Low, 0=Normal, 1=High
    .Item("urn:schemas:httpmail:priority") = b_Importancia
    .Item("urn:schemas:mailheader:X-Priority") = b_Importancia
    ' Importancia 0=Low, 1=Normal, 2=High
    .Item("urn:schemas:httpmail:importance") = b_Importancia
    ' Actualiza los datos antes de enviar
    .Update
  End With
  
  ' Mensaje
  oMensaje.From = s_Remitente
  oMensaje.To = s_Destinatario
  If Not Trim(s_Copia) = vbNullString Then
    oMensaje.CC = s_Copia
  End If
  If Not Trim(s_CopiaOculta) = vbNullString Then
    oMensaje.BCC = s_CopiaOculta
  End If
  oMensaje.Subject = s_Asunto
  oMensaje.TextBody = s_Mensaje
  ' adjunto/s si existe
  If Not Trim(s_Adjunto) = vbNullString Then
    aMatriz() = Split(s_Adjunto, ";")
    For nSecuencia = 0 To UBound(aMatriz)
      If Not dir$(Trim(aMatriz(nSecuencia))) = vbNullString Then oMensaje.AddAttachment Trim$(aMatriz(nSecuencia))
    Next nSecuencia
  End If
  ' Envío el mensaje
  oMensaje.Send
  EnviaCorreoCDOWeb = True

EnviaCorreoCDOWeb_Salir:
  Set oMensaje = Nothing
  On Error GoTo 0
  Exit Function
   
EnviaCorreoCDOWeb_TratamientoErrores:
  EnviaCorreoCDOWeb = False
  MsgBox "Error " & Err.Number & " proceso de enviar correo (" & Err.Description & ")", vbCritical + vbOKOnly
  Resume EnviaCorreoCDOWeb_Salir

End Function
Sub EnviaCorreoOutlook(ByVal s_Destinatario As String, ByVal s_Asunto As String, ByVal s_Mensaje As String, Optional ByVal s_Copia As String, Optional ByVal s_CopiaOculta As String, Optional s_Adjunto As String, Optional b_Importancia As Byte = 1)
  
  Dim oAplicacion As Object           ' Outlook.Application
  
  Dim oMensaje As Object             ' Outlook.MailItem
  Dim oDestinatario As Object        ' Outlook.Recipient
  Dim oAdjunto As Object               ' Outlook.Attachment
  Dim aMatriz() As String, sMsgError As String
  Dim nSecuencia As Long
  
  'On Error GoTo EnviaCorreoOutlook_TratamientoErrores

  ' valido destinatarios
  If Trim(s_Destinatario) = vbNullString Then
    MsgBox "No hay ningún destinatario; Verificar", vbExclamation + vbOKOnly
    GoTo EnviaCorreoOutlook_Salir
  End If
  
  ' genero instancia outlook
  Set oAplicacion = CreateObject("Outlook.Application")
  ' creo un mensaje
  Set oMensaje = oAplicacion.CreateItem(0)
  With oMensaje
    ' añado destinatarios
    If Not Trim(s_Destinatario) = vbNullString Then
      aMatriz() = Split(s_Destinatario, ";")
      For nSecuencia = 0 To UBound(aMatriz)
        Set oDestinatario = .Recipients.Add(aMatriz(nSecuencia))
        oDestinatario.Type = 1
      Next nSecuencia
    End If
    If Not Trim(s_Copia) = vbNullString Then
      aMatriz() = Split(s_Copia, ";")
      For nSecuencia = 0 To UBound(aMatriz)
        Set oDestinatario = .Recipients.Add(aMatriz(nSecuencia))
        oDestinatario.Type = 2
      Next nSecuencia
    End If
    If Not Trim(s_CopiaOculta) = vbNullString Then
      aMatriz() = Split(s_CopiaOculta, ";")
      For nSecuencia = 0 To UBound(aMatriz)
        Set oDestinatario = .Recipients.Add(aMatriz(nSecuencia))
        oDestinatario.Type = 3
      Next nSecuencia
    End If
    ' aplico el asunto, el mensaje y la importancia del mensaje
    .Subject = s_Asunto
    .Body = s_Mensaje
    .Importance = b_Importancia
  
    ' añado los adjuntos si los hubiera
    If Not Trim(s_Adjunto) = vbNullString Then
      aMatriz() = Split(s_Adjunto, ";")
      For nSecuencia = 0 To UBound(aMatriz)
        Set oAdjunto = .Attachments.Add(aMatriz(nSecuencia))
      Next nSecuencia
    End If
    
    ' verifico la validez de los destinatarios
    For Each oDestinatario In .Recipients
      oDestinatario.Resolve
      If Not oDestinatario.Resolved Then
        sMsgError = sMsgError & Chr$(34) & oDestinatario & Chr$(34) & ", "
      End If
    Next
    If Len(sMsgError) > 0 Then
      MsgBox "Los siguientes destinatarios no son correctos" & vbNewLine & Left$(sMsgError, Len(sMsgError) - 2), vbCritical + vbOKOnly
      .Display   ' muestro el mensaje para dar opción a rectificar el error
      GoTo EnviaCorreoOutlook_Salir
    End If
    ' envío el mensaje
    .Send
  End With

EnviaCorreoOutlook_Salir:
  ' cierro objetos
  Set oMensaje = Nothing
  Set oAplicacion = Nothing
  On Error GoTo 0
  Exit Sub

'EnviaCorreoOutlook_TratamientoErrores:
'  If Err = 287 Then
'    MsgBox "Outlook debe estar abierto para que el proceso funcione", vbCritical + vbOKOnly
'  Else
'    MsgBox "Error " & Err.Number & " proceso de enviar correo (" & Err.Description & ")", vbCritical + vbOKOnly
'  End If
'  Resume EnviaCorreoOutlook_Salir
'  Resume Next
End Sub
Sub EnviarTecla(ByVal n_Tecla As Long)
  SendKeysEvent n_Tecla, 0, 0, 0
  SendKeysEvent n_Tecla, 0, KEYEVENTF_KEYUP, 0
End Sub
Function FormVisible(ByVal s_Formulario As String) As Boolean
  Dim i As Integer
  ' Barro todos los Formularios cargadas
  FormVisible = False
  For i = 0 To Forms.Count - 1
    If Forms(i).Name = s_Formulario Then
      FormVisible = True
      Exit Function
    End If
  Next i
End Function
Sub IniciaCalculo()
  
  Dim sSQL As String
  Dim nTotalFunciones As Long, nTotalVariables As Long
  Dim nTotalConceptos As Long, nTotalValores As Long
  Dim nPosition As Integer

  nTotalFunciones = 0
  ' Obtengo el número de funciones de cálculo
  sSQL = "SELECT COUNT(*) AS nRegistros FROM plvarfunc WHERE tipo = 'F'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    nTotalFunciones = CLng(porstRecordset!nRegistros)
    porstRecordset.Close
  End If
  
  nTotalVariables = 0
  ' Obtengo el número de variables de cálculo
  sSQL = "SELECT COUNT(*) AS nRegistros FROM plvarfunc WHERE tipo = 'V'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    nTotalVariables = CLng(porstRecordset!nRegistros)
    porstRecordset.Close
  End If
  
  nTotalConceptos = 0
  ' Obtengo el número de conceptos de cálculo
  sSQL = "SELECT COUNT(*) AS nRegistros "
  sSQL = sSQL & "FROM plconceplanilla cxc, plconcepto cpc "
  sSQL = sSQL & "WHERE cxc.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND cpc.codcpc=cxc.codcpc"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    nTotalConceptos = CLng(porstRecordset!nRegistros)
    porstRecordset.Close
  End If
  
  nTotalValores = 0
  ' Obtengo el número de valores de tablas basicas de cálculo
  sSQL = "SELECT COUNT(*) AS nRegistros "
  sSQL = sSQL & "FROM pltablabase tbl "
  sSQL = sSQL & "WHERE tbl.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND tbl.pdoano='" & ps_Anyo & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    nTotalValores = CLng(porstRecordset!nRegistros)
    porstRecordset.Close
  End If
  
  nPosition = 1
  ' Carga las funciones de cálculo
  sSQL = "SELECT nombre FROM plvarfunc WHERE tipo = 'F' ORDER BY orden"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  ReDim sys_lst_func(1 To nTotalFunciones)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    Do While Not porstRecordset.EOF
      sys_lst_func(nPosition) = porstRecordset!nombre
      nPosition = nPosition + 1
      porstRecordset.MoveNext
    Loop
    porstRecordset.Close
  End If
    
  nPosition = 1
  ' Cargo las variables de cálculo
  sSQL = "SELECT nombre FROM plvarfunc WHERE tipo = 'V' ORDER BY orden"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  ReDim sys_lst_const(1 To nTotalVariables)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    Do While Not porstRecordset.EOF
      sys_lst_const(nPosition) = porstRecordset!nombre
      nPosition = nPosition + 1
      porstRecordset.MoveNext
    Loop
    porstRecordset.Close
  End If
  
  nPosition = 1
  ' Cargo los conceptos de cálculo
  sSQL = "SELECT cxc.codcpc "
  sSQL = sSQL & "FROM plconceplanilla cxc, plconcepto cpc "
  sSQL = sSQL & "WHERE cxc.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND cpc.codcpc=cxc.codcpc"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  ReDim sys_lst_concpt(1 To nTotalConceptos)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    Do While Not porstRecordset.EOF
      sys_lst_concpt(nPosition) = "C" & porstRecordset!codcpc
      nPosition = nPosition + 1
      porstRecordset.MoveNext
    Loop
    porstRecordset.Close
  End If
  
  nPosition = 1
  ' Cargos los valores de tablas basicas de cálculo
  sSQL = "SELECT tbl.codtbl "
  sSQL = sSQL & "FROM pltablabase tbl "
  sSQL = sSQL & "WHERE tbl.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND tbl.pdoano='" & ps_Anyo & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  ReDim sys_lst_valores(1 To nTotalValores)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    Do While Not porstRecordset.EOF
      sys_lst_valores(nPosition) = "K_" & porstRecordset!codtbl
      nPosition = nPosition + 1
      porstRecordset.MoveNext
    Loop
    porstRecordset.Close
  End If
  sys_num_func = nTotalFunciones
  sys_num_const = nTotalVariables
  sys_num_concpt = nTotalConceptos
  sys_num_valores = nTotalValores

End Sub
Sub Inputbox_Password(ByVal o_Formulario As Form)
  SetTimer o_Formulario.hwnd, &H5000&, 100, AddressOf TimerProceso
End Sub
Sub LoadOpcion(ByVal o_Form As Form, ByVal s_Opcion As String, ByVal s_Index As String, ByVal n_Proceso As Integer, Optional ByVal n_Show As Integer)
  
  On Error GoTo Err
  
  Dim n_Posicion As Integer, s_CapOption As String
   
  s_Sentencia = "SELECT mnu.codmdl, mnu.opcion, mnu.orden, mnu.detmdl"
  s_Sentencia = s_Sentencia & " FROM sgmdl mnu LEFT JOIN sgpms opc USING(codsis, codmdl)"
  s_Sentencia = s_Sentencia & " WHERE mnu.codsis='" & ps_CodSistema & "'"
  s_Sentencia = s_Sentencia & " AND mnu.opcion='" & s_Opcion & "'"
  s_Sentencia = s_Sentencia & " AND mnu.orden='" & Format(s_Index, "00") & "'"
  s_Sentencia = s_Sentencia & " AND opc.codemp='" & ps_CodEmpresa & "'"
  s_Sentencia = s_Sentencia & " AND opc.codusr='" & ps_Usuario & "'"
  
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenKeyset, adLockReadOnly, adUseClient, s_Sentencia)
  
  If (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount = 0 Then
    MsgBox "Opción restringida", vbInformation
    ' Cierro Tabla de Usuario
    porstRecordset.Close: Set porstRecordset = Nothing
    Exit Sub
  End If
  
  s_CapOption = porstRecordset!detmdl
  ' Cierro Tabla de Usuario
  porstRecordset.Close: Set porstRecordset = Nothing
  
  n_Posicion = InStr(s_CapOption, "&")
  If n_Posicion <> 0 Then
    s_CapOption = Mid(s_CapOption, 1, n_Posicion - 1) & Mid(s_CapOption, n_Posicion + 1)
  End If
  
  o_Form.Caption = Choose(n_Proceso + 1, "Registro ", "Proceso ", "Reporte ", "Listado ", "Consulta ", "Provisión ", "") & s_CapOption
  o_Form.Show n_Show

  Exit Sub
Err:
  MsgBox "Error: " & Err.Number & " : " & Err.Description

End Sub
Sub MuestraMensaje(ByVal s_Mensaje As String)
  fMenu.panMessage.Caption = " " & Trim$(s_Mensaje)
End Sub
Function OpenConnection(ByVal s_Servidor, ByVal s_BaseDatos As String) As String
  Dim n_PosicionIni As Integer, n_PosicionFin As Integer
  Dim s_Expresion As String
  
  ' Instancia de Componente de conexion
  Set gdl_Conexion = CreateObject("conexionbd.conexion")    ' No se refencia crea objeto
  'Set gdl_Conexion = New Conexionbd.conexion                ' Se refencia proyecto clase
  With gdl_Conexion
    .Proveedor = ps_Provider
    .Driver = ps_DriverCnn
    .NameDsn = ps_NombreDsn
    .Servidor = s_Servidor
    .BaseDatos = s_BaseDatos
    .Usuario = ps_UserId
    .Password = ps_Password
    If Not .ActivaConexion Then Exit Function
  End With
  ' Incremento el passwor si no existe
  s_Expresion = "PASSWORD="
  n_PosicionIni = InStr(gdl_Conexion.ps_CadenaConexion, s_Expresion)
  If n_PosicionIni = 0 Then s_Expresion = "PWD="
  n_PosicionIni = InStr(gdl_Conexion.ps_CadenaConexion, s_Expresion)
  gdl_Conexion.ps_CadenaConexion = gdl_Conexion.ps_CadenaConexion & IIf(n_PosicionIni = 0, ";PWD=" & ps_Password, "")
  ' Obtengo cadena con base de datos
  s_Expresion = "DB="
  n_PosicionIni = InStr(gdl_Conexion.ps_CadenaConexion, s_Expresion)
  If n_PosicionIni = 0 Then s_Expresion = "DATABASE="
  n_PosicionIni = InStr(gdl_Conexion.ps_CadenaConexion, s_Expresion)
  n_PosicionFin = InStr(n_PosicionIni, gdl_Conexion.ps_CadenaConexion, ";")
  OpenConnection = Left(gdl_Conexion.ps_CadenaConexion, (n_PosicionIni - 1)) & Mid(gdl_Conexion.ps_CadenaConexion, (n_PosicionFin + 1)) & ";" & s_Expresion
  
End Function
Function OpenRecordset(ByVal s_Conexion As String, ByVal n_CursorType As CursorTypeEnum, ByVal n_LockType As LockTypeEnum, ByVal n_Location As CursorLocationEnum, ByVal s_Sentencia As String) As ADODB.Recordset

  ' iguala el nuevo cursor para la selección
  Set porstRecord = New ADODB.Recordset
  With porstRecord
    If .State = adStateOpen Then .Close
    .CursorType = n_CursorType
    .LockType = n_LockType
    .CursorLocation = n_Location
    .Open s_Sentencia, s_Conexion
  End With
  Set OpenRecordset = porstRecord
  Set porstRecord = Nothing

End Function
Sub ReadImagen(ByVal o_rstImagen As ADODB.Recordset, ByVal o_Imagen As Object, ByVal s_Campo As String)
  Dim o_Stream As ADODB.Stream
  ' Inicializo y verifico parametros
  o_Imagen.Picture = LoadPicture()
  If Not (o_rstImagen Is Nothing) And s_Campo <> "" Then
    If o_rstImagen.State = adStateClosed Then GoTo ErrImagen
    If IsNull(o_rstImagen.Fields(s_Campo).Value) Then GoTo ErrImagen
    On Error GoTo ErrImagen
    
    ' Instancio los objetos y propiedades
    Set o_Stream = New ADODB.Stream
    
    ' Cargo la imagene en el objeto
    o_Stream.Type = adTypeBinary
    o_Stream.Open
    o_Stream.Write o_rstImagen.Fields(s_Campo).Value
    ' Guardo la imagen en un archivo temporal
    o_Stream.SaveToFile "imgtempo", adSaveCreateOverWrite
    ' Cierro el objeto
    o_Stream.Close
    
    ' Cargo la imagen en el control
    o_Imagen.Picture = LoadPicture("imgtempo")
    ' Elimino el archivo temporal
    If dir$("imgtempo", vbNormal) <> "" Then Kill "imgtempo"
  End If
  o_Imagen.Refresh
  
ErrImagen:
  If Err.Number <> s_Estado_Ina Then: MsgBox Err.Description
  Set o_Stream = Nothing

End Sub
Function Records_Del(ByVal s_Tabla As String, ByVal a_Condicion, ByVal a_Valores, ByVal a_Tipos) As Boolean
  Dim s_CadCondicion As String
  Dim n_Index As Integer

  Records_Del = False
  s_Sentencia = "DELETE FROM " & s_Tabla
  s_CadCondicion = ""
  For n_Index = 0 To UBound(a_Condicion)
    s_CadCondicion = s_CadCondicion & a_Condicion(n_Index) & " = "
    ' Inicio de la cadena de cada valor
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_CadCondicion = s_CadCondicion & "'"
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        s_CadCondicion = s_CadCondicion & "CONVERT(DATETIME, '"
      Else
        s_CadCondicion = s_CadCondicion & "NULL"
      End If
    End If
    ' Valores de cada campo
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_CadCondicion = s_CadCondicion & gdl_Funcion.SacaEntRetApos(a_Valores(n_Index))
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        s_CadCondicion = s_CadCondicion & a_Valores(n_Index)
      End If
    ElseIf a_Tipos(n_Index) = TipoDato.Logico Then
      s_CadCondicion = s_CadCondicion & IIf(a_Valores(n_Index), 1, 0)
    Else
      s_CadCondicion = s_CadCondicion & a_Valores(n_Index)
    End If
    ' Fin de la cadena de cada valor
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_CadCondicion = s_CadCondicion & "'"
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        s_CadCondicion = s_CadCondicion & "', 103)"
      End If
    End If
    If n_Index <> UBound(a_Condicion) Then s_CadCondicion = s_CadCondicion & " AND "
  Next n_Index
  s_Sentencia = s_Sentencia & " WHERE " & s_CadCondicion
  ' Ejecuto la sentencia
  Records_Del = gdl_Conexion.Execucion(s_Sentencia, Elimina)

End Function
Function Records_Ins(ByVal s_Tabla As String, ByVal a_Campos, ByVal a_Valores, ByVal a_Tipos) As Boolean
  Dim s_CadCampos As String, s_CadValores As String
  Dim n_Index As Integer
  Dim s_Valor As String, s_Sentencia As String
  
  Records_Ins = False
  s_Sentencia = "INSERT INTO " & s_Tabla & " "
  s_CadCampos = "(": s_CadValores = "("
  For n_Index = 0 To UBound(a_Campos)
    s_CadCampos = s_CadCampos & a_Campos(n_Index)
    If n_Index <> UBound(a_Campos) Then s_CadCampos = s_CadCampos & ", "
    ' Inicio de la cadena de cada valor
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_Valor = gdl_Funcion.SacaEntRetApos(a_Valores(n_Index))
      s_CadValores = s_CadValores & IIf(s_Valor = "", s_Valor, "'")
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        s_CadValores = s_CadValores & "DATE_FORMAT('"
      Else
        s_CadValores = s_CadValores & "NULL"
      End If
    End If
    ' Valores de cada campo
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_CadValores = s_CadValores & IIf(s_Valor = "", "NULL", s_Valor)
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        s_CadValores = s_CadValores & a_Valores(n_Index)
      End If
    ElseIf a_Tipos(n_Index) = TipoDato.Logico Then
      s_CadValores = s_CadValores & IIf(a_Valores(n_Index), 1, 0)
    Else
      s_CadValores = s_CadValores & a_Valores(n_Index)
    End If
    ' Fin de la cadena de cada valor
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_CadValores = s_CadValores & IIf(s_Valor = "", s_Valor, "'")
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        s_CadValores = s_CadValores & "', '" & s_FmtFechMysql_1 & "')"
      End If
    End If
    If n_Index <> UBound(a_Valores) Then s_CadValores = s_CadValores & ", "
  Next n_Index
  
  s_CadCampos = s_CadCampos & ")"
  s_CadValores = IIf(Right$(s_CadValores, 1) <> ",", s_CadValores, Left$(s_CadValores, Len(s_CadValores) - 1)) & ")"
  s_Sentencia = s_Sentencia & s_CadCampos & " VALUES " & s_CadValores
  ' Ejecuto la sentencia
  Records_Ins = gdl_Conexion.Execucion(s_Sentencia, Inserta)

End Function
Function Records_Upd(ByVal s_Tabla As String, ByVal a_Campos, ByVal a_Valores, ByVal a_Tipos, ByVal a_Condicion) As Boolean
  Dim s_CadCampos As String, s_CadCondicion As String
  Dim n_Index As Integer, n_NumWhere As Integer
  Dim s_Valor As String, s_Sentencia As String

  Records_Upd = False
  s_Sentencia = "UPDATE " & s_Tabla & " SET "
  s_CadCampos = "": s_CadCondicion = "": n_NumWhere = UBound(a_Condicion)
  For n_Index = 0 To UBound(a_Campos)
    If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & a_Campos(n_Index) & "="
    If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & a_Condicion(n_Index) & "="
    ' Inicio de la cadena de cada valor
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      s_Valor = gdl_Funcion.SacaEntRetApos(a_Valores(n_Index))
      If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & IIf(s_Valor = "", s_Valor, "'")
      If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & "'"
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & "DATE_FORMAT('"
        If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & "DATE_FORMAT('"
      Else
        If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & "NULL"
        If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & "NULL"
      End If
    End If
    ' Valores de cada campo
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & IIf(s_Valor = "", "NULL", s_Valor)
      If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & gdl_Funcion.SacaEntRetApos(a_Valores(n_Index))
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & a_Valores(n_Index)
        If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & a_Valores(n_Index)
      End If
    ElseIf a_Tipos(n_Index) = TipoDato.Logico Then
      If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & IIf(a_Valores(n_Index), 1, 0)
      If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & IIf(a_Valores(n_Index), 1, 0)
    Else
      If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & a_Valores(n_Index)
      If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & a_Valores(n_Index)
    End If
    ' Fin de la cadena de cada valor
    If a_Tipos(n_Index) = TipoDato.Caracter Then
      If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & IIf(s_Valor = "", s_Valor, "'")
      If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & "'"
    ElseIf a_Tipos(n_Index) = TipoDato.FECHA Then
      If IsDate(a_Valores(n_Index)) Then
        If n_NumWhere < n_Index Then s_CadCampos = s_CadCampos & "', '" & s_FmtFechMysql_1 & "')"
        If n_NumWhere >= n_Index Then s_CadCondicion = s_CadCondicion & "', '" & s_FmtFechMysql_1 & "')"
      End If
    End If
    If (n_Index <> UBound(a_Campos)) And (n_NumWhere < n_Index) Then s_CadCampos = s_CadCampos & ", "
    If n_Index < n_NumWhere Then s_CadCondicion = s_CadCondicion & " AND "
  Next n_Index
  s_Sentencia = s_Sentencia & s_CadCampos & " WHERE " & s_CadCondicion
  ' Ejecuto la sentencia
  Records_Upd = gdl_Conexion.Execucion(s_Sentencia, Modifica)

End Function
Sub Registro_Texto(ByVal sLinea As String, ByVal nColumnas As Integer, ByRef aRegistros)
  Dim n_Campo As Integer
  Dim n_Inicio As Integer, n_Longitud As Integer
  ReDim Preserve aRegistros(nColumnas)
  
  n_Inicio = 1
  For n_Campo = 1 To nColumnas
    n_Longitud = Abs(InStr(n_Inicio, sLinea, "|") - n_Inicio)
    aRegistros(n_Campo) = Mid$(sLinea, n_Inicio, n_Longitud)
    n_Inicio = n_Inicio + (n_Longitud + 1)
  Next n_Campo

End Sub
Function fRetornaPosArreglo(ByVal o_ArrBusqueda, ByVal nFilaInicial As Long, ByVal nColBusqueda As Long, ByVal sValBusqueda As String) As Long
  Dim nRegistro As Long
  
  fRetornaPosArreglo = 0
  'Bucle para recorrer a través de la matriz
  For nRegistro = nFilaInicial To UBound(o_ArrBusqueda, 2)
    If o_ArrBusqueda(nColBusqueda, nRegistro) = Trim(sValBusqueda) Then
      fRetornaPosArreglo = nRegistro
      Exit For
    End If
  Next nRegistro

End Function
Function fSeleccionDirectorio() As String
  On Error Resume Next ' por si el usuario pulsa {esc} y no selecciona nada
  With CreateObject("shell.application")
    fSeleccionDirectorio = .BrowseForFolder(0, "Seleccionar carpeta de proceso", 0, "").Items.Item.path
  End With
: On Error GoTo 0
  If fSeleccionDirectorio = "" Then MsgBox "No se ha seleccionado carpeta; Verifique", , "Operación cancelada !!!"

End Function
Sub TimerProceso(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  Dim nControl_InputBox As Long

  ' Captura texto del control
  nControl_InputBox = FindWindowEx(FindWindow("#32770", "Clave de Acceso"), vbEmpty, "Edit", vbNullString)
  ' Establece PasswordChar
   SendMessageLongRef nControl_InputBox, &HCC&, Asc("*"), vbEmpty
  ' Finaliza Timer
  KillTimer hwnd, idEvent

End Sub
Function WriteImagen(ByVal o_rstImagen As ADODB.Recordset, ByVal o_Imagen As Object, ByVal s_Campo As String) As Boolean
  Dim o_Stream As ADODB.Stream
    
  ' Verifico parametros
  If Not (o_rstImagen Is Nothing) And s_Campo <> "" Then
    If o_rstImagen.State = adStateClosed Then GoTo ErrImagen
    On Error GoTo ErrImagen
    
    ' Inicializamos la imagen
    o_rstImagen.Fields(s_Campo).Value = Null
    ' Instancio los objetos y propiedades
    Set o_Stream = New ADODB.Stream
    ' Cargo la imagen
    If Not (o_Imagen.Picture = s_Estado_Ina) Then
      ' Guardo la imagen en un archivo temporal
      SavePicture o_Imagen.Picture, "imgtempo"
      ' Grabo la imagen en el registro de la tabla
      o_Stream.Type = adTypeBinary
      o_Stream.Open
      o_Stream.LoadFromFile "imgtempo"
      ' Insertamos la imagen
      o_rstImagen.Fields(s_Campo).Value = o_Stream.Read
      ' Cierro el objeto
      o_Stream.Close
      ' Elimino el archivo temporal
      If dir$("imgtempo", vbNormal) <> "" Then Kill "imgtempo"
    End If
    ' Actualizo la infortmación
    o_rstImagen.Update
  End If
  WriteImagen = True
  
ErrImagen:
  If Err.Number <> s_Estado_Ina Then: MsgBox Err.Description
  Set o_Stream = Nothing

End Function
