Attribute VB_Name = "modPrcFnc"
Option Explicit
Private Declare Function SystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Sub CaptionBotones(ByVal o_Form As Form, Optional ByVal lAceptar As Boolean, Optional ByVal lCancelar As Boolean, Optional ByVal lNuevo As Boolean, Optional ByVal lRevisar As Boolean, Optional ByVal lEliminar As Boolean, Optional ByVal lRefrescar As Boolean, Optional ByVal lVisualiza As Boolean, Optional ByVal lImprimir As Boolean, Optional lConfigura As Boolean, Optional ByVal lCorregir As Boolean, Optional ByVal lGrabar As Boolean, Optional ByVal lDeshacer As Boolean, Optional ByVal lSalir As Boolean, Optional ByVal aEtiqueta)
  
  Dim nContador As Integer
  
  '[ Cargo los mensajes de botones
  If lAceptar Then o_Form.cmdAceptar.Caption = Choose(gsIdioma, "&Aceptar", "&Accept")
  If lCancelar Then o_Form.cmdCancelar.Caption = Choose(gsIdioma, "&Cancelar", "&Cancel")
  If lRefrescar Then o_Form.cmdRefrescar.Caption = Choose(gsIdioma, "Re&frescar", "Re&fresh")
  If lSalir Then o_Form.cmdSalir.Caption = Choose(gsIdioma, "&Salir", "&Exit")
  
  If lNuevo Then o_Form.cmdNuevo.Caption = Choose(gsIdioma, "&Nuevo", "&New")
  If lRevisar Then o_Form.cmdRevisar.Caption = Choose(gsIdioma, "&Revisar", "&Review")
  If lEliminar Then o_Form.cmdEliminar.Caption = Choose(gsIdioma, "&Eliminar", "&Delete")
  If lVisualiza Then o_Form.cmdImprimir(0).Caption = Choose(gsIdioma, "&Preliminar", "&Preview")
  If lImprimir Then o_Form.cmdImprimir(1).Caption = Choose(gsIdioma, "&Imprimir", "Pr&int")
  If lConfigura Then o_Form.cmdConfig.Caption = Choose(gsIdioma, "&Configuración de Impresora", "Print &Setup")

  If lCorregir Then o_Form.cmdCorregir.Caption = Choose(gsIdioma, "&Corregir", "&Correct")
  If lGrabar Then o_Form.cmdGrabar.Caption = Choose(gsIdioma, "&Grabar", "&Save")
  If lDeshacer Then o_Form.cmdDeshacer.Caption = Choose(gsIdioma, "&Deshacer", "&Undo")
  '[ Cargo los mensajes de las etiquetas
  If UBound(aEtiqueta) > 0 Then
    For nContador = 0 To UBound(aEtiqueta, 1) - 1
      o_Form.lblTexto(nContador).Caption = aEtiqueta(nContador, CInt(gsIdioma) - 1)
    Next nContador
  End If
  
End Sub

'[ARREGLAR. Cambiar por un recordset fijo en Diario.
Public Function gfRetornaValor(cCadenaConec As String, cSource As String) As String
'cCadenaConec   Cadena de coneccion
'cSource        Sentecia de origen

   Static porstRetorno As ADODB.Recordset
   
   Set porstRetorno = New ADODB.Recordset
   
    gfRetornaValor = ""
    With porstRetorno
        .ActiveConnection = cCadenaConec
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = cSource
        .Open
    End With
    If Not (porstRetorno.EOF And porstRetorno.BOF) Then
        gfRetornaValor = porstRetorno(0)
    End If
    porstRetorno.Close
    Set porstRetorno = Nothing

End Function

']ARREGLAR.
Public Function gfCeros(tsCadena As String, _
                        tnTamaño As Integer, _
                        tnIncremento As Integer, _
                        tsCaracter As String)
'tsCadena         Cadena a procesar.
'tnTamaño         Tamaño resultante de la cadena.
'tnIncremento     Valor en que se incrementa la cadena.
'tsCaracter       Caracter que se añade a la izquierda de la cadena.

   If IsNumeric(tsCadena) Then tsCadena = Str(tsCadena)
   tsCadena = Trim(Str(Val(tsCadena) + tnIncremento))
   If tnTamaño > Len(tsCadena) Then
      tsCadena = String(tnTamaño - Len(tsCadena), tsCaracter) & tsCadena
   Else
      '[ARREGLAR: Qué hacer si tTamaño < Len(tCadena)
      ']ARREGLAR.
   End If
   
   gfCeros = tsCadena
End Function
Public Function gfDateText(ByVal dExpresion As Date) As String
  'dExpresion       Expresion de fecha
  Dim s_Expresion As String, s_DiaSemana As String, s_NombreMes As String
  Dim a_DiaSemana() As String
  ReDim a_DiaSemana(7, 2)
  
  a_DiaSemana(1, 1) = "Lunes": a_DiaSemana(1, 2) = "Monday"
  a_DiaSemana(2, 1) = "Martes": a_DiaSemana(2, 2) = "Tuesday"
  a_DiaSemana(3, 1) = "Miercoles": a_DiaSemana(3, 2) = "Wednesday"
  a_DiaSemana(4, 1) = "Jueves": a_DiaSemana(4, 2) = "Thurday"
  a_DiaSemana(5, 1) = "Viernes": a_DiaSemana(5, 2) = "Friday"
  a_DiaSemana(6, 1) = "Sabado": a_DiaSemana(6, 2) = "Saturday"
  a_DiaSemana(7, 1) = "Domingo": a_DiaSemana(7, 2) = "Sunday"
  
  s_NombreMes = gfMesLet(dExpresion, 0, "", 1, "", 0)
  s_DiaSemana = a_DiaSemana(DatePart("w", dExpresion, vbMonday, vbFirstFourDays), gsIdioma)
  If gsIdioma = "1" Then
    s_Expresion = s_DiaSemana & ", " & Format(dExpresion, "dd") & " de " & s_NombreMes & " de " & Format(dExpresion, "yyyy")
  Else
    s_Expresion = s_DiaSemana & Format(dExpresion, " dd") & ", " & s_NombreMes & Format(dExpresion, " yyyy")
  End If
   gfDateText = s_Expresion
End Function
Public Function gfEnmasc(tsCadena As String)
'tsCadena         Cadena a procesar.

   Dim dsCadInv As String, dsLista As String, _
       dnContador As Single, dnPosicion As Single
   
   If VarType(tsCadena) <> vbString Then Exit Function
   
  'Invierte la cadena.
   For dnContador = Len(tsCadena) To 1 Step -1
      dsCadInv = dsCadInv + Mid(tsCadena, dnContador, 1)
   Next
   
  'Genera la cadena enmascarada/desenmascarada.
   tsCadena = ""
   dsLista = " 192  96  96 128 -96 -96-192-128"
   For dnContador = 1 To Len(dsCadInv)
      dnPosicion = ((Int(Asc(Mid(dsCadInv, dnContador, 1)) / 32) + 1) * 4) - 4
      tsCadena = tsCadena + Chr(Asc(Mid(dsCadInv, dnContador, 1)) + Val(Mid(dsLista, IIf(dnPosicion = 1, 1, dnPosicion + 1), 4)))
   Next
   
   gfEnmasc = tsCadena
End Function

Public Function gfMesLet(tdFecha As Variant, tnDia As Integer, tsSepara1 As String, tnMes As Integer, tsSepara2 As String, tnAno As Integer)
'tdFecha          Fecha a procesar. Puede ser tipo Date o String. Si este último será como ddmmyyyy.
'tnDia            Formato como se mostrará el día.
'tnSepara1        Cadena que separa el día del mes.
'tnMes            Formato como se mostrará el mes.
'tnSepara2        Cadena que separa el mes del año.
'tnAno            Formato como se mostrará el año.

  Dim dsCadena As String
  Dim dsDia As String
  Dim dsMes As String
  Dim dsAno As String
  Dim dbTipoCadena As Boolean

  If VarType(tdFecha) <> vbDate And VarType(tdFecha) <> vbString Then
     gfMesLet = "ERR"
  End If

  dbTipoCadena = (VarType(tdFecha) = vbString)

  If tnDia <> 0 Then
     dsDia = IIf(dbTipoCadena, Left(tdFecha, 2), CStr(Day(tdFecha)))
     If tnDia = 1 And Left(dsDia, 1) = " " Then
         dsDia = "0" + LTrim(dsDia)
     ElseIf tnDia = 2 And Left(dsDia, 1) = " " Then
         dsDia = LTrim(dsDia)
     End If
  End If

  If tnMes <> 0 Then
     dsMes = IIf(dbTipoCadena, Mid(tdFecha, 3, 2), CStr(Month(tdFecha)))

     If tnMes < 3 Then
         dsCadena = Choose(gsIdioma, "00APERTURA  01Enero     02Febrero   03Marzo     04Abril     05Mayo      06Junio     07Julio     08Agosto    09Septiembre10Octubre   11Noviembre 12Diciembre 13CIERRE    14Enero     15Febrero   16Marzo     17Abril     ", _
                                     "00OPENING   01January   02February  03March     04Abril     05May       06June      07July      08August    09September 10October   11November  12December  13CLOSING   14January   15February  16March     17April     ")
         dsMes = IIf(Len(dsMes) = 1, "0" + dsMes, dsMes)
         dsMes = Mid(dsCadena, InStr(dsCadena, dsMes) + 2, 10)
         dsMes = IIf(tnMes = 2, Left(dsMes, 3), Trim(dsMes))
     ElseIf tnMes = 3 And Len(dsMes) = 1 Then
         dsMes = "0" + dsMes
     End If
  End If

  If tnAno <> 0 Then
     dsAno = Right(IIf(dbTipoCadena, Right(tdFecha, 4), CStr(Year(tdFecha))), IIf(tnAno = 1, 4, 2))
  End If

  gfMesLet = Trim(dsDia + tsSepara1 + dsMes + tsSepara2 + dsAno)
End Function
Function gfNumComprobante(ByVal s_Ano As String, ByVal s_Mes As String, ByVal s_Diario As String) As String
  
  ' s_Ano             Año donde  se genera
  ' s_Mes             Mes donde  se genera
  ' s_Diario          Copdigo de diario para generar numero
    
  Dim porstRetorno As ADODB.Recordset
  Dim s_Sentencia As String
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), '000000') AS cNumMaxCpb "
  s_Sentencia = s_Sentencia & "FROM CoCpbCab "
  s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND pdoano='" & s_Ano & "' "
  s_Sentencia = s_Sentencia & "AND MesPvs='" & s_Mes & "' "
  s_Sentencia = s_Sentencia & "AND CodDro='" & s_Diario & "'"
  Set porstRetorno = New ADODB.Recordset
  With porstRetorno
    .ActiveConnection = CONNSTRG & gsNomBDS
    '        .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With
  gfNumComprobante = gfCeros(porstRetorno!cNumMaxCpb, 6, 1, "0")
  porstRetorno.Close
  Set porstRetorno = Nothing

End Function
Public Function gfNumLet(tnNumero As Double, tsMoneda As String)
'tnNumero          Número a procesar.
'tsMoneda          Moneda en que se expresa el número (0:No Aplica N:MN Nuevos Soles E:ME Dólares Americanos).

'dsParteEntera     Parte decimal original.
'dsParteDecimal    Parte decimal resultante.
'dsUnidades1       Unidades del 0 al 9 en letras.
'dsUnidades2       Unidades del 10 al 19 en letras.
'dsDecenas         Decenas en letras.
'dsCentenas        Centenas en letras.
'dsLetras1         Expresión en letras de millares para abajo.
'dsLetras2         Expresión en letras de millones para arriba.
'daDigitos(1 To 15) Almacén de los dígitos del número a procesar.
'dnContador1       Contador.
'dnContador2       Contador.

   Dim dsParteEntera As String
   Dim dsParteDecimal As String
   Dim dsUnidades1 As String
   Dim dsUnidades2 As String
   Dim dsDecenas As String
   Dim dsCentenas As String
   Dim dsLetras1 As String
   Dim dsLetras2 As String
   Dim daDigitos(1 To 15) As String
   Dim dnContador1 As Integer
   Dim dnContador2 As Integer
   
   dsUnidades1 = "0      1UNO   2DOS   3TRES  4CUATRO5CINCO 6SEIS  7SIETE 8OCHO  9NUEVE "
   dsUnidades2 = "10DIEZ      11ONCE      12DOCE      13TRECE     14CATORCE   15QUINCE    16DIECISEIS 17DIECISIETE18DIECIOCHO 19DIECINUEVE"
   dsDecenas = "0         2VEINTE   3TREINTA  4CUARENTA 5CINCUENTA6SESENTA  7SETENTA  8OCHENTA  9NOVENTA   "
   dsCentenas = "0             1CIENTO       2DOSCIENTOS   3TRESCIENTOS  4CUATROCIENTOS5QUINIENTOS   6SEISCIENTOS  7SETECIENTOS  8OCHOCIENTOS  9NOVECIENTOS   "
   dsLetras1 = " "
   dsLetras2 = " "
 
   dsParteEntera = CStr(Int(tnNumero))
   dnContador1 = Len(dsParteEntera)
   dnContador2 = 0
   Do While dnContador1 > 0
      dnContador2 = dnContador2 + 1
      daDigitos(dnContador2) = Mid(dsParteEntera, dnContador1, 1)
      dnContador1 = dnContador1 - 1
   Loop

   dnContador2 = 1
   Do While dnContador2 <= Len(dsParteEntera)
      If daDigitos(dnContador2) <> "0" Then
         If dnContador2 = 1 Or dnContador2 = 4 Or dnContador2 = 7 Or dnContador2 = 10 Or dnContador2 = 13 Or dnContador2 = 16 Then
           'UNIDADES
            If dnContador2 <> 1 And daDigitos(dnContador2) = "1" Then
               dsLetras1 = "UN"
            Else
               dsLetras1 = Trim(Mid(dsUnidades1, InStr(dsUnidades1, daDigitos(dnContador2)) + 1, 6))
            End If

         ElseIf dnContador2 = 2 Or dnContador2 = 5 Or dnContador2 = 8 Or dnContador2 = 11 Or dnContador2 = 14 Then
           'DECENAS
            If daDigitos(dnContador2) = "1" Then
               dsLetras1 = Trim(Mid(dsUnidades2, InStr(dsUnidades2, daDigitos(dnContador2) + daDigitos(dnContador2 - 1)) + 2, 10))
            Else
               If daDigitos(dnContador2 - 1) = "0" Then
                  dsLetras1 = Trim(Mid(dsDecenas, InStr(dsDecenas, daDigitos(dnContador2)) + 1, 9))
               ElseIf daDigitos(dnContador2) <> "0" Then
                  dsLetras1 = Trim(Mid(dsDecenas, InStr(dsDecenas, daDigitos(dnContador2)) + 1, 9)) + " Y " + dsLetras1
               End If
            End If
        
         ElseIf dnContador2 = 3 Or dnContador2 = 6 Or dnContador2 = 9 Or dnContador2 = 12 Or dnContador2 = 15 Then
           'CENTENAS
            If daDigitos(dnContador2 - 1) = "0" And daDigitos(dnContador2 - 2) = "0" Then
               If daDigitos(dnContador2) = "1" Then
                  dsLetras1 = "CIEN"
               Else
                  dsLetras1 = Trim(Mid(dsCentenas, InStr(dsCentenas, daDigitos(dnContador2)) + 1, 13))
               End If
            Else
               dsLetras1 = Trim(Mid(dsCentenas, InStr(dsCentenas, daDigitos(dnContador2)) + 1, 13)) + " " + dsLetras1
            End If
         End If
      End If
      If (dnContador2 = 3 Or dnContador2 = 6 Or dnContador2 = 9 Or dnContador2 = 12 Or dnContador2 = 15 Or dnContador2 = 16 Or Len(dsParteEntera) = dnContador2) And dsLetras1 <> " " Then
         If Len(dsParteEntera) = dnContador2 Then
            dsLetras2 = dsLetras1 + IIf(dnContador2 = 4 Or dnContador2 = 5 Or dnContador2 = 6 Or dnContador2 = 10 Or dnContador2 = 11 Or dnContador2 = 12, " MIL ", IIf(dnContador2 = 7 Or dnContador2 = 8 Or dnContador2 = 9, IIf(dsLetras1 = "UN", " MILLON ", " MILLONES "), IIf(dnContador2 = 13 Or dnContador2 = 14 Or dnContador2 = 15, IIf(dsLetras1 = "UN", " BILLON ", " BILLONES "), ""))) + dsLetras2
         Else
            dsLetras2 = dsLetras1 + IIf(dnContador2 = 16 Or dnContador2 = 12 Or dnContador2 = 6, " MIL ", IIf(dnContador2 = 9, IIf(dsLetras1 = "UN", " MILLON ", " MILLONES "), IIf(dnContador2 = 15, IIf(dsLetras1 = "UN", " BILLON ", " BILLONES "), ""))) + dsLetras2
         End If
         dsLetras1 = " "
      End If
      dnContador2 = dnContador2 + 1
   Loop

   dsLetras2 = Trim(dsLetras2)
   dsParteDecimal = Mid(CDec(tnNumero) - Int(tnNumero), 3)
   If dsParteDecimal = "" Then
      dsParteDecimal = "00"
   ElseIf Len(dsParteDecimal) = 1 Then
      dsParteDecimal = dsParteDecimal + "0"
   ElseIf Len(dsParteDecimal) > 2 Then
      dsParteDecimal = Right(CStr(gfRedond(tnNumero, 2)), 2)
   End If
   
   gfNumLet = IIf(Len(dsLetras2) > 0, dsLetras2 + " Y ", "") + dsParteDecimal + "/100" + IIf(tsMoneda = "0", "", IIf(tsMoneda = TPOMON_NAC, " NUEVOS SOLES", " DOLARES AMERICANOS"))
End Function
Function gfPadC(ByVal Expresion, ByVal n_Longitud As Integer, ByVal s_Caracter As String) As String

Expresion = IIf(IsNull(Expresion) Or IsEmpty(Expresion), "", Expresion)
gfPadC = String$((n_Longitud - Len(Expresion)) / 2, s_Caracter) + Expresion + String$((n_Longitud - Len(Expresion)) / 2, s_Caracter)

End Function
Function gfPadL(ByVal Expresion, ByVal n_Longitud As Integer, ByVal s_Caracter As String) As String

Expresion = IIf(IsNull(Expresion) Or IsEmpty(Expresion), "", Expresion)
gfPadL = String$(n_Longitud - Len(Expresion), s_Caracter) & Expresion
    
End Function
Function gfPadR(ByVal Expresion, ByVal n_Longitud As Integer, ByVal s_Caracter As String) As String
Static s_Cadena As String

Expresion = IIf(IsNull(Expresion) Or IsEmpty(Expresion), "", Expresion)
If n_Longitud > Len(Expresion) Then
  s_Cadena = Expresion & String$(n_Longitud - Len(Expresion), s_Caracter)
Else
  s_Cadena = Expresion
End If
gfPadR = Left$(s_Cadena, n_Longitud)
    
End Function

Public Function gfParaOracle() As String
  
  Static s_Proveedor As String
  Static s_Servidor As String, s_BaseDatos As String
  Static s_UserId As String, s_Password As String

  Static s_Buffer As String, n_Size As Long
  Static s_Archivo As String, s_Linea As String
  Static o_fsoFileCfg As New FileSystemObject, o_fTexto As TextStream
  
  gfParaOracle = ""
  s_Proveedor = "OraOLEDB.Oracle.1"
  ' Reconoce el directorio de windows.
  s_Buffer = Space$(255)
  n_Size = Len(s_Buffer)
  SystemDirectory s_Buffer, n_Size
  s_Buffer = Left(s_Buffer, Len(Trim(s_Buffer)) - 1)

  ' Verifico que exista el Archivo de Configuracion
  s_Archivo = s_Buffer & "\" & pFileCfg
  If StrConv(Dir$(s_Archivo, vbHidden), vbLowerCase) <> LCase(pFileCfg) Then
    MsgBox TEXT_8002, vbCritical
    Exit Function
  End If

  ' Abro Archivo de Configuracion
  Set o_fTexto = o_fsoFileCfg.OpenTextFile(s_Archivo, ForReading)
  Do While Not o_fTexto.AtEndOfStream
    s_Linea = o_fTexto.ReadLine
    If Left$(s_Linea, 8) = "[Server]" Then s_Servidor = Mid$(s_Linea, InStr(s_Linea, "=") + 1)
    If Left$(s_Linea, 8) = "[UserId]" Then s_UserId = Mid$(s_Linea, InStr(s_Linea, "=") + 1)
    If Left$(s_Linea, 10) = "[Password]" Then s_Password = Mid$(s_Linea, InStr(s_Linea, "=") + 1)
    If Left$(s_Linea, 11) = "[BaseDatos]" Then s_BaseDatos = Mid$(s_Linea, InStr(s_Linea, "=") + 1)
  Loop
  o_fTexto.Close
  Set o_fsoFileCfg = Nothing
  Set o_fTexto = Nothing
  gfParaOracle = "provider=" & s_Proveedor & ";server=" & s_Servidor & ";Data Source=" & s_BaseDatos & ";user id=" & s_UserId & ";password=" & s_Password & ";Persist Security Info=True;"

End Function

Public Function gfRedond(tnNumero As Double, _
                         tnTotalDecimales As Integer) As Currency
'tnNumero          Número a procesar.
'tnTotalDecimales  Cantidad de decimales resultantes.

'dsParteDecimal1   Parte decimal original.
'dsParteDecimal2   Parte decimal resultante.

   Dim dsParteDecimal1 As String
   Dim dsParteDecimal2 As String

   dsParteDecimal1 = Mid(CStr(CDec(tnNumero) - Int(tnNumero)), 2)
   If Len(dsParteDecimal1) - 1 <= tnTotalDecimales Then
      gfRedond = tnNumero
      Exit Function
   Else
      dsParteDecimal1 = dsParteDecimal1 + String(9, "0")
      dsParteDecimal2 = Mid(dsParteDecimal1, 1, tnTotalDecimales + 1)
      If InStr("56789", Mid(dsParteDecimal1, tnTotalDecimales + 2, 1)) <> 0 Then
         If Mid(dsParteDecimal1, tnTotalDecimales + 1, 1) = "9" Then
            gfRedond = (Int(tnNumero) + Val(dsParteDecimal2) + Val("." + String(tnTotalDecimales - 1, "0") + "1"))
            Exit Function
         End If
         dsParteDecimal2 = Mid(dsParteDecimal1, 1, tnTotalDecimales) + CStr(Val(Mid(dsParteDecimal1, tnTotalDecimales + 1, 1)) + 1)
      End If
      gfRedond = (Int(tnNumero) + Val(dsParteDecimal2))
   End If
End Function

Function gfSacaEntRetApos(ByVal s_Expresion As String) As String

  s_Expresion = IIf(IsNull(s_Expresion) Or IsEmpty(s_Expresion), "", s_Expresion)
  If s_Expresion <> "" Then
    ' saco los Enters de la cadena de caracteres
    While InStr(s_Expresion, Chr(13)) <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(13)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(13)) + 1)
    Wend
    ' saco los Retornos de la cadena de caracteres
    While InStr(s_Expresion, Chr(10)) <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(10)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(10)) + 1)
    Wend
    ' saco los Apostrofes de la cadena de caracteres
    While InStr(s_Expresion, "'") <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "'") - 1) & "´" & Mid$(s_Expresion, InStr(s_Expresion, "'") + 1)
    Wend
    ' saco las Rayas de la cadena de caracteres
    While InStr(s_Expresion, "|") <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "|") - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, "|") + 1)
    Wend
  End If
  
''ini 2016-04-04 correccion teo ple diario, mayor
   'TC 29/03/2016
    While InStr(s_Expresion, "/") <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "/") - 1) & "-" & Mid$(s_Expresion, InStr(s_Expresion, "/") + 1)
    Wend
''fin 2016-04-04 correccion teo ple diario, mayor
  
  gfSacaEntRetApos = Trim$(s_Expresion)

End Function

Public Function gfSuma(tsCadena As String, _
                       tnCantidad As Integer)

'tsCadena         Cadena a procesar.
'tnCantidad       Cantidad a sumar. Para restar usar números negativos.

   Dim dsCadena1 As String             'Cadena.
   Dim dbTipoCadena As String             'Cadena.
   Dim dsCadena3 As String             'Cadena temporal.
   Dim dnInicio As Integer             'Primer dígito de la cadena.
   Dim dnFin As Integer                'Ultimo dígito de la cadena.
   Dim dbInicioFin As Boolean          'Para buscar el inicio/fin de la cadena.
   Dim dnContador As Integer           'Contador.

   dsCadena1 = tsCadena

   dnInicio = 0
   dnFin = Len(dsCadena1)
   dbInicioFin = True
   
   dnContador = 1
   Do While dnContador <= Len(dsCadena1)
      dsCadena3 = Mid(dsCadena1, dnContador, 1)
      If dbInicioFin Then
         If InStr("123456789", dsCadena3) <> 0 Then
            dnInicio = dnContador
            dnFin = Len(dsCadena1)
            dbInicioFin = False
         End If
      Else
         If InStr("0123456789", dsCadena3) = 0 Then
            dnFin = dnContador - 1
            dbInicioFin = True
         End If
      End If
      dnContador = dnContador + 1
   Loop
   If dnInicio = 0 Then
      dnInicio = dnFin
   End If
   
   dbTipoCadena = Mid(dsCadena1, dnInicio, dnFin + 1 - dnInicio)
   dbTipoCadena = RTrim(LTrim(CStr(Val(dbTipoCadena) + tnCantidad)))
   
   If dnInicio = 1 Then
      dsCadena1 = dbTipoCadena + Space(Len(dsCadena1) - Len(dbTipoCadena))
   Else
      If Len(dsCadena1) = Len(dbTipoCadena) Then
         dsCadena1 = dbTipoCadena
      Else
         dsCadena1 = Left(dsCadena1, dnInicio - IIf(Len(dbTipoCadena) > dnFin + 1 - dnInicio, 2, 1)) + dbTipoCadena
         If Len(dbTipoCadena) < Len(tsCadena) Then
            dsCadena1 = dsCadena1 + Space(Len(tsCadena) - Len(dsCadena1))
         End If
      End If
   End If
   
   gfSuma = dsCadena1
End Function

Public Function gfUltDia(tvFecha As Variant)
'tvFecha          Fecha a procesar.
'ddFecha          Utilizado durante el proceso.
   
   Dim ddFecha As Date
   
   ddFecha = CDate("28/" & Mid(CStr(tvFecha), 4)) + 5
   ddFecha = CDate("01/" & Mid(CStr(ddFecha), 4))
'   Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)
'   Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)

   gfUltDia = DateAdd("d", -1, ddFecha)
End Function
'ini 2015-08-27/09-02 ctr obligac sunat
'sirve para hallar el mes anterior sin importar el año donde este
Public Function gfMesAnte(tvFecha As Variant)
'tvFecha          Fecha a procesar.
'ddFecha          Utilizado durante el proceso.
   
   Dim ddFecha As Date
   
   'ddFecha = CDate("28/" & Mid(CStr(tvFecha), 4)) + 5
   ddFecha = CDate("01/" & Mid(CStr(tvFecha), 4))
'   Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)
'   Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)

   gfMesAnte = DateAdd("d", -1, ddFecha)
End Function
'fin 2015-08-27/09-02 ctr obligac sunat

'Public Function fValidaTxt(fContext As Object)
'   fValidaTxt = True
'   With fContext
'      If .Text = "" Then 'Necesario por por problemas con valores nulos.
'         .Text = " "
'      End If
''        If Not IsNumeric(.Text) And .Text <> "" Then
''            MsgBox "El Campo es un Valor Texto, Poner Datos Correctos"
''            .Text = ""
''            FValTxt = False
''        End If
'   End With
'End Function

'Función necesaria porque por problemas con valores nulos (Access).
'Public Function fValidaNum(fContext As Object)
'   fValidaNum = True
'   With fContext
'      If .Text = "" Then 'Necesario por por problemas con valores nulos.
'         .Text = "0"
'      End If
'      If Not IsNumeric(.Text) And .Text <> "" Then
'         MsgBox "El dato debe ser de tipo numérico. Corregir."
'         .Text = ""
'         fValidaNum = False
'      End If
'   End With
'End Function

'Public Function fValidaFec(fContext As Object, Optional fEnter As Integer)
'   fValidaFec = True
'   With fContext
'      If Not IsDate(.Text) And _
'         .Text <> "" Then
'         MsgBox "El dato debe ser de tipo fecha con formato dd/mm/aaaa. Corregir."
'         With frmUCogeFecha
'            .Show vbModal
'            fContext.Text = frmUCogeFecha.CbText(0) & "-" & frmUCogeFecha.CbText(1) & "-" & frmUCogeFecha.CbText(2)
'         End With
'         Unload frmUCogeFecha
''        Cancel = True
'         If fEnter <> 1 Then
'            SendKeys "{Tab}"
'         End If
'         fValidaFec = False
'      End If
''     If .Text = "" Then
''        'por estar la validacion en .t. en lugar de estar
''        'en Not FValDat()
''        FValDat = False
''     End If
'   End With
'End Function

'Public Sub gpTeclasGrid(tiKeyCode As Integer, tiShift As Integer, _
'                        toForm As Form, ParamArray taTeclas() As Variant)
'   Select Case tiKeyCode
'   Case vbKeyInsert
'      If taTeclas(0) Then
'         toForm.cmdNuevo_Click
'      End If
''   Case vbKeyEnter
''      If taTeclas(1) Then
''         cmdCorregir_Click
''      End If
'   Case vbKeyDelete
'      If taTeclas(2) Then
'         toForm.cmdEliminar_Click
'      End If
''   Case vbKey...
''      If taTeclas(3) Then
''        cmdImprimir_Click
''      End If
'   Case vbKeyF7
''      cmdImprimir_Click
'   Case vbKeyF8
''
'   End Select
'End Sub

Public Sub gpEncabezadoRpt(toRpt As CrystalReport, tsTit As String, tsFEm As Date, tbImpMesAno As Boolean, Optional ByVal tImpFecha As Boolean, Optional o_Data As Object)

  Dim dnContador As Integer
  Dim s_Fecha As String
   
  'Inicializa a "" todas las fórmulas, para no tener problemas de un reporte a otro. Esto es necesario por estar usando un único objeto rptMain de Crystal.
  For dnContador = 0 To 45: toRpt.Formulas(dnContador) = "":     toRpt.ParameterFields(dnContador) = "": Next dnContador
  s_Fecha = gfDateText(tsFEm) & " / " & Format(Time(), "hh:mm:ss AMPM")
  s_Fecha = IIf(tImpFecha, s_Fecha, "")
  With toRpt
    .Formulas(0) = "mSistema='" & gsNomSis & "'"
    .Formulas(1) = "mEmpresa='" & Trim(gsRazEmp) & "'"
    .Formulas(2) = "mTitulo='" & tsTit & "'"
    .WindowTitle = tsTit
    .Formulas(3) = "mFeReporte='" & s_Fecha & "'"
    ' .Formulas(4) = "mHrReporte='" & Format(Time(), "hh:mm:ss AMPM") & "'"
    If tbImpMesAno Then .Formulas(5) = "mPeriodo='" & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
    .Formulas(6) = "mRucEmpresa='" & Trim(gsRUCEmp) & "'"
    ' Inicializo las formulas de seleccion
    .SelectionFormula = ""
    .ParameterFields(0) = "Idioma;" & gsIdioma & ";true"
    ' ?.WindowShowCancelBtn = True
    ' ?.ProgressDialog = False 'True
    .WindowShowCloseBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = False
    .WindowShowSearchBtn = True
    .WindowShowZoomCtl = True
    .SetTablePrivateData 0, 3, o_Data
    .DiscardSavedData = False
 End With
End Sub

Public Sub gpEncabezadoRptPresup(toRpt As CrystalReport, tsTit As String, tsFEm As Date, tbImpMesAno As Boolean, Optional ByVal tImpFecha As Boolean, Optional o_Data As Object, Optional Presupuesto As String)
  Dim dnContador As Integer
  Dim s_Fecha As String
   
  'Inicializa a "" todas las fórmulas, para no tener problemas de un reporte a otro. Esto es necesario por estar usando un único objeto rptMain de Crystal.
  For dnContador = 0 To 45: toRpt.Formulas(dnContador) = "":     toRpt.ParameterFields(dnContador) = "": Next dnContador
  s_Fecha = gfDateText(tsFEm) & " / " & Format(Time(), "hh:mm:ss AMPM")
  s_Fecha = IIf(tImpFecha, s_Fecha, "")
  With toRpt
    .Formulas(0) = "mSistema='" & gsNomSis & "'"
    .Formulas(1) = "mEmpresa='" & Trim(gsRazEmp) & "'"
    .Formulas(2) = "mTitulo='" & tsTit & "'"
    .WindowTitle = tsTit
    .Formulas(3) = "mFeReporte='" & s_Fecha & "'"
    '      .Formulas(4) = "mHrReporte='" & Format(Time(), "hh:mm:ss AMPM") & "'"
    If tbImpMesAno Then .Formulas(5) = "mPeriodo='" & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
    .Formulas(6) = "mRucEmpresa='" & Trim(gsRUCEmp) & "'"
    ' Inicializo las formulas de seleccion
    .SelectionFormula = ""
    .ParameterFields(0) = "Idioma;" & gsIdioma & ";true"
    .ParameterFields(1) = "Presupuesto;" & Presupuesto & ";true"
    '?.WindowShowCancelBtn = True
    '?.ProgressDialog = False 'True
    .WindowShowCloseBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = False
    .WindowShowSearchBtn = True
    .WindowShowZoomCtl = True
    .SetTablePrivateData 0, 3, o_Data
    .DiscardSavedData = False
   End With

End Sub
Public Sub gpEncabezadoRptLibros(toRpt As CrystalReport, tsTit As String, tsFEm As Date, tbImpMesAno As Boolean, Optional ByVal tImpFecha As Boolean, Optional o_Data As Object)

  Dim dnContador As Integer
  Dim s_Fecha As String
   
  'Inicializa a "" todas las fórmulas, para no tener problemas de un reporte a otro. Esto es necesario por estar usando un único objeto rptMain de Crystal.
  For dnContador = 0 To 45: toRpt.Formulas(dnContador) = "":     toRpt.ParameterFields(dnContador) = "": Next dnContador
  s_Fecha = gfDateText(tsFEm) & " / " & Format(Time(), "hh:mm:ss AMPM")
  s_Fecha = IIf(tImpFecha, s_Fecha, "")
  With toRpt
    .Formulas(0) = "mSistema='" & gsNomSis & "'"
    .Formulas(1) = "mEmpresa='" & Trim(gsRazEmp) & "'"
    '.Formulas(2) = "mTitulo='" & rTitulo & "'"
    .Formulas(2) = "mTitulo='" & tsTit & "'"
    .WindowTitle = tsTit
    .Formulas(3) = "mFeReporte='" & s_Fecha & "'"
    ' .Formulas(4) = "mHrReporte='" & Format(Time(), "hh:mm:ss AMPM") & "'"
    If tbImpMesAno Then .Formulas(5) = "mPeriodo='" & gfMesLet("01" & xqmes & gsAnoAct, 0, "", 1, " ", 1) & "'"
    .Formulas(6) = "mRucEmpresa='" & Trim(gsRUCEmp) & "'"
    ' Inicializo las formulas de seleccion
    .SelectionFormula = ""
    .ParameterFields(0) = "Idioma;" & gsIdioma & ";true"
    ' ?.WindowShowCancelBtn = True
    ' ?.ProgressDialog = False 'True
    .WindowShowCloseBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = False
    .WindowShowSearchBtn = True
    .WindowShowZoomCtl = True
    .SetTablePrivateData 0, 3, o_Data
    .DiscardSavedData = False
   End With

End Sub
Public Sub gpComprimirArchivo(ByVal sArchivo As String)
  ' sArchivo directorio y archivo a comprimir
  Dim sArchivoZip As String, sComandoRun As String
  Dim nWindowStyle As Integer, nWaitOnReturn As Boolean
  Dim oWShell As Object
  
  sArchivoZip = Left(sArchivo, Len(sArchivo) - 3) & "zip"
  If Dir$(sArchivo, vbNormal) = "" Then
    MsgBox "El archivo TXT no existe en la ubicación especificada."
    Exit Sub
  End If
  
  ' Genera archivo comprimido
  If Dir$(sArchivoZip, vbNormal) <> "" Then Kill sArchivoZip
  sComandoRun = "Powershell -command Compress-Archive -Path '" & sArchivo & "' -DestinationPath '" & sArchivoZip & "'"
  Set oWShell = CreateObject("WScript.Shell")
  
  nWaitOnReturn = True
  nWindowStyle = 7
  oWShell.Run sComandoRun, nWindowStyle, nWaitOnReturn

  ' Saco de memoria
  Set oWShell = Nothing

End Sub
Public Sub gpEncabezadoMRp(toMRp As MRViewerObject, tsTit As String, tsFEm As Date, tbImpMesAno As Boolean, Optional ByVal tImpFecha As Boolean)
Dim s_Fecha As String, s_Hora As String
   
s_Fecha = gfDateText(tsFEm)
s_Fecha = IIf(tImpFecha, s_Fecha, "")
s_Hora = IIf(tImpFecha, Format(Time(), "hh:mm:ss AMPM"), "")
  
With toMRp
    .Parameters("mSistema") = gsNomSis
    'If tbImpMesAno Then .Parameters("mPeriodo") = Format(CDate(gfMesAct(gsMesAct) & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
    If tbImpMesAno Then .Parameters("mPeriodo") = gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1)
    .Parameters("mEmpresa") = Trim(gsRazEmp)
    .Parameters("mRucempresa") = Trim(gsRUCEmp)
    .Parameters("mTitulo") = UCase(tsTit)
    .Parameters("mFeReporte") = s_Fecha
    .Parameters("mHrReporte") = s_Hora
End With
  
End Sub

Public Sub gpErrores()
   MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Public Sub gpTeclasGrid(tiKeyCode As Integer, tiShift As Integer, _
                        toForm As Form, ParamArray taTeclas() As Variant)

'   Select Case tiKeyCode
'   Case vbKeyInsert
'      toForm.cmdNuevo_Click
'   Case vbKeyEnter
'      toForm.cmdRevisar_click
'   Case vbKeyDelete
'      toForm.cmdEliminar_Click
'   Case vbKeyF4
'      toForm.cmdRefrescar_Click
'   Case vbKeyF8
'      toForm.cmdImprimir_Click
'   End Select

End Sub

Public Sub gpTeclasData(tiKeyCode As Integer, tiShift As Integer, _
                        toForm As Form, ParamArray taTeclas() As Variant)
'   Select Case tiKeyCode
'   Case vbKeyEscape
'      toForm.cmdSalir_Click
'   Case vbKeyEnter
'      toForm.cmdCorregir_Click
'   Case vbKeyReturn
'      SendKeys "{TAB}"
'      tiKeyCode = 0
'   End Select
End Sub

Public Sub gpTUg_Resize(toFormGrid As Form)
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   toFormGrid.cmdSalir.Left = toFormGrid.Width - 820
   toFormGrid.fraBuscar.Width = toFormGrid.cmdSalir.Left - toFormGrid.fraBuscar.Left - 50
   toFormGrid.txtBuscar.Width = toFormGrid.fraBuscar.Width - 240
   toFormGrid.dgrMain.Height = toFormGrid.ScaleHeight - 30 - toFormGrid.picOpciones.Height '- uctEstado.Height
End Sub
'ini 2014-08-05 RR.HH afecto afp/onp
Public Sub gpTUg_Resize2(toFormGrid As Form)
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   toFormGrid.cmdSalir.Left = toFormGrid.Width - (820 + 0)
   toFormGrid.fraBuscar.Width = toFormGrid.cmdSalir.Left - toFormGrid.fraBuscar.Left - 50
   toFormGrid.txtBuscar.Width = toFormGrid.fraBuscar.Width - 240
   toFormGrid.dgrMain.Height = toFormGrid.ScaleHeight - 30 - toFormGrid.picOpciones.Height '- uctEstado.Height
End Sub
'fin 2014-08-05 RR.HH afecto afp/onp

Public Sub gpTUg_Nuevo(toFormGrid As Form, toForm As Form)
  On Error GoTo Err

   With toForm
    .zbNuevo = True   'Tiene que ir primero para que el load lo coja evaluado.
    .Caption = TEXT_NUEVO & " " & toFormGrid.Caption
    .upDatosPredeterminados
    .Show vbModal
   End With
   
  toFormGrid.dgrMain.SetFocus
  
  Exit Sub
Err:
  gpErrores
End Sub

Public Sub gpTUg_Refrescar(toFormGrid As Form)
   toFormGrid.uorstMain.Requery
   toFormGrid.ppDatosGrid
   
   toFormGrid.dgrMain.SetFocus
End Sub

Public Sub gpTUe_Retroceder(toRSet As Recordset, toForm As Form)
   With toRSet
      .MovePrevious
      If .BOF Then
         .MoveFirst
      End If
   End With
   toForm.upDatosDesconectados 1
End Sub

Public Sub gpTUe_Avanzar(toRSet As Recordset, toForm As Form)
   With toRSet
      .MoveNext
      If .EOF Then
         .MoveLast
      End If
   End With
   toForm.upDatosDesconectados 1
End Sub

Public Sub gpTUe_Deshacer(toForm As Form)
   toForm.upDatosDesconectados 1
   toForm.cmdRetroceder.Enabled = True
   toForm.cmdAvanzar.Enabled = True
   toForm.cmdCorregir.Enabled = True
   toForm.cmdGrabar.Enabled = False
   toForm.cmdDeshacer.Enabled = False
   toForm.upHabilitacion False
End Sub

Public Sub gpTVd_Nuevo(toFormGrid As Form, toForm As Form)
   On Error GoTo Err

   With toForm
      .zbNuevo = True   'Tiene que ir primero para que el load lo coja evaluado.
      .Caption = TEXT_NUEVO & " " & toFormGrid.Caption
      .upDatosPredeterminados
   
      .Show vbModal
   End With
   
   toFormGrid.dgrDetalle.SetFocus
  
   Exit Sub
Err:
   gpErrores
End Sub
Public Function ppNumeroLinea(ByVal s_Expresion As String) As Integer
  Dim nLen As Integer, nContador As Integer
  Dim nInicio As Integer, nFinal As Integer, nLongitud As Integer

  If s_Expresion <> "" Then
    ' saco los Enters de la cadena de caracteres
    While InStr(s_Expresion, Chr(13)) <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(13)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(13)) + 1)
    Wend
    ' saco los Retornos de la cadena de caracteres
    While InStr(s_Expresion, Chr(10)) <> 0
      nInicio = nFinal + 1
      nFinal = (InStr(nInicio, s_Expresion, Chr(10)))
      nLen = Len(Mid$(s_Expresion, nInicio, (nFinal - nInicio)))
      nContador = (nLen Mod 70)
      nContador = (nLen \ 70) + IIf(nContador > 0, 1, 0)
      ppNumeroLinea = ppNumeroLinea + nContador
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(10)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(10)) + 1)
    Wend
    ' Verifico el final de la linea
    nLongitud = Len(s_Expresion)
    If nLongitud > nFinal Then
      nInicio = nFinal + 1
      nLen = Len(Mid$(s_Expresion, nInicio, (nLongitud - nInicio)))
      nContador = (nLen Mod 70)
      nContador = (nLen \ 70) + IIf(nContador > 0, 1, 0)
      ppNumeroLinea = ppNumeroLinea + nContador
    End If
  End If

End Function
Public Sub ppSelRango(ByVal o_Grilla As Object, ByVal n_Indice As Integer, ByRef nInicio As Long, ByRef nFinal As Long, ByRef l_Seleccion As Boolean)
  Dim nFila As Long, nColumna As Integer
  
  On Error GoTo ErrRango
  
  Select Case n_Indice
   Case 0         ' Inicio de rango
    If nInicio = o_Grilla.Rows - 2 Then Exit Sub
    l_Seleccion = True
    For nColumna = 1 To o_Grilla.cols - 2
      o_Grilla.Col = nColumna
      o_Grilla.CellBackColor = &H8000000D
      o_Grilla.CellForeColor = &H80000018
    Next nColumna
    o_Grilla.row = nInicio: o_Grilla.Col = 1
   Case 1         ' Final de rango
    If nInicio < 1 Then
      MsgBox Choose(gsIdioma, "No Selecciono el Registro Inicial; Verifique", "Did not Select Initial Record; Verify"), vbExclamation
      o_Grilla.SetFocus
      Exit Sub
    End If
    If nFinal = o_Grilla.Rows - 1 Then Exit Sub
    
    For nFila = nInicio To nFinal
      o_Grilla.row = nFila
      For nColumna = 1 To o_Grilla.cols - 1
        o_Grilla.Col = nColumna
        o_Grilla.CellBackColor = &H8000000D
        o_Grilla.CellForeColor = &H80000018
      Next nColumna
    Next nFila
    o_Grilla.row = nFinal: o_Grilla.Col = 1
    nFinal = 0: nInicio = 0
   Case 2         ' Inicializa rango
    l_Seleccion = False
    For nFila = 1 To o_Grilla.Rows - 2
      o_Grilla.row = nFila
      For nColumna = 1 To o_Grilla.cols - 1
        o_Grilla.Col = nColumna
        o_Grilla.CellBackColor = &H80000018
        o_Grilla.CellForeColor = QBColor(0)
      Next nColumna
    Next nFila
    o_Grilla.row = 1: o_Grilla.Col = 1
  End Select
  Exit Sub

ErrRango:
  MsgBox Err.Description

End Sub
Public Function ValidadConexion(ByVal s_DataBase As String) As Boolean
  Dim pocnnValida As ADODB.Connection
  
  On Error GoTo CapturaError
    
  ValidadConexion = False
  ' Realizo la Conexión
  Set pocnnValida = New ADODB.Connection
  With pocnnValida
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & s_DataBase
    .Open
  End With
  ValidadConexion = True
  pocnnValida.Close
  
CapturaError:
   Set pocnnValida = Nothing
  
End Function
Public Function ValidaEjercicio(ByVal s_Empresa As String, ByVal s_Ejercicio As String) As Boolean
  Dim porstValidar As ADODB.Recordset
  Dim psSentencia As String
  
  ' Verifico y genero la base de datos
  psSentencia = "SELECT COUNT(a.codemp) AS nRegistro "
  psSentencia = psSentencia & "FROM cocfg a, tgcfg b "
  psSentencia = psSentencia & "WHERE a.codemp='" & s_Empresa & "' "
  psSentencia = psSentencia & "AND a.pdoano='" & s_Ejercicio & "' "
  psSentencia = psSentencia & "AND b.codemp=a.codemp "
  psSentencia = psSentencia & "AND b.pdoano=a.pdoano"
  Set porstValidar = New ADODB.Recordset
  With porstValidar
    .ActiveConnection = CONNSTRG & gsNomBDS
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = psSentencia
    .Open
  End With
  ValidaEjercicio = (porstValidar!nRegistro > 0)
  porstValidar.Close
  Set porstValidar = Nothing

End Function
