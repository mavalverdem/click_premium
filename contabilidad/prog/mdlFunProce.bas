Attribute VB_Name = "mdlFunProce"
'2015-01-07 funciones que viene del sistema de planilla
'2015-04-06 tambien funciones creadas por rcs
Option Explicit
Sub ReadImagen(ByVal o_rstImagen As ADODB.Recordset, ByVal o_Imagen As Object, ByVal s_Campo As String)
'ini 2015-01-07 adiciono imagen empresa
Dim s_Estado_Ina As Integer
s_Estado_Ina = 0
'fin 2015-01-07 adiciono imagen empresa

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
    If Dir$("imgtempo", vbNormal) <> "" Then Kill "imgtempo"
  End If
  o_Imagen.Refresh
  
ErrImagen:
  If Err.Number <> s_Estado_Ina Then: MsgBox Err.Description
  '2015-01-07 adiciono imagen empresa MsgBox Err.Number & " / " & Err.Description
  Set o_Stream = Nothing

End Sub

'2015-04-06 funcion que sirve para verficar si existe campo en el Recordset
Public Function ExistFieldInRS(oRs As ADODB.Recordset, _
sFieldName As String) As Boolean
On Error GoTo ErrEF
Dim x
x = oRs(sFieldName)
ExistFieldInRS = True
Exit Function
ErrEF:
End Function

'2015-06-05 estado de la mayorizacion
Public Function fEstMaySeek() As Integer
'fEstMaySeek= 0, no existe registro,
'fEstMaySeek= 1, existe registros mayorizar
  On Error GoTo ErrorRs
  Dim sEstado As Integer
  sEstado = 0
  Dim pCnn As ADODB.Connection
  Set pCnn = New ADODB.Connection
     
  With pCnn
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
   
  Dim pRst As ADODB.Recordset
  Set pRst = New ADODB.Recordset
  
  
  Dim x_Sentencia As String
  x_Sentencia = "SELECT * FROM CoCieMes "
  x_Sentencia = x_Sentencia & "WHERE "
  x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
  x_Sentencia = x_Sentencia & "AND indProcMay=1 "
  x_Sentencia = x_Sentencia & "ORDER BY pdoano, mescie" 'ini 2015-08-25 teo mesas anteriores al control mes proc dejar utilizar opcion
  'x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
  'x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
  'ORDER BY pdoano, mescie
  With pRst
    .ActiveConnection = pCnn
    .Source = x_Sentencia
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    '.MoveFirst 2015-09-01 error eof cuando no existe registro
    If .RecordCount > 0 Then .MoveFirst
    If Not .EOF Then
      'ini 2015-08-25 teo mesas anteriores al control mes proc dejar utilizar opcion
      'sEstado = 1
      If gsAnoAct & gsMesAct >= .Fields("pdoano") & .Fields("mescie") Then
        sEstado = 1
      End If
      'fin 2015-08-25 teo mesas anteriores al control mes proc dejar utilizar opcion
    End If
  End With
  
  If Not pRst Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
    'If prst.State = adStateOpen Then prst.Close: Set prst = Nothing
    If pRst.State = adStateOpen Then pRst.Close
    Set pRst = Nothing
  End If
  
  If Not pCnn Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
    If pCnn.State = adStateOpen Then pCnn.Close
    Set pCnn = Nothing
  End If
  
  
  fEstMaySeek = sEstado
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Function


'2015-09-25 se necesita que reconozca el mes diferente al actual
'Optional pmes As String
Public Sub fEstMayUpd(Optional pindProcMay As Integer, Optional pmes As String)
'pindProcMay= 0, graba 1 en tabla que significa "peridos por procesar mayorizacion"
'pindProcMay= -1, graba 0 en tabla que significa "peridos ya fue procesado"
  On Error GoTo ErrorRs
    If pindProcMay = 0 Then
      pindProcMay = 1
    End If
    If pindProcMay = -1 Then
      pindProcMay = 0
    End If
    Dim pCnn As ADODB.Connection
    Set pCnn = New ADODB.Connection
    
   With pCnn
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
    
    
    Dim x_Sentencia As String
    x_Sentencia = "UPDATE CoCieMes SET indProcMay=" & Str(pindProcMay) & " "
    x_Sentencia = x_Sentencia & "WHERE "
    x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
    x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
'ini 2015-09-25 se necesita que reconozca el mes diferente al actual
'    x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
    If Len(Trim(pmes)) = 0 Then
        x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
    Else
        x_Sentencia = x_Sentencia & "AND MesCie='" & pmes & "' "
    End If
'fin 2015-09-25 se necesita que reconozca el mes diferente al actual
    
    With pCnn
        .Execute x_Sentencia
    End With
    
    If Not pCnn Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
      If pCnn.State = adStateOpen Then pCnn.Close
      Set pCnn = Nothing
    End If
    
    'fEstMayUpd = 0
  Exit Sub
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub

'se necesita que reconozca el mes diferente al actual
Public Sub fEstMayUpd_2015_09_25(Optional pindProcMay As Integer)
'pindProcMay= 0, graba 1 en tabla que significa "peridos por procesar mayorizacion"
'pindProcMay= -1, graba 0 en tabla que significa "peridos ya fue procesado"
  On Error GoTo ErrorRs
    If pindProcMay = 0 Then
      pindProcMay = 1
    End If
    If pindProcMay = -1 Then
      pindProcMay = 0
    End If
    Dim pCnn As ADODB.Connection
    Set pCnn = New ADODB.Connection
    
   With pCnn
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
    
    
    Dim x_Sentencia As String
    x_Sentencia = "UPDATE CoCieMes SET indProcMay=" & Str(pindProcMay) & " "
    x_Sentencia = x_Sentencia & "WHERE "
    x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
    x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
    x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
    
    With pCnn
        .Execute x_Sentencia
    End With
    
    If Not pCnn Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
      If pCnn.State = adStateOpen Then pCnn.Close
      Set pCnn = Nothing
    End If
    
    'fEstMayUpd = 0
  Exit Sub
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub



''Public Sub fEstMayUpd(pCnn As ADODB.Connection, Optional pindProcMay As Integer)
'''pindProcMay= 0, graba 1 en tabla que significa "peridos por procesar mayorizacion"
'''pindProcMay= -1, graba 0 en tabla que significa "peridos ya fue procesado"
''  On Error GoTo ErrorRs
''    If pindProcMay = 0 Then
''      pindProcMay = 1
''    End If
''    If pindProcMay = -1 Then
''      pindProcMay = 0
''    End If
''    'Set pCnn = New ADODB.Connection
''    Dim x_Sentencia As String
''    x_Sentencia = "UPDATE CoCieMes SET indProcMay=" & Str(pindProcMay) & " "
''    x_Sentencia = x_Sentencia & "WHERE "
''    x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
''    x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
''    x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
''
''    With pCnn
''        .Execute x_Sentencia
''    End With
''
''    If Not pCnn Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
''      If pCnn.State = adStateOpen Then pCnn.Close
''      Set pCnn = Nothing
''    End If
''
''    'fEstMayUpd = 0
''  Exit Sub
''ErrorRs:
''  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
''
''End Sub


''Public Function fEstMayUpd(pCnn As ADODB.Connection, Optional pindProcMay As Integer) As Integer
'''pindProcMay= 0, graba 1 en tabla que significa "peridos por procesar mayorizacion"
'''pindProcMay= -1, graba 0 en tabla que significa "peridos ya fue procesado"
''  On Error GoTo ErrorRs
''    If pindProcMay = 0 Then
''      pindProcMay = 1
''    End If
''    If pindProcMay = -1 Then
''      pindProcMay = 0
''    End If
''    'Set pCnn = New ADODB.Connection
''    Dim x_Sentencia As String
''    x_Sentencia = "UPDATE CoCieMes SET indProcMay=" & Str(pindProcMay) & " "
''    x_Sentencia = x_Sentencia & "WHERE "
''    x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
''    x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
''    x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
''
''    With pCnn
''        .Execute x_Sentencia
''    End With
''
''    If Not pCnn Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
''      If pCnn.State = adStateOpen Then pCnn.Close
''      Set pCnn = Nothing
''    End If
''
''    fEstMayUpd = 0
''  Exit Function
''ErrorRs:
''  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
''End Function
''
'Public Function fTsaDetrac(pRst As ADODB.Recordset, sServicio As String) As Double
Public Function fTsaDetrac(pCnn As ADODB.Connection, sServicio As String) As Double
    Dim nDetraccion As Double
    nDetraccion = 0
    Dim pRst As ADODB.Recordset
    Set pRst = New ADODB.Recordset
    Set pRst = fRstDetrac(pCnn, pRst)
    With pRst
        If .RecordCount > 0 Then .MoveFirst
            .Find "coddetrac3='" & sServicio & "'"
            If Not .EOF Then
                '2015-07-08 cambio de decima a % nDetraccion = !pctdetrac
                nDetraccion = fDiv0(!pctdetrac, 100)
            End If
    End With
    pRst.Close
    Set pRst = Nothing
    fTsaDetrac = nDetraccion
End Function

Public Function fRstDetrac(pCnn As ADODB.Connection, pRst As ADODB.Recordset) As ADODB.Recordset
  On Error GoTo ErrorRs
  
    'Set pRst = New ADODB.Recordset
    With pRst '
           .ActiveConnection = pCnn
           .Source = "SELECT coddetrac, " & Choose(gsIdioma, "detdetrac", "detdetracx") & " AS DetDetrac,pctdetrac ,  "
           .Source = .Source & "left(coddetrac,3) coddetrac3, "
           .Source = .Source & "codemp "
           .Source = .Source & "FROM codetrac  "
           .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
           .Source = .Source & "AND estdetrac ='" & ESTDETRAC_ACT & "' "
           '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
           '.Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
           .CursorType = adOpenDynamic
           .LockType = adLockOptimistic
           .Open
    End With
    Set fRstDetrac = pRst
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Function

'2015-07-08 validar en el texbox que solo ingrese numeros enteros
Public Function fValidInt(Index As Integer, KeyAscii As Integer, pLabel As String) As Integer
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
            ' El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace
            'Exit Function
        Else
            KeyAscii = 0
            MsgBox "Solo números para registrar " & pLabel _
            & " sin puntos, " & "ni comas, ni cualquier caracter especial!!"
        End If

    fValidInt = KeyAscii
End Function

'2015-07-08 validar en el texbox que solo ingrese numeros decimales
Public Function fValidDeci(Index As Integer, KeyAscii As Integer, pLabel As String) As Integer
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46 Then
            ' El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace, 46 punto decimal
            'Exit Function
        Else
            KeyAscii = 0
            MsgBox "Solo números para registrar " & pLabel _
            & " sin comas, ni cualquier caracter especial!!"
        End If

    fValidDeci = KeyAscii
End Function


'ini 2015-06-23 control flag mayoriza
'0=mes abierto 1=mes cerrado
Function gcCierre(p_anno As String, p_mes As String) As Integer
        Dim sxMes As String
        sxMes = p_mes
        Dim arr_cie() As Integer
        Dim i As Integer
        For i = 1 To Int(sxMes)
        'Verificación de Mes Cerrado.
           sxMes = Format(i, "00")
           'Dim arr_cie() As Long
           arr_cie() = gpCieMes_arr(p_anno, sxMes) 'verificar todos los periodos
           If arr_cie(4) = 1 Then
              MsgBox TEXT_9016 & " " & sxMes & "/" & gsAnoAct, vbCritical
              gcCierre = 1
              Exit Function
           End If
        Next
    
        gcCierre = 0
End Function

'envia estados de cierre en array 1...5
Function gpCieMes_arr(p_anno As String, p_mes As String) As Integer() 'As Long() '
    Dim arr_cie(6) As Integer
    'Dim arr_cie(6) As Long
   Dim docnnMain As ADODB.Connection
   Dim dorstMain As ADODB.Recordset

   Set docnnMain = New ADODB.Connection
   Set dorstMain = New ADODB.Recordset
   With docnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
   End With
   With dorstMain
    .ActiveConnection = docnnMain
    .Source = "SELECT IndCpr, IndVta, IndHpr, IndCpb,IndProcMay "
    .Source = .Source & "FROM COCieMes "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & p_anno & "' "
    .Source = .Source & "AND mescie='" & p_mes & "'"
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
   
    arr_cie(1) = !IndCpr
    arr_cie(2) = !IndVta
    arr_cie(3) = !IndHpr
    arr_cie(4) = !IndCpb
    arr_cie(5) = !IndProcMay
    .Close
   End With

   docnnMain.Close
   Set dorstMain = Nothing
   Set docnnMain = Nothing
   gpCieMes_arr = arr_cie
End Function
'fin 2015-06-23 control flag mayoriza

Function gfMesAct(tsMesAct As String)
'tsFecha          Mes a procesar.
   
   gfMesAct = IIf(gsMesAct = "00", gsMesApe, IIf(gsMesAct = "13", gsMesCie, gsMesAct))
End Function

'ini 2015-08-27 ctr obligac sunat
Sub gfctrl_obliga_sunat()
    '2015-10-15 solo el periodo actu segun now If gsAnoAct <= "2013" Then
    If Val(gsAnoAct) <> Year(Now) Then
        Exit Sub
    End If
'ini 2015-08-27/09-02 ctr obligac sunat
    Dim xx_ano_act As String
    Dim xx_mes_act As String
    Dim xx_fe_ante As Date
    xx_fe_ante = gfMesAnte(CDate("01/" + Format(Month(Now), "00") + "/" + Format(Year(Now), "0000")))
    'gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
'   xx_ano_act = Format(xx_fe_ante, "0000")
'    xx_mes_act = Format(xx_fe_ante, "00")
    xx_ano_act = Format(Year(xx_fe_ante), "0000")
    xx_mes_act = Format(Month(xx_fe_ante), "00")
'fin 2015-08-27/09-02 ctr obligac sunat

    Dim docnnMain As ADODB.Connection
    Dim dorstMain As ADODB.Recordset
    Dim dorstMain2 As ADODB.Recordset
    
    Set docnnMain = New ADODB.Connection
    Set dorstMain = New ADODB.Recordset
    Set dorstMain2 = New ADODB.Recordset
   
    Dim x_Sentencia As String
    x_Sentencia = "SELECT * "
    'x_Sentencia = x_Sentencia & "FROM tgCtrObli "
    x_Sentencia = x_Sentencia & "FROM " & gsNomBDC & fPunto & "tgCtrObli " '2015-08-27/09-02 ctr obligac sunat
   'x_Sentencia = x_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
    x_Sentencia = x_Sentencia & "WHERE pdotribu='" & xx_ano_act & xx_mes_act & "' " '2015-08-27/09-02 ctr obligac sunat
    'x_Sentencia = x_Sentencia & "WHERE pdotribu='" & gsAnoAct & gsMesAct & "' "
    'x_Sentencia = x_Sentencia & "AND mescie='" & gsMesAct & "'"
    
    Set docnnMain = fCnnOpen(docnnMain)
    Set dorstMain = fRstOpenBuscar(docnnMain, dorstMain, x_Sentencia)
    
    'gsRUCEmp
    With dorstMain
       'si no existe registro crea todo el periodo de obligacion
        If .EOF Then
'ini 2015-12-17 error sin no tiene internet
            'gfctrl_obliga_sunat_html
            If gfctrl_obliga_sunat_html() = 1 Then
                MsgBox ("Sin acceso a interne, no se puede validar fecha declaracion")
                Exit Sub
            End If
'fin 2015-12-17 error sin no tiene internet
            Set dorstMain = fRstOpenBuscar(docnnMain, dorstMain, x_Sentencia)
        End If
        Dim xruc_fec_vence As Date
        xruc_fec_vence = dorstMain.Fields("fecVence" & Right(gsRUCEmp, 1))
        If gsBuenContriEmp = ESTDBUEN_CONTRI_ACT Then
            xruc_fec_vence = dorstMain.Fields("buencontri")
        End If
    End With
        
        x_Sentencia = "SELECT * "
        x_Sentencia = x_Sentencia & "FROM " & gsNomBDC & fPunto & "tgCtrObliDet "
        x_Sentencia = x_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
        x_Sentencia = x_Sentencia & "AND pdotribu='" & xx_ano_act & xx_mes_act & "' " '2015-08-27/09-02 ctr obligac sunat
        'x_Sentencia = x_Sentencia & "AND pdotribu='" & gsAnoAct & gsMesAct & "' "
        'x_Sentencia = x_Sentencia & "AND mescie='" & gsMesAct & "'"
        
        Set dorstMain2 = fRstOpenBuscar(docnnMain, dorstMain, x_Sentencia)
        With dorstMain2
            'presento formulario solo si no existe y
            'la fecha es menor a igual al de la pc now-actual
            If .EOF Then
                'If xruc_fec_vence <= Now Then
                'If xruc_fec_vence <= Format(Now, "dd/mm/yyyy") Then
                If Format(Now, "dd/mm/yyyy") <= xruc_fec_vence Then   '2015-08-27/09-02 ctr obligac sunat
                    If MsgBox(TEXT_9024 & Format(xruc_fec_vence, "dd/mm/yyyy") & TEXT_9025 & ")", vbYesNo + vbQuestion + vbDefaultButton2, "Pregunta") = vbYes Then
                     frmZCtrObliSunat.Show vbModal
                    End If
                End If
            End If
        End With
        
    fRstClose dorstMain
    fRstClose dorstMain2
    'Set dorstMain = Nothing
    'Set dorstMain2 = Nothing
    fCnnClose docnnMain
    'Set docnnMain = Nothing
End Sub

'2015-12-17 error sin no tiene internet Sub gfctrl_obliga_sunat_html()
Function gfctrl_obliga_sunat_html() As Integer
  On Error GoTo ErrorRs
 '---------------------------
    Dim IE As InternetExplorer
    Dim HTMLdoc As HTMLDocument
    Dim HTMLdoc2 As HTMLDocument '2015-07-21 creado por rcs
    
    Dim TDelements As IHTMLElementCollection
    Dim TDelement As HTMLTableCell
    Dim url As String
    Dim mtext1 As String
    mtext1 = "http://www.sunat.gob.pe/cl-ti-itcronobligme/Opciones.jsp?per=" & gsAnoAct
    'url = Trim(Text1.Text)
    url = Trim(mtext1)
    Set IE = New InternetExplorer
    
    With IE
        .navigate url
        .Visible = False
        'Esperamos que toda la web cargue
        While .Busy Or .readyState <> READYSTATE_COMPLETE: DoEvents: Wend
        Set HTMLdoc = .document
    End With
    With HTMLdoc.selectForm
        .accion.Value = "rptGral"
        .submit

        '            .mes.selectedIndex = Tc_Mes
        '            .anho.Value = tc_año
        '            .submit
    End With
    Set HTMLdoc2 = IE.document
    While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE: DoEvents: Wend
    Dim mtext2 As String
    mtext2 = "p"
    Set TDelements = HTMLdoc2.getElementsByTagName(Trim(mtext2))
    'Set TDelements = HTMLdoc2.getElementsByTagName(Trim(Text2.Text))
    Dim xite1 As Integer
    Dim xite2 As Integer
    xite1 = 1
    Dim xinsert As String
    xinsert = ""
    'xite2 = 1 'es para realizar pausa cada 20 lineas
    Dim xannio_tribu As String
    xannio_tribu = ""
    Dim xmes_tribu As String
    xmes_tribu = ""
    For Each TDelement In TDelements
        'Debug.Print "<<*" & Trim(Str(xite1)) & "*>> <" & TDelement.ClassName & "> - " & TDelement.innerText
        
'        xite1 = xite1 + 1
        xmes_tribu = gfctrl_obliga_sunat_mes(Right(TDelement.innerText, 3))
        Select Case xite1
        Case 15, 27, 39, 51, 63, 75, 87, 99, 111, 123, 135, 147
            'xinsert = "INSERT INTO tgCtrObli"
            'xinsert = "'" & gsAnoAct & gsMesAct & "','" & TDelement.innerText & "'"
             xmes_tribu = gfctrl_obliga_sunat_mes(Left(TDelement.innerText, 3))
             xannio_tribu = Trim(Str(Val(Right(TDelement.innerText, 2)) + 2000))
            'xinsert = "'" & xannio_tribu & gfctrl_obliga_sunat_mes(Left(TDelement.innerText, 3)) & "'"
            
            xinsert = "'" & xannio_tribu & xmes_tribu & "'"
            'xinsert = "'" & gsAnoAct & gsMesAct & "'"
        Case 15 + 11, 27 + 11, 39 + 11, 51 + 11, 63 + 11, 75 + 11, 87 + 11, 99 + 11, 111 + 11, 123 + 11, 135 + 11, 147 + 11
            'xinsert = xinsert & ",'" & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            
            'xinsert = xinsert & ",'" & xannio_tribu & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            'periodos de enero a diciembre
            If xite1 >= 15 And xite1 <= 146 Then
            'xinsert = xinsert & ",'" & xannio_tribu & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            'fDateFmt
            'xinsert = xinsert & ",'" & CDate(Left(TDelement.innerText, 2) & "/" & xmes_tribu & "/" & xannio_tribu) & "'"
            xinsert = xinsert & "," & fDateFmt(xannio_tribu & "-" & xmes_tribu & "-" & Left(TDelement.innerText, 2)) & " "
            End If
            'periodo solo de diciembre, para poner siguiente año
            If xite1 >= 147 And xite1 <= 158 Then
             'xinsert = xinsert & ",'" & Str(Val(xannio_tribu) + 1) & xmes_tribu & Left(TDelement.innerText, 2) & "'"
             'xinsert = xinsert & ",'" & CDate(Left(TDelement.innerText, 2) & "/" & xmes_tribu & "/" & Str(Val(xannio_tribu) + 1)) & "'"
             xinsert = xinsert & "," & fDateFmt(Str(Val(xannio_tribu) + 1) & "-" & xmes_tribu & "-" & Left(TDelement.innerText, 2)) & " "
            End If
          gfctrl_obliga_sunat_insert xinsert
       Case Else
            'If xite1 >= 15 And xite1 <= 158 Then
            'periodos de enero a diciembre
            If xite1 >= 15 And xite1 <= 146 Then
            'xinsert = xinsert & ",'" & xannio_tribu & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            'xinsert = xinsert & ",'" & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            'xinsert = xinsert & ",'" & CDate(Left(TDelement.innerText, 2) & "/" & xmes_tribu & "/" & xannio_tribu) & "'"
            xinsert = xinsert & "," & fDateFmt(xannio_tribu & "-" & xmes_tribu & "-" & Left(TDelement.innerText, 2)) & " "
           End If
            'periodo solo de diciembre, para poner siguiente año
            If xite1 >= 147 And xite1 <= 158 Then
            'xannio_tribu = Str(Val(xannio_tribu) + 1)
            'xinsert = xinsert & ",'" & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            'xinsert = xinsert & ",'" & Str(Val(xannio_tribu) + 1) & xmes_tribu & Left(TDelement.innerText, 2) & "'"
            'xinsert = xinsert & ",'" & CDate(Left(TDelement.innerText, 2) & "/" & xmes_tribu & "/" & Str(Val(xannio_tribu) + 1)) & "'"
             xinsert = xinsert & "," & fDateFmt(Str(Val(xannio_tribu) + 1) & "-" & xmes_tribu & "-" & Left(TDelement.innerText, 2)) & " "
          End If
        End Select

        xite1 = xite1 + 1
'        Select Case xite1
'        Case 15 To 26 'ene
'        Case 27 To 38 'feb
'        Case 39 To 50 'mar
'        Case 51 To 62 'abri
'        Case 63 To 74 'may
'        Case 75 To 86 'jun
'        Case 87 To 98 'jul
'        Case 99 To 110 'ago
'        Case 111 To 122 'set
'        Case 123 To 134 'oct
'        Case 135 To 146 'nov
'        Case 147 To 158 'dic
'        End Select
        
        
        'Debug.Print "<<*" & Trim(Str(xite1)) & "*>> <" & TDelement.ClassName & "> - " & TDelement.innerText
        'xite2 = xite2 + 1
        'If xite2 = 20 Then
        '    xite2 = 0
        'End If
    Next
    IE.Quit
  
  gfctrl_obliga_sunat_html = 0 '2015-12-17 error sin no tiene internet
 '---------------------------
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  gfctrl_obliga_sunat_html = 1 '2015-12-17 error sin no tiene internet

End Function

Sub gfctrl_obliga_sunat_insert(pvalues As String)
'Function gfctrl_obliga_sunat_insert(pvalues As String) As Integer
   Dim docnnMain As ADODB.Connection
   Set docnnMain = New ADODB.Connection
   
    Dim x_Sentencia As String
    x_Sentencia = "INSERT INTO " & gsNomBDC & fPunto & "tgCtrObli ("
    x_Sentencia = x_Sentencia & "pdotribu, fecVence0,  fecVence1,  fecVence2,  fecVence3,"
    x_Sentencia = x_Sentencia & "fecVence4,fecVence5,  fecVence6,  fecVence7,  fecVence8,"
    x_Sentencia = x_Sentencia & "fecVence9 , buencontri, usrcre, fyhcre"
    x_Sentencia = x_Sentencia & ")"
    x_Sentencia = x_Sentencia & " VALUES ("
    x_Sentencia = x_Sentencia & pvalues
    x_Sentencia = x_Sentencia & ",'" + gsAbvUsr & "'"
    x_Sentencia = x_Sentencia & "," + fDateNow() & " "
    x_Sentencia = x_Sentencia & ")"
'    x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
'    x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
'    x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
   
   With docnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
    .Execute x_Sentencia
   End With
   
fCnnClose docnnMain

'End Function
End Sub

Function gfctrl_obliga_sunat_mes(pmes As String) As String

    Dim dsCadena As String
    dsCadena = "Ene01Feb02Mar03Abr04May05Jun06Jul07Ago08Set09Oct10Nov11Dic12"
    pmes = Mid(dsCadena, InStr(dsCadena, pmes) + 3, 2)
    gfctrl_obliga_sunat_mes = pmes
End Function

'Function gfctrl_obliga_sunat_select() As ADODB.Recordset 'As Long() '
'fin 2015-08-27 ctr obligac sunat

'ini 2016-02-02.06  correccion ple
 Function gf_tb_sunat(xCodTabla As String) As String
    Dim xSource As String
    If xCodTabla = CODSUNAT_004 Then
    xSource = "SELECT codsunat, CONCAT(TRIM(campo03),'-',detsunat) detsunat,campo03 "
    Else
    xSource = "SELECT codsunat, detsunat,campo03 "
    End If
    xSource = xSource & "FROM tgsunat  "
    xSource = xSource & "WHERE codtabla='" & xCodTabla & "' "
    xSource = xSource & "AND estsunat ='" & ESTSUNAT_ACT & "' "

    gf_tb_sunat = xSource
End Function
Sub gpcbo_sunat_ins(xcombo As ComboBox, xrst As Recordset)
      xcombo.AddItem Choose(gsIdioma, "Ninguna", "Neither"), 0
        With xrst
            If Not (.EOF And .BOF()) Then .MoveFirst
            If Not .EOF Then
                Do While Not .EOF
                     xcombo.AddItem Trim(!CodSunat) & "-" & Left(Trim(!detsunat), 40)
                    .MoveNext
                Loop
            End If
        End With
End Sub
Sub gpcbo_sunat_update(xcombo As ComboBox, xrst As Recordset, xCampo As String, _
xCampoLen As Integer, xrstMain As Recordset)
'Sub gpcbo_sunat_update(xcombo As ComboBox, xrst As Recordset, xCampo As String, _
'xCampoLen As Integer, xform As Form)
         With xrst
            If Not (.EOF And .BOF) Then .MoveFirst
            .Find "CodSunat='" & Left(xcombo.Text, xCampoLen) & "'"
            If .EOF Then
               'frmTCprGrd.uorstMain.Fields(xCampo) = Null
               xrstMain.Fields(xCampo) = Null
            Else
               'frmTCprGrd.uorstMain.Fields(xCampo) = IIf(xcombo.ListIndex = 0, Null, !CodSunat)
               xrstMain.Fields(xCampo) = IIf(xcombo.ListIndex = 0, Null, !CodSunat)
           End If
         End With
End Sub
Sub gpcbo_sunat_index(xcombo As ComboBox, xCampo As String, _
xCampoLen As Integer, xrstMain As Recordset)
'Sub gpcbo_sunat_index(xcombo As ComboBox, xCampo As String, _
'xCampoLen As Integer, xform As Form)

  Dim dnContador As Integer
    xcombo.ListIndex = 0
    'With frmTCprGrd.uorstMain
    'With xform.uorstMain
    With xrstMain
     If Not IsNull(.Fields(xCampo).Value) Then
        For dnContador = 0 To xcombo.ListCount - 1
          If Left(.Fields(xCampo).Value, xCampoLen) = Left(xcombo.List(dnContador), xCampoLen) Then
            xcombo.ListIndex = dnContador
            Exit For
          End If
        Next dnContador
      End If
    End With
End Sub
Sub gpcbo_sunat_index2(xcombo As ComboBox, xCampo As String, xCampoLen As Integer)
  Dim dnContador As Integer
    xcombo.ListIndex = 0
        For dnContador = 0 To xcombo.ListCount - 1
          If xCampo = Left(xcombo.List(dnContador), xCampoLen) Then
            xcombo.ListIndex = dnContador
            Exit For
          End If
        Next dnContador
End Sub


' Function gf_tb_sunat_seek(xrst As Recordset, xCampo As String, xDato As String, xLabel As Label) As Boolean
 Function gf_tb_sunat_seek(xrst As Recordset, xDato As String, xLabel As Label) As Boolean
         With xrst
            'If .RecordCount > 0 Then .MoveFirst
            If Not (.EOF And .BOF) Then .MoveFirst
            'CodSunat
            '.Find xCampo & "='" & xDato & "'"
            .Find "CodSunat='" & xDato & "'"
            If .EOF Then
               'MsgBox TEXT_8006, vbExclamation
               gf_tb_sunat_seek = True
            Else
               'lblDatoDeta(tnIndex).Caption = " " & !razAux
               xLabel.Caption = " " & !detsunat
               gf_tb_sunat_seek = False
            End If
         End With
    'gf_tb_sunat = xSource
End Function

'fin 2016-02-02.06  correccion ple

