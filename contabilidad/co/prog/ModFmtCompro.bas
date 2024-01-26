Attribute VB_Name = "ModFmtCompro"
Function rcgfMesAct(tsMesAct As String)
'tsFecha          Mes a procesar.
   rcgfMesAct = tsMesAct
   'rcrcgfMesAct = IIf(gsMesAct = "00", gsMesApe, IIf(gsMesAct = "13", gsMesCie, gsMesAct))
End Function
Function rcgfNumComprobante(ByVal s_Ano As String, ByVal s_Mes As String, ByVal s_Diario As String) As String
  ' s_Ano             Año donde  se genera
  ' s_Mes             Mes donde  se genera
  ' s_Diario          Copdigo de diario para generar numero
    
  Dim porstRetorno As ADODB.Recordset
  Dim s_Sentencia As String
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), '000000') AS cNumMaxCpb "
  s_Sentencia = s_Sentencia & "FROM ComaCpbCab "
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
  rcgfNumComprobante = gfCeros(porstRetorno!cNumMaxCpb, 6, 1, "0")
  porstRetorno.Close
  Set porstRetorno = Nothing

End Function


