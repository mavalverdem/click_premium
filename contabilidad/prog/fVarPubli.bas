Attribute VB_Name = "fVarPubli"
Option Explicit
  
'ini 2016-06-02 adicion campo EstAct en Empresa
Public Const ESTEMPR_ACT As String = "A", _
             ESTEMPR_INA As String = "I"
'fin 2016-06-02 adicion campo EstAct en Empresa
  
'ini 2016-05-27/28 nivel=asisten no elimin datos
Public Const NVLUSR_ADMIN As String = "0", _
             NVLUSR_ASIS As String = "2"
'fin 2016-05-27/28 nivel=asisten no elimin datos
  
  
'ini 2015-10-09 aumento col detra, const.nro. y fech
Public Const INDCONS_DPO_0 As Byte = 0, _
             INDCONS_DPO_1 As Byte = 1
'fin 2015-10-09 aumento col detra, const.nro. y fech


'ini 2015-08-27 ctr obligac sunat
Public Const ESTDBUEN_CONTRI_ACT As String = "1", _
             ESTDBUEN_CONTRI_INA As String = "0"
'fin 2015-08-27 ctr obligac sunat

'ini 2015-07-02 adic tabla detrac
Public Const ESTDETRAC_ACT As String = "A", _
             ESTDETRAC_INA As String = "I"
'fin 2015-07-02 adic tabla detrac

'ini 2014-08-05 RR.HH afecto afp/onp
'2014-08-25 error flag Public Const INDCOMI_MIX As Byte = 1, _
'             INDCOMI_FLU As Byte = 0
Public Const INDCOMI_MIX As Byte = 0, _
             INDCOMI_FLU As Byte = 1
             
Public TPOCOMI_MIXTA_TXT As String, _
       TPOCOMI_FLUJO_TXT As String
Public Const s_FmtFeMysql_0 As String = "yyyy-mm-dd" ' Formato de fecha mysql visualizacion
       
'fin 2014-08-05 RR.HH afecto afp/onp

'ini 2016-02-02-03.05  correccion ple
Public Const ESTSUNAT_ACT As String = "A", _
             ESTSUNAT_INA As String = "I"
'codigos de tabla anexo sunat
'uorstTpoBns=30
'uorstCodMon=4
'uorstPais=35
'uorstCnveDobImpo=25
'uorstTpoRta=31
Public Const CODSUNAT_030 As String = "030", _
             CODSUNAT_004 As String = "004", _
             CODSUNAT_035 As String = "035", _
             CODSUNAT_025 As String = "025", _
             CODSUNAT_031 As String = "031"
             
Public Const CODMON_NAC As String = "PEN", _
             CODMON_EXT As String = "USD"
             
'fin 2016-02-02-03.05 correccion ple

'Public aDe
'2014-04-04 Codigo de detraccion
'Public aDtraccCod(205) As String

'ini 2015-07-02 adic tabla detrac
'Public aDtraccDet(205) As String
'Public aDtraccPor(205) As Double
'Public aDtraccEst(205) As Double
'fin 2015-07-02 adic tabla detrac


'2014-05-29
'Código del Plan de Cuentas utilizado por el deudor tributario
Public aCodPlCta(9) As String
Public aDetPlCta(9) As String
Public gnCodPlaCata As String

'2014-05-22
'estado del calculo del documento pedido
'CODPDO_IGV=igv no grabado
'CODPDO_IGVG=igv grabado 2014-07-18
Public Const CODPDO_IGV As String = "1", _
             CODPDO_HPR As String = "0", _
             CODPDO_IGVG As String = "2"
Public CODPDO_IGV_TXT As String, _
       CODPDO_HPR_TXT As String, _
       CODPDO_IGVG_TXT As String

Public Sub Mensajes2()

CODPDO_IGV_TXT = Choose(gsIdioma, "IGV Gravado ", "Excise Engraving ")
CODPDO_HPR_TXT = Choose(gsIdioma, "Honorarios ", "Fee ")
CODPDO_IGVG_TXT = Choose(gsIdioma, "IGV No Gravado   ", "Excise NoEngraving")
End Sub
          

Public Function fDiv0(fNumer As Double, fDenomi As Double) As Double
    
    fDiv0 = IIf(fDenomi = 0, 0, fNumer / fDenomi)
End Function

Public Function ValidateIdentificationDocumentPeru(identificationDocument As String) As Boolean
    If Not IsNull(identificationDocument) Then
        If Not Len(Trim(identificationDocument)) = 0 Then
            Dim addition As Integer
            addition = 0
            Dim hash(10) As Integer
            Dim n As Integer
            n = 0
            n = n + 1: hash(n) = 5
            n = n + 1: hash(n) = 4
            n = n + 1: hash(n) = 3
            n = n + 1: hash(n) = 2
            n = n + 1: hash(n) = 7
            n = n + 1: hash(n) = 6
            n = n + 1: hash(n) = 5
            n = n + 1: hash(n) = 4
            n = n + 1: hash(n) = 3
            n = n + 1: hash(n) = 2
            Dim identificationDocumentLength  As Integer
            identificationDocumentLength = Len(Trim(identificationDocument))
            Dim identificationComponent As String
            identificationComponent = Mid(identificationDocument, 0, identificationDocumentLength)
            Dim identificationComponentLength  As Integer
            identificationComponentLength = Len(identificationComponent)
            Dim diff As Integer
            diff = UBound(hash, 1) - identificationDocumentLength
            Dim i As Integer
            For i = identificationComponentLength - 1 To i >= 0 Step -1
            'MsgBox "x = " + Str(X)
            addition = addition + 1
            Next
        End If
    End If
    
    ValidateIdentificationDocumentPeru = False
End Function

