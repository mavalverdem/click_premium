Attribute VB_Name = "modMainCO"
Option Explicit

Public Const CODEMP_COMPASS As String = "001"
Public Const CODEMP_NORMAL As String = "000"

Public Const s_FmtFeHoMysql_0 As String = "yyyy-mm-dd hh:mm:ss" ' Formato de fecha y hora larga mysql visualizacion
Public Const CTAFZD_CTA As String = "FF"
Public Const BSEPCT_ACT As Byte = 1, _
             BSEPCT_INA As Byte = 0
Public Const ESTAUX_ACT As String = "A", _
             ESTAUX_INA As String = "I", _
             ESTCCO_ACT As String = "A", _
             ESTCCO_INA As String = "I", _
             ESTCTA_ACT As String = "A", _
             ESTCTA_INA As String = "I"
Public ESTAUX_ACT_TXT As String, _
       ESTAUX_INA_TXT As String, _
       ESTCCO_ACT_TXT As String, _
       ESTCCO_INA_TXT As String, _
       ESTCTA_ACT_TXT As String, _
       ESTCTA_INA_TXT As String
Public Const FORMATO_NUM_1 As String = "###,###,##0.00", _
             FORMATO_NUM_2 As String = "#0.0000", _
             FORMATO_NUM_3 As String = "#0.000000", _
             FORMATO_NUM_4 As String = "#0.00"
Public Const INDCNV_VER As Byte = 1, _
             INDCNV_FAL As Byte = 0, _
             INDLAT_DER As Byte = 0, _
             INDLAT_IZQ As Byte = 1
Public Const INDAJD_ACT As Byte = 1, _
             INDAJD_INA As Byte = 0, _
             INDAUX_CLI_ACT As Byte = 1, _
             INDAUX_CLI_INA As Byte = 0, _
             INDAUX_PRV_ACT As Byte = 1, _
             INDAUX_PRV_INA As Byte = 0, _
             INDAUX_OTR_ACT As Byte = 1, _
             INDAUX_OTR_INA As Byte = 0, _
             INDCCO_ACT As Byte = 1, _
             INDCCO_INA As Byte = 0, _
             INDCDT_ACT As Byte = 1, _
             INDCDT_INA As Byte = 0, _
             INDDOC_ACT As Byte = 1, _
             INDAUX_ACT As Byte = 0, _
             INDDOC_INA As Byte = 0, _
             INDMNE_ACT As Byte = 1, _
             INDMNE_INA As Byte = 0, _
             INDMOE_ACT As Byte = 1, _
             INDMOE_INA As Byte = 0, _
             INDNCU_VER As Byte = 1, _
             INDNCU_FAL As Byte = 0
Public Const INDPREGEN_ACT As Byte = 1, _
             INDPREGEN_INA As Byte = 0, _
             INDPREGEN_ACTx As Byte = 1, _
             INDPREGEN_INAx As Byte = 0, _
             INDPSP_ACT As Byte = 1, _
             INDPSP_INA As Byte = 0, _
             INDSDO_AMB As Byte = 0, _
             INDSDO_POS As Byte = 1, _
             INDSDO_NEG As Byte = 2
Public INDSDO_POS_TXT As String, _
       INDSDO_NEG_TXT As String
Public Const INDAMO_ACT As Byte = 1, _
             INDAMO_INA As Byte = 0, _
             INDCCT_ACT As Byte = 1, _
             INDCCT_INA As Byte = 0, _
             INDHTR_ACT As Byte = 1, _
             INDHTR_INA As Byte = 0, _
             NATCTA_DEU As Byte = 0, _
             NATCTA_ACR As Byte = 1, _
             SGNTDC_POS As Byte = 1, _
             SGNTDC_NEG As Byte = 0
'             INDAJD_ACT_TXT As String = "Aplicar Ajuste DC", _
             INDAJD_INA_TXT As String = "No Aplicar Ajuste DC", _
             INDCCO_ACT_TXT As String = "Solicitar CC", _
             INDCCO_INA_TXT As String = "No Solicitar CC", _
             INDDOC_ACT_TXT As String = "Solicitar CC", _
             INDDOC_INA_TXT As String = "No Solicitar CC"
Public NATCTA_DEU_TXT As String, _
       NATCTA_ACR_TXT As String, _
       SGNTDC_POS_TXT As String, _
       SGNTDC_NEG_TXT As String
Public Const TPOANL_CTA As Byte = 0, _
             TPOANL_AUX As Byte = 1, _
             TPOANL_DOC As Byte = 2, _
             TPOCTA_TIT As Byte = 0, _
             TPOCTA_TRA As Byte = 1, _
             TPOCTB_DEB As String = "D", _
             TPOCTB_HAB As String = "H", _
             TPOCNC_TOT_CPR As Byte = 11, _
             TPOCNC_TOT_HPR As Byte = 1, _
             TPOCNC_TOT_VTA As Byte = 7
Public TPOANL_CTA_TXT As String, _
       TPOANL_AUX_TXT As String, _
       TPOANL_DOC_TXT As String, _
       TPOCTA_TIT_TXT As String, _
       TPOCTA_TRA_TXT As String

Public Const TPOGNR_DRO As Byte = 0, _
             TPOGNR_CPR As Byte = 1, _
             TPOGNR_VTA As Byte = 2, _
             TPOGNR_HPR As Byte = 3, _
             TPOGNR_DST As Byte = 4, _
             TPOGNR_DCA As Byte = 5, _
             TPOGNR_APE As Byte = 6, _
             TPOGNR_CIE As Byte = 7, _
             TPOGNR_DRP As Byte = 8, _
             TPOGNR_BAN As Byte = 9, _
             TPOHTR_HT1 As Byte = 0, _
             TPOHTR_HT2 As Byte = 1, _
             TPOHTR_HT3 As Byte = 2, _
             TPOHT1_SAL As Byte = 0, _
             TPOHT1_DEP As Byte = 1
Public TPOGNR_DRO_TXT As String, _
       TPOGNR_CPR_TXT As String, _
       TPOGNR_VTA_TXT As String, _
       TPOGNR_HPR_TXT As String, _
       TPOGNR_DST_TXT As String, _
       TPOGNR_DCA_TXT As String, _
       TPOGNR_APE_TXT As String, _
       TPOGNR_CIE_TXT As String, _
       TPOGNR_DRP_TXT As String, _
       TPOGNR_BAN_TXT As String, _
       TPOHT1_SAL_TXT As String, _
       TPOHT1_DEP_TXT As String
Public Const TPOLIN_TIT As Byte = 0, _
             TPOLIN_STO As Byte = 1, _
             TPOLIN_TOT As Byte = 2, _
             TPOLIN_OPE As Byte = 3, _
             TPOLIN_MAS As Byte = 4, _
             TPOMON_NAC As String = "N", _
             TPOMON_EXT As String = "E", _
             TPOMON_NAC_IND As Byte = 0, _
             TPOMON_EXT_IND As Byte = 1
Public TPOLIN_CTA_TXT As String, _
       TPOLIN_TIT_TXT As String, _
       TPOLIN_STO_TXT As String, _
       TPOLIN_TOT_TXT As String, _
       TPOLIN_OPE_TXT As String, _
       TPOLIN_MAS_TXT As String, _
       TPOMON_NAC_TXT_0 As String, _
       TPOMON_EXT_TXT_0 As String, _
       TPOMON_NAC_TXT_1 As String, _
       TPOMON_EXT_TXT_1 As String, _
       TPOMON_NAC_TXT_2 As String, _
       TPOMON_EXT_TXT_2 As String
Public Const TPOMON_NAC_TXT As String = "MN", _
             TPOMON_EXT_TXT As String = "ME"
'             TPOMON_NAC_SIG As String = "S/.", _
             TPOMON_EXT_SIG As String = "US$",
Public Const TPOPER_NAT As String = "N", _
             TPOPER_JUR As String = "J", _
             TPOPER_DOM As String = "D"
Public Const TPOPVS_PVS As String = "P", _
             TPOPVS_CAN As String = "C", _
             TPOPVS_OTR As String = "O", _
             TPOPVS_PVS_VER As Boolean = True, _
             TPOPVS_CAN_VER As Boolean = True, _
             TPOPVS_OTR_VER As Boolean = True, _
             TPOPVS_PVS_FAL As Boolean = False, _
             TPOPVS_CAN_FAL As Boolean = False, _
             TPOPVS_OTR_FAL As Boolean = False, _
             TPOSDO_INV As String = "I", _
             TPOSDO_RES As String = "R", _
             TPOSDO_FUN As String = "F", _
             TPOSDO_NAT As String = "N", _
             TPOSDO_AMB As String = "A", _
             TPOTCB_CPR As String = "C", _
             TPOTCB_VTA As String = "V", _
             TPOTCB_CPR_IND As Byte = 1, _
             TPOTCB_VTA_IND As Byte = 0
Public TPOSDO_INV_TXT As String, _
       TPOSDO_RES_TXT As String, _
       TPOSDO_FUN_TXT As String, _
       TPOSDO_NAT_TXT As String, _
       TPOSDO_AMB_TXT As String, _
       TPOTCB_CPR_TXT As String, _
       TPOTCB_VTA_TXT As String
Public Const TPOGRU1_TXT_0 As String = "A", _
             TPOGRU2_TXT_0 As String = "B", _
             TPOGRU3_TXT_0 As String = "C", _
             TPOGRU4_TXT_0 As String = "9", _
             TPOGRU1_IND As Byte = 0, _
             TPOGRU2_IND As Byte = 1, _
             TPOGRU3_IND As Byte = 2
Public TPOGRU1_TXT_1 As String, _
       TPOGRU2_TXT_1 As String, _
       TPOGRU3_TXT_1 As String, _
       TPOGRU4_TXT_1 As String

Public Const TPOFJO_ING As Byte = 0, _
             TPOFJO_EGR As Byte = 1, _
             INDFJO_ACT As Byte = 1, _
             INDFJO_INA As Byte = 0
Public TPOFJO_ING_TXT As String, _
       TPOFJO_EGR_TXT As String

Public Const TPOEFE_OPE As Byte = 0, _
             TPOEFE_INV As Byte = 1, _
             TPOEFE_FIN As Byte = 2
Public TPOEFE_OPE_TXT As String, _
       TPOEFE_INV_TXT As String, _
       TPOEFE_FIN_TXT As String

'TEXT_9017=ini 2014-07-10 validacion T.Doc=05
''ini 2016-05-27/28 TEXT_9018=asistente no puede eliminar datos
Public TEXT_9011 As String, _
       TEXT_9012 As String, _
       TEXT_9013 As String, _
       TEXT_9015 As String, _
       TEXT_9016 As String, _
       TEXT_9017 As String
       
'ini 2015-05-18 validacion frm
Public TEXT_9018 As String, _
       TEXT_9019 As String
'fin 2015-05-18 validacion frm
'ini 2015-06-06 Si Mayorizo o no . Estado Mayorizacion
Public TEXT_9020 As String
'fin 2015-06-06 Si Mayorizo o no . Estado Mayorizacion
Public TEXT_9021 As String '2015-06-30 correccion tipo mon cta

Public TEXT_9022 As String '2015-07-17 error eliminar detraccion

Public TEXT_9023 As String '2015-07-21 t.cambio sunat

'ini 2015-08-27 ctr obligac sunat
Public TEXT_9024 As String, _
       TEXT_9025 As String
'fin 2015-08-27 ctr obligac sunat

'ini 2016-05-27/28 nivel=asisten no elimin datos
Public TEXT_9026 As String
'fin 2016-05-27/28 nivel=asisten no elimin datos


Public Const TPOBAN_ING As Byte = 0, _
             TPOBAN_EGR As Byte = 1
Public TPOBAN_ING_TXT As String, _
       TPOBAN_EGR_TXT As String

Public Const TPODOC_DPS_IND As Byte = 1, _
             TPODOC_GRO_IND As Byte = 2, _
             TPODOC_TRA_IND As Byte = 3, _
             TPODOC_ORD_IND As Byte = 4, _
             TPODOC_DEB_IND As Byte = 5, _
             TPODOC_CRE_IND As Byte = 6, _
             TPODOC_CHQ_IND As Byte = 7, _
             TPODOC_OTR_IND As Byte = 8, _
             TPODOC_EFE_IND As Byte = 9, _
             TPODOC_PEX_IND As Byte = 10, _
             TPODOC_LTR_IND As Byte = 11, _
             TPODOC_CGE_IND As Byte = 12
Public TPODOC_DPS_TXT As String, _
       TPODOC_GRO_TXT As String, _
       TPODOC_TRA_TXT As String, _
       TPODOC_ORD_TXT As String, _
       TPODOC_DEB_TXT As String, _
       TPODOC_CRE_TXT As String, _
       TPODOC_CHQ_TXT As String, _
       TPODOC_OTR_TXT As String, _
       TPODOC_EFE_TXT As String, _
       TPODOC_PEX_TXT As String, _
       TPODOC_LTR_TXT As String, _
       TPODOC_CGE_TXT As String

Public TPOCTA_AHO_TXT_2 As String, _
       TPOCTA_COR_TXT_2 As String, _
       TPOCTA_MAE_TXT_2 As String, _
       TPOCTA_SIN_TXT_2 As String
Public Const TPOCTA_COR As String = "0", _
             TPOCTA_AHO As String = "1", _
             TPOCTA_MAE As String = "2", _
             TPOCTA_SIN As String = "3", _
             TPOCTA_COR_IND As Byte = 0, _
             TPOCTA_AHO_IND As Byte = 1, _
             TPOCTA_MAE_IND As Byte = 2, _
             TPOCTA_SIN_IND As Byte = 3

Public gbCieCpr As Boolean             'Indicador: Cerrado Compras.
Public gbCieVta As Boolean             'Indicador: Cerrado Ventas.
Public gbCieHpr As Boolean             'Indicador: Cerrado Honorarios.
Public gbCieCpb As Boolean             'Indicador: Cerrado Diario.

Public gsTpoMon_Fnc As String          'Moneda Funcional
Public gnIndMNE As Byte                '0:1 Moneda (Nacional) / 1:2 Monedas (Nacional y Extranjera)
Public gsTpoMon_Sgn_MN As String       'Signo de Moneda Nacional.
Public gsTpoMon_Sgn_ME As String       'Signo de Moneda Extranjera.
Public gsNivCta As String              'Niveles de Cuenta
Public gsNivCCo As String              'Niveles de Centro de Costo
Public gsIniAno(24) As String          'Saldo Inicial del Año.
'Public gsIniMes(24) As String          'Saldo Inicial del Mes.
Public gsAcuAnt(24) As String          'Importes Acumulados al Mes Anterior.
Public gsAcuMes(24) As String          'Importes Acumulados del Mes.

'Para Tablas:   1:Cta Debe MN
'               2:Cta Debe ME
'               3:Cta Haber MN
'               4:Cta Haber ME
'               5:Aux Debe MN
'               6:Aux Debe ME
'               7:Aux Haber MN
'               8:Aux Haber ME
'               9:CCo Debe MN
'              10:CCo Debe ME
'              11:CCo Haber MN
'              12:CCo Haber ME
'Para Reportes rpt:
'              13:Cta Debe MN
'              14:Cta Debe ME
'              15:Cta Haber MN
'              16:Cta Haber ME
'              17:Aux Debe MN
'              18:Aux Debe ME
'              19:Aux Haber MN
'              20:Aux Haber ME
'              21:CCo Debe MN
'              22:CCo Debe ME
'              23:CCo Haber MN
'              24:CCo Haber ME

'Sub PublicasModulo()
'Sólo existe para poder cargar las constantes públicas al módulo.
'End Sub

Sub gpCamposSaldos()
   Dim dsTablaCta As String, dsTablaAux, dsTablaCCo, _
       dsCi As String, dsCF As String, _
       dnContador As Byte

   dsTablaCta = "COCtaAcu."
   dsTablaAux = "COAuxAcu."
   dsTablaCCo = "COCCoAcu."
'   dsCi = "{"
'   dsCF = "}"

   gsIniAno(1) = dsTablaCta & "AcuD00_MN"
'   gsIniMes(1) = gsIniAno(1)
   gsAcuMes(1) = dsTablaCta & "AcuD" & gsMesAct & "_MN"
   gsIniAno(2) = dsTablaCta & "AcuD00_ME"
'   gsIniMes(2) = gsIniAno(2)
   gsAcuMes(2) = dsTablaCta & "AcuD" & gsMesAct & "_ME"
   gsIniAno(3) = dsTablaCta & "AcuH00_MN"
'   gsIniMes(3) = gsIniAno(3)
   gsAcuMes(3) = dsTablaCta & "AcuH" & gsMesAct & "_MN"
   gsIniAno(4) = dsTablaCta & "AcuH00_ME"
'   gsIniMes(4) = gsIniAno(4)
   gsAcuMes(4) = dsTablaCta & "AcuH" & gsMesAct & "_ME"
   
   gsIniAno(5) = dsTablaAux & "AcuD00_MN"
'   gsIniMes(5) = gsIniAno(5)
   gsAcuMes(5) = dsTablaAux & "AcuD" & gsMesAct & "_MN"
   gsIniAno(6) = dsTablaAux & "AcuD00_ME"
'   gsIniMes(6) = gsIniAno(6)
   gsAcuMes(6) = dsTablaAux & "AcuD" & gsMesAct & "_ME"
   gsIniAno(7) = dsTablaAux & "AcuH00_MN"
'   gsIniMes(7) = gsIniAno(7)
   gsAcuMes(7) = dsTablaAux & "AcuH" & gsMesAct & "_MN"
   gsIniAno(8) = dsTablaAux & "AcuH00_ME"
'   gsIniMes(8) = gsIniAno(8)
   gsAcuMes(8) = dsTablaAux & "AcuH" & gsMesAct & "_ME"
   
   gsIniAno(9) = dsTablaCCo & "AcuD00_MN"
'   gsIniMes(9) = gsIniAno(9)
   gsAcuMes(9) = dsTablaCCo & "AcuD" & gsMesAct & "_MN"
   gsIniAno(10) = dsTablaCCo & "AcuD00_ME"
'   gsIniMes(10) = gsIniAno(10)
   gsAcuMes(10) = dsTablaCCo & "AcuD" & gsMesAct & "_ME"
   gsIniAno(11) = dsTablaCCo & "AcuH00_MN"
'   gsIniMes(11) = gsIniAno(11)
   gsAcuMes(11) = dsTablaCCo & "AcuH" & gsMesAct & "_MN"
   gsIniAno(12) = dsTablaCCo & "AcuH00_ME"
'   gsIniMes(12) = gsIniAno(12)
   gsAcuMes(12) = dsTablaCCo & "AcuH" & gsMesAct & "_ME"
   
'   gsIniAno(1 + 12) = dsCi & dsTablaCta & "AcuD00_MN" & dsCF
'   gsIniMes(1 + 12) = gsIniAno(1 + 12)
'   gsAcuMes(1 + 12) = dsCi & dsTablaCta & "AcuD" & gsMesAct & "_MN" & dsCF
'   gsIniAno(2 + 12) = dsCi & dsTablaCta & "AcuD00_ME" & dsCF
'   gsIniMes(2 + 12) = gsIniAno(2 + 12)
'   gsAcuMes(2 + 12) = dsCi & dsTablaCta & "AcuD" & gsMesAct & "_ME" & dsCF
'   gsIniAno(3 + 12) = dsCi & dsTablaCta & "AcuH00_MN" & dsCF
'   gsIniMes(3 + 12) = gsIniAno(3 + 12)
'   gsAcuMes(3 + 12) = dsCi & dsTablaCta & "AcuH" & gsMesAct & "_MN" & dsCF
'   gsIniAno(4 + 12) = dsCi & dsTablaCta & "AcuH00_ME" & dsCF
'   gsIniMes(4 + 12) = gsIniAno(4 + 12)
'   gsAcuMes(4 + 12) = dsCi & dsTablaCta & "AcuH" & gsMesAct & "_ME" & dsCF

'   gsIniAno(5 + 12) = dsCi & dsTablaAux & "AcuD00_MN" & dsCF
'   gsIniMes(5 + 12) = gsIniAno(5 + 12)
'   gsAcuMes(5 + 12) = dsCi & dsTablaAux & "AcuD" & gsMesAct & "_MN" & dsCF
'   gsIniAno(6 + 12) = dsCi & dsTablaAux & "AcuD00_ME" & dsCF
'   gsIniMes(6 + 12) = gsIniAno(6 + 12)
'   gsAcuMes(6 + 12) = dsCi & dsTablaAux & "AcuD" & gsMesAct & "_ME" & dsCF
'   gsIniAno(7 + 12) = dsCi & dsTablaAux & "AcuH00_MN" & dsCF
'   gsIniMes(7 + 12) = gsIniAno(7 + 12)
'   gsAcuMes(7 + 12) = dsCi & dsTablaAux & "AcuH" & gsMesAct & "_MN" & dsCF
'   gsIniAno(8 + 12) = dsCi & dsTablaAux & "AcuH00_ME" & dsCF
'   gsIniMes(8 + 12) = gsIniAno(8 + 12)
'   gsAcuMes(8 + 12) = dsCi & dsTablaAux & "AcuH" & gsMesAct & "_ME" & dsCF
   
'   gsIniAno(9 + 12) = dsCi & dsTablaCCo & "AcuD00_MN" & dsCF
'   gsIniMes(9 + 12) = gsIniAno(9 + 12)
'   gsAcuMes(9 + 12) = dsCi & dsTablaCCo & "AcuD" & gsMesAct & "_MN" & dsCF
'   gsIniAno(10 + 12) = dsCi & dsTablaCCo & "AcuD00_ME" & dsCF
'   gsIniMes(10 + 12) = gsIniAno(10 + 12)
'   gsAcuMes(10 + 12) = dsCi & dsTablaCCo & "AcuD" & gsMesAct & "_ME" & dsCF
'   gsIniAno(11 + 12) = dsCi & dsTablaCCo & "AcuH00_MN" & dsCF
'   gsIniMes(11 + 12) = gsIniAno(11 + 12)
'   gsAcuMes(11 + 12) = dsCi & dsTablaCCo & "AcuH" & gsMesAct & "_MN" & dsCF
'   gsIniAno(12 + 12) = dsCi & dsTablaCCo & "AcuH00_ME" & dsCF
'   gsIniMes(12 + 12) = gsIniAno(12 + 12)
'   gsAcuMes(12 + 12) = dsCi & dsTablaCCo & "AcuH" & gsMesAct & "_ME" & dsCF

   If gsMesAct > "00" Then
   
      gsAcuAnt(1) = "(" & dsTablaCta & "AcuD00_MN"
      gsAcuAnt(2) = "(" & dsTablaCta & "AcuD00_ME"
      gsAcuAnt(3) = "(" & dsTablaCta & "AcuH00_MN"
      gsAcuAnt(4) = "(" & dsTablaCta & "AcuH00_ME"

      gsAcuAnt(5) = "(" & dsTablaAux & "AcuD00_MN"
      gsAcuAnt(6) = "(" & dsTablaAux & "AcuD00_ME"
      gsAcuAnt(7) = "(" & dsTablaAux & "AcuH00_MN"
      gsAcuAnt(8) = "(" & dsTablaAux & "AcuH00_ME"

      gsAcuAnt(9) = "(" & dsTablaCCo & "AcuD00_MN"
      gsAcuAnt(10) = "(" & dsTablaCCo & "AcuD00_ME"
      gsAcuAnt(11) = "(" & dsTablaCCo & "AcuH00_MN"
      gsAcuAnt(12) = "(" & dsTablaCCo & "AcuH00_ME"

      gsAcuAnt(1 + 12) = "(" & dsCi & dsTablaCta & "AcuD00_MN" & dsCF
      gsAcuAnt(2 + 12) = "(" & dsCi & dsTablaCta & "AcuD00_ME" & dsCF
      gsAcuAnt(3 + 12) = "(" & dsCi & dsTablaCta & "AcuH00_MN" & dsCF
      gsAcuAnt(4 + 12) = "(" & dsCi & dsTablaCta & "AcuH00_ME" & dsCF

      gsAcuAnt(5 + 12) = "(" & dsCi & dsTablaAux & "AcuD00_MN" & dsCF
      gsAcuAnt(6 + 12) = "(" & dsCi & dsTablaAux & "AcuD00_ME" & dsCF
      gsAcuAnt(7 + 12) = "(" & dsCi & dsTablaAux & "AcuH00_MN" & dsCF
      gsAcuAnt(8 + 12) = "(" & dsCi & dsTablaAux & "AcuH00_ME" & dsCF

      gsAcuAnt(9 + 12) = "(" & dsCi & dsTablaCCo & "AcuD00_MN" & dsCF
      gsAcuAnt(10 + 12) = "(" & dsCi & dsTablaCCo & "AcuD00_ME" & dsCF
      gsAcuAnt(11 + 12) = "(" & dsCi & dsTablaCCo & "AcuH00_MN" & dsCF
      gsAcuAnt(12 + 12) = "(" & dsCi & dsTablaCCo & "AcuH00_ME" & dsCF
      
      If gsMesAct > "01" Then
      
         gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD01_MN"
         gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD01_ME"
         gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH01_MN"
         gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH01_ME"
      
         gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD01_MN"
         gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD01_ME"
         gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH01_MN"
         gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH01_ME"
      
         gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD01_MN"
         gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD01_ME"
         gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH01_MN"
         gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH01_ME"
         
         gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD01_MN" & dsCF
         gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD01_ME" & dsCF
         gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH01_MN" & dsCF
         gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH01_ME" & dsCF

         gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD01_MN" & dsCF
         gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD01_ME" & dsCF
         gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH01_MN" & dsCF
         gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH01_ME" & dsCF
      
         gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD01_MN" & dsCF
         gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD01_ME" & dsCF
         gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH01_MN" & dsCF
         gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH01_ME" & dsCF
            
         If gsMesAct > "02" Then
            gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD02_MN"
            gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD02_ME"
            gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH02_MN"
            gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH02_ME"
         
            gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD02_MN"
            gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD02_ME"
            gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH02_MN"
            gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH02_ME"
         
            gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD02_MN"
            gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD02_ME"
            gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH02_MN"
            gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH02_ME"
            
            gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD02_MN" & dsCF
            gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD02_ME" & dsCF
            gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH02_MN" & dsCF
            gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH02_ME" & dsCF
   
            gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD02_MN" & dsCF
            gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD02_ME" & dsCF
            gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH02_MN" & dsCF
            gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH02_ME" & dsCF
         
            gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD02_MN" & dsCF
            gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD02_ME" & dsCF
            gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH02_MN" & dsCF
            gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH02_ME" & dsCF
            
            If gsMesAct > "03" Then
               gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD03_MN"
               gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD03_ME"
               gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH03_MN"
               gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH03_ME"
         
               gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD03_MN"
               gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD03_ME"
               gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH03_MN"
               gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH03_ME"
         
               gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD03_MN"
               gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD03_ME"
               gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH03_MN"
               gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH03_ME"
            
               gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD03_MN" & dsCF
               gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD03_ME" & dsCF
               gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH03_MN" & dsCF
               gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH03_ME" & dsCF
   
               gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD03_MN" & dsCF
               gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD03_ME" & dsCF
               gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH03_MN" & dsCF
               gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH03_ME" & dsCF
         
               gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD03_MN" & dsCF
               gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD03_ME" & dsCF
               gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH03_MN" & dsCF
               gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH03_ME" & dsCF
   
               If gsMesAct > "04" Then
                  gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD04_MN"
                  gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD04_ME"
                  gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH04_MN"
                  gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH04_ME"
         
                  gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD04_MN"
                  gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD04_ME"
                  gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH04_MN"
                  gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH04_ME"
         
                  gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD04_MN"
                  gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD04_ME"
                  gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH04_MN"
                  gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH04_ME"
            
                  gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD04_MN" & dsCF
                  gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD04_ME" & dsCF
                  gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH04_MN" & dsCF
                  gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH04_ME" & dsCF
   
                  gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD04_MN" & dsCF
                  gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD04_ME" & dsCF
                  gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH04_MN" & dsCF
                  gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH04_ME" & dsCF
         
                  gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD04_MN" & dsCF
                  gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD04_ME" & dsCF
                  gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH04_MN" & dsCF
                  gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH04_ME" & dsCF
   
                  If gsMesAct > "05" Then
                     gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD05_MN"
                     gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD05_ME"
                     gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH05_MN"
                     gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH05_ME"
         
                     gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD05_MN"
                     gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD05_ME"
                     gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH05_MN"
                     gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH05_ME"
         
                     gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD05_MN"
                     gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD05_ME"
                     gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH05_MN"
                     gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH05_ME"
            
                     gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD05_MN" & dsCF
                     gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD05_ME" & dsCF
                     gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH05_MN" & dsCF
                     gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH05_ME" & dsCF
   
                     gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD05_MN" & dsCF
                     gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD05_ME" & dsCF
                     gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH05_MN" & dsCF
                     gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH05_ME" & dsCF
         
                     gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD05_MN" & dsCF
                     gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD05_ME" & dsCF
                     gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH05_MN" & dsCF
                     gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH05_ME" & dsCF
   
                     If gsMesAct > "06" Then
                        gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD06_MN"
                        gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD06_ME"
                        gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH06_MN"
                        gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH06_ME"
         
                        gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD06_MN"
                        gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD06_ME"
                        gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH06_MN"
                        gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH06_ME"
         
                        gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD06_MN"
                        gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD06_ME"
                        gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH06_MN"
                        gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH06_ME"
            
                        gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD06_MN" & dsCF
                        gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD06_ME" & dsCF
                        gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH06_MN" & dsCF
                        gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH06_ME" & dsCF
   
                        gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD06_MN" & dsCF
                        gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD06_ME" & dsCF
                        gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH06_MN" & dsCF
                        gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH06_ME" & dsCF
         
                        gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD06_MN" & dsCF
                        gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD06_ME" & dsCF
                        gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH06_MN" & dsCF
                        gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH06_ME" & dsCF
   
                        If gsMesAct > "07" Then
                           gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD07_MN"
                           gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD07_ME"
                           gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH07_MN"
                           gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH07_ME"
         
                           gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD07_MN"
                           gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD07_ME"
                           gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH07_MN"
                           gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH07_ME"
         
                           gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD07_MN"
                           gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD07_ME"
                           gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH07_MN"
                           gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH07_ME"
            
                           gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD07_MN" & dsCF
                           gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD07_ME" & dsCF
                           gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH07_MN" & dsCF
                           gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH07_ME" & dsCF
   
                           gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD07_MN" & dsCF
                           gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD07_ME" & dsCF
                           gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH07_MN" & dsCF
                           gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH07_ME" & dsCF
         
                           gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD07_MN" & dsCF
                           gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD07_ME" & dsCF
                           gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH07_MN" & dsCF
                           gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH07_ME" & dsCF
   
                           If gsMesAct > "08" Then
                              gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD08_MN"
                              gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD08_ME"
                              gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH08_MN"
                              gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH08_ME"
         
                              gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD08_MN"
                              gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD08_ME"
                              gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH08_MN"
                              gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH08_ME"
         
                              gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD08_MN"
                              gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD08_ME"
                              gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH08_MN"
                              gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH08_ME"
            
                              gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD08_MN" & dsCF
                              gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD08_ME" & dsCF
                              gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH08_MN" & dsCF
                              gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH08_ME" & dsCF
   
                              gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD08_MN" & dsCF
                              gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD08_ME" & dsCF
                              gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH08_MN" & dsCF
                              gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH08_ME" & dsCF
         
                              gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD08_MN" & dsCF
                              gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD08_ME" & dsCF
                              gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH08_MN" & dsCF
                              gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH08_ME" & dsCF
   
                              If gsMesAct > "09" Then
                                 gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD09_MN"
                                 gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD09_ME"
                                 gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH09_MN"
                                 gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH09_ME"
         
                                 gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD09_MN"
                                 gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD09_ME"
                                 gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH09_MN"
                                 gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH09_ME"
         
                                 gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD09_MN"
                                 gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD09_ME"
                                 gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH09_MN"
                                 gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH09_ME"
            
                                 gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD09_MN" & dsCF
                                 gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD09_ME" & dsCF
                                 gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH09_MN" & dsCF
                                 gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH09_ME" & dsCF
   
                                 gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD09_MN" & dsCF
                                 gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD09_ME" & dsCF
                                 gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH09_MN" & dsCF
                                 gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH09_ME" & dsCF
         
                                 gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD09_MN" & dsCF
                                 gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD09_ME" & dsCF
                                 gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH09_MN" & dsCF
                                 gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH09_ME" & dsCF
   
                                 If gsMesAct > "10" Then
                                    gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD10_MN"
                                    gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD10_ME"
                                    gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH10_MN"
                                    gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH10_ME"
         
                                    gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD10_MN"
                                    gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD10_ME"
                                    gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH10_MN"
                                    gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH10_ME"
         
                                    gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD10_MN"
                                    gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD10_ME"
                                    gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH10_MN"
                                    gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH10_ME"
            
                                    gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD10_MN" & dsCF
                                    gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD10_ME" & dsCF
                                    gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH10_MN" & dsCF
                                    gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH10_ME" & dsCF
   
                                    gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD10_MN" & dsCF
                                    gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD10_ME" & dsCF
                                    gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH10_MN" & dsCF
                                    gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH10_ME" & dsCF
         
                                    gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD10_MN" & dsCF
                                    gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD10_ME" & dsCF
                                    gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH10_MN" & dsCF
                                    gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH10_ME" & dsCF
   
                                    If gsMesAct > "11" Then
                                       gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD11_MN"
                                       gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD11_ME"
                                       gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH11_MN"
                                       gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH11_ME"
         
                                       gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD11_MN"
                                       gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD11_ME"
                                       gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH11_MN"
                                       gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH11_ME"
         
                                       gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD11_MN"
                                       gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD11_ME"
                                       gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH11_MN"
                                       gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH11_ME"
            
                                       gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD11_MN" & dsCF
                                       gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD11_ME" & dsCF
                                       gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH11_MN" & dsCF
                                       gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH11_ME" & dsCF
   
                                       gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD11_MN" & dsCF
                                       gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD11_ME" & dsCF
                                       gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH11_MN" & dsCF
                                       gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH11_ME" & dsCF
         
                                       gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD11_MN" & dsCF
                                       gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD11_ME" & dsCF
                                       gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH11_MN" & dsCF
                                       gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH11_ME" & dsCF
   
                                       If gsMesAct > "12" Then
                                          gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD12_MN"
                                          gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD12_ME"
                                          gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH12_MN"
                                          gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH12_ME"
         
                                          gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD12_MN"
                                          gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD12_ME"
                                          gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH12_MN"
                                          gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH12_ME"
         
                                          gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD12_MN"
                                          gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD12_ME"
                                          gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH12_MN"
                                          gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH12_ME"
            
                                          gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD12_MN" & dsCF
                                          gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD12_ME" & dsCF
                                          gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH12_MN" & dsCF
                                          gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH12_ME" & dsCF
   
                                          gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD12_MN" & dsCF
                                          gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD12_ME" & dsCF
                                          gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH12_MN" & dsCF
                                          gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH12_ME" & dsCF
            
                                          gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD12_MN" & dsCF
                                          gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD12_ME" & dsCF
                                          gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH12_MN" & dsCF
                                          gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH12_ME" & dsCF
      
                                          If gsMesAct > "13" Then
                                             gsAcuAnt(1) = gsAcuAnt(1) & "+" & dsTablaCta & "AcuD13_MN"
                                             gsAcuAnt(2) = gsAcuAnt(2) & "+" & dsTablaCta & "AcuD13_ME"
                                             gsAcuAnt(3) = gsAcuAnt(3) & "+" & dsTablaCta & "AcuH13_MN"
                                             gsAcuAnt(4) = gsAcuAnt(4) & "+" & dsTablaCta & "AcuH13_ME"
         
                                             gsAcuAnt(5) = gsAcuAnt(5) & "+" & dsTablaAux & "AcuD13_MN"
                                             gsAcuAnt(6) = gsAcuAnt(6) & "+" & dsTablaAux & "AcuD13_ME"
                                             gsAcuAnt(7) = gsAcuAnt(7) & "+" & dsTablaAux & "AcuH13_MN"
                                             gsAcuAnt(8) = gsAcuAnt(8) & "+" & dsTablaAux & "AcuH13_ME"
         
                                             gsAcuAnt(9) = gsAcuAnt(9) & "+" & dsTablaCCo & "AcuD13_MN"
                                             gsAcuAnt(10) = gsAcuAnt(10) & "+" & dsTablaCCo & "AcuD13_ME"
                                             gsAcuAnt(11) = gsAcuAnt(11) & "+" & dsTablaCCo & "AcuH13_MN"
                                             gsAcuAnt(12) = gsAcuAnt(12) & "+" & dsTablaCCo & "AcuH13_ME"
            
                                             gsAcuAnt(1 + 12) = gsAcuAnt(1 + 12) & "+" & dsCi & dsTablaCta & "AcuD13_MN" & dsCF
                                             gsAcuAnt(2 + 12) = gsAcuAnt(2 + 12) & "+" & dsCi & dsTablaCta & "AcuD13_ME" & dsCF
                                             gsAcuAnt(3 + 12) = gsAcuAnt(3 + 12) & "+" & dsCi & dsTablaCta & "AcuH13_MN" & dsCF
                                             gsAcuAnt(4 + 12) = gsAcuAnt(4 + 12) & "+" & dsCi & dsTablaCta & "AcuH13_ME" & dsCF
   
                                             gsAcuAnt(5 + 12) = gsAcuAnt(5 + 12) & "+" & dsCi & dsTablaAux & "AcuD13_MN" & dsCF
                                             gsAcuAnt(6 + 12) = gsAcuAnt(6 + 12) & "+" & dsCi & dsTablaAux & "AcuD13_ME" & dsCF
                                             gsAcuAnt(7 + 12) = gsAcuAnt(7 + 12) & "+" & dsCi & dsTablaAux & "AcuH13_MN" & dsCF
                                             gsAcuAnt(8 + 12) = gsAcuAnt(8 + 12) & "+" & dsCi & dsTablaAux & "AcuH13_ME" & dsCF
         
                                             gsAcuAnt(9 + 12) = gsAcuAnt(9 + 12) & "+" & dsCi & dsTablaCCo & "AcuD13_MN" & dsCF
                                             gsAcuAnt(10 + 12) = gsAcuAnt(10 + 12) & "+" & dsCi & dsTablaCCo & "AcuD13_ME" & dsCF
                                             gsAcuAnt(11 + 12) = gsAcuAnt(11 + 12) & "+" & dsCi & dsTablaCCo & "AcuH13_MN" & dsCF
                                             gsAcuAnt(12 + 12) = gsAcuAnt(12 + 12) & "+" & dsCi & dsTablaCCo & "AcuH13_ME" & dsCF
   
'[REVISAR: Ver si es útil.
'gsIngAntCtaRpt = gsIngAntCtaRpt & "+" & dsCi & dsTablaCta & "CtaI11" & dsCF
'gsSalAntCtaRpt = gsSalAntCtaRpt & "+" & dsCi & dsTablaCta & "CtaS11" & dsCF
']REVISAR.
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      For dnContador = 1 To 24
         gsAcuAnt(dnContador) = gsAcuAnt(dnContador) & ")"
'         gsIniMes(dnContador) = gsIniMes(dnContador) & "+" & gsAcuAnt(dnContador)
      Next
   End If
End Sub

Sub gpCieMes()
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
    .Source = "SELECT IndCpr, IndVta, IndHpr, IndCpb "
    .Source = .Source & "FROM COCieMes "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND mescie='" & gsMesAct & "'"
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
   
    gbCieCpr = !IndCpr
    gbCieVta = !IndVta
    gbCieHpr = !IndHpr
    gbCieCpb = !IndCpb
    .Close
   End With

   docnnMain.Close
   Set dorstMain = Nothing
   Set docnnMain = Nothing
End Sub


