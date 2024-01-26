Attribute VB_Name = "modMainGV"
Option Explicit



Public Const CODTDC_FAC As String = "01", _
             CODTDC_HPR As String = "02", _
             CODTDC_BOL As String = "03", _
             CODTDC_NCR As String = "07", _
             CODTDC_TIC As String = "12", _
             CODCCO_AJD As String = "99099"
             
'Public Const xCODTDC_RET As String = "34", _
             xCODTDC_PCP As String = "35"
Public Const INDANU_VER As Integer = 1, _
             INDANU_FAL As Integer = 0, _
             ESTMDL_ACT As String = "A", _
             ESTMDL_INA As String = "I", _
             ESTUSR_ACT As String = "A", _
             ESTUSR_INA As String = "I"
Public ESTMDL_ACT_TXT As String, _
       ESTMDL_INA_TXT As String, _
       ESTUSR_ACT_TXT As String, _
       ESTUSR_INA_TXT As String
'Public Const INDCLI_SI As Boolean = True, _
             INDCLI_NO As Boolean = False, _
             SITCLI_ACT As String = "A", _
             SITCLI_INA As String = "I", _
             SITCLI_ACT_TXT As String = "Activo  ", _
             SITCLI_INA_TXT As String = "Inactivo"
'Public Const TPOMON_NAC As String = "N", _
             TPOMON_EXT As String = "E"

Public gnPctIGV As Single, gnPctIGV1 As Single, gnPctIGV2 As Single
Public gnPctISC As Single, _
       gnPctIR4 As Single, _
       gnPctIES As Single, _
       gnPctRtc As Single, _
       gnPctPcp As Single, _
       gnImpUIT As Single
       
Public gsCodTDc_Pcp As String, _
       gsCodTDc_Rtc As String, _
       gsCodCta_Pcp As String, _
       gsCodCta_Rtc As String
Public gsIndRtc As String           ' Indicador de agente de retención
Public gsIndPcp As String           ' Indicador de agente de percepción
Public gsTpoGlo_Rtc As String       ' Tipo de glosa de retencion
Public gsGloDoc_Rtc(2) As String    ' Arreglo de glosa de retencion

Public gsCodDro_Ing As String       ' Codigo de diario de caja ingreso
Public gsCodDro_Egr As String       ' Codigo de diario de caja egreso

Public Enum TipoImpuesto                   ' Tipo de tributo
  Ninguno = 0: ResponsableInscrito = 1: ResponsableMonotributo = 2: Exepto = 3: NoAlcanzado = 4: ConsumidosFinal = 5: TipoSemanal = 6: TipoQuincenal = 7: TipoMensual = 8
End Enum
Public Enum CategoriaDocumento             ' Categoria del docuemnto
  Ninguno = 0: ImpuestoDetallado = 1: FacturaPublica = 2: FacturaContador = 3: RetencionIva = 4: RetencionIB = 5: RetenconIG = 6: RetencionSuss = 7: RetencionOtro = 8
End Enum

Public TEXT_Ninguno  As String, TEXT_ResponsableInscrito As String, _
       TEXT_ResponsableMonotributo As String, TEXT_Exepto As String, _
       TEXT_NoAlcanzado As String, TEXT_ConsumidosFinal As String
Public TEXT_ImpuestoDetallado As String, TEXT_FacturaPublica As String, _
       TEXT_FacturaContador As String, TEXT_RetencionIva As String, _
       TEXT_RetencionIB As String, TEXT_RetencionIG As String, _
       TEXT_RetencionSuss As String, TEXT_RetencionOtro As String
