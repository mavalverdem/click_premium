Attribute VB_Name = "modAyuBus"
Option Explicit

Public Sub Asi_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT codasi, " & Choose(gsIdioma, "detasi", "detasix") & " AS detasi "
      .usConnStrgSele = .usConnStrgSele & "FROM coasitipo "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Codigo", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6000)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codasi='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6030 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub
Public Sub Med_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT codmed, abvmed,desmed,indmod "
      .usConnStrgSele = .usConnStrgSele & "FROM bnmediopago "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgOrde = " ORDER BY 1 "
      .uaTitulos = Array(Choose(gsIdioma, "Codigo", "Code"), Choose(gsIdioma, "Abreviatura", "Abreviatura"), Choose(gsIdioma, "Descripción", "Description"), Choose(gsIdioma, "Indicador", "Indicador"))
      .uaAncho = Array(800, 1000, 5000, 500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "", "", "")
      .uaOrden = Array("", "", "", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codmed='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6030 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .uvDato3Posicion = 2
      .uvDato4Posicion = 3
      .unElementos = 4
      
      .Show vbModal
   End With
End Sub
Public Sub Lib_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT codlib, deslib "
      .usConnStrgSele = .usConnStrgSele & "FROM colib "
      .usConnStrgOrde = " ORDER BY 1 "
      .uaTitulos = Array(Choose(gsIdioma, "Codigo", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6000)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codlib='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6030 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub
Public Sub Aux_Det(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT RazAux, CodAux "
      .usConnStrgSele = .usConnStrgSele & "FROM TGAux "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Razón Social", "Firm Name"), Choose(gsIdioma, "Código", "Code"))
      .uaAncho = Array(4000, 1100)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodAux='" & pvDato1Previo & "'"

      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 4030 + 1130 + 640)

      .uvDato1Posicion = 1
      .uvDato2Posicion = 0
      .unElementos = 2

      .Show vbModal
   End With
End Sub

Public Sub Bco_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT codbco, " & Choose(gsIdioma, "detbco", "detbcox") & " AS detbco, codent "
      .usConnStrgSele = .usConnStrgSele & "FROM cobco "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 4500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codbco='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 4530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Bco_CodBan(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frm0AyuBusBan
   
      .usConnStrgSele = "SELECT codbco, " & Choose(gsIdioma, "detbco", "detbcox") & " AS detbco,codent,forimp,ctactemn,ctacteme "
      .usConnStrgSele = .usConnStrgSele & "FROM cobco "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"), Choose(gsIdioma, "Entidad", "Entity"), Choose(gsIdioma, "Formato", "Format"), Choose(gsIdioma, "CtaCteMN", "CtaCteMN"), Choose(gsIdioma, "CtaCteMe", "CtaCteME"))
      .uaAncho = Array(800, 4500, 0, 0, 0, 0)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "", "", "", "", "")
      .uaOrden = Array("", "", "", "", "", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codbco='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 4530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .uvDato3Posicion = 2
      .uvDato4Posicion = 3
      .uvDato5Posicion = 4
      .uvDato6Posicion = 5
      .unElementos = 6
      
      .Show vbModal
      
   End With
End Sub



Public Sub Cta_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
      .usConnStrgSele = .usConnStrgSele & "FROM COCta "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Codigo", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodCta='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub CCo_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS DetCCo "
      .usConnStrgSele = .usConnStrgSele & "FROM CoCCo "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 4000)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodCCo='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 4030 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Cfg_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT DISTINCTROW codcfg, detcfg "
      .usConnStrgSele = .usConnStrgSele & "FROM coccocfg "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codcfg='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub DPe_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT coddpe, " & Choose(gsIdioma, "detdpe", "detdpex") & " AS detdpe "
      .usConnStrgSele = .usConnStrgSele & "FROM codpe "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      '.usConnStrgSele = .usConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "coddpe='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Dro_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
      .usConnStrgSele = .usConnStrgSele & "FROM CODro "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodDro='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Efe_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT CodEfe, " & Choose(gsIdioma, "DetEfe", "DetEfex") & " AS DetEfe "
      .usConnStrgSele = .usConnStrgSele & "FROM CoEfe "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodEfe='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub EFi_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .ubBDConfiguracion = False    'Para BD de Configuración.
      .usConnStrgSele = "SELECT CodEFi, " & Choose(gsIdioma, "DetEFi", "DetEFix") & " AS DetEFi "
      .usConnStrgSele = .usConnStrgSele & "FROM COEFi "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodEFi='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Emp_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .ubBDConfiguracion = True     'Para BD de Configuración.
      .usConnStrgSele = "SELECT CodEmp, RazEmp "
      .usConnStrgSele = .usConnStrgSele & "FROM TGEmp "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "WHERE " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Razón Social", "Firm Name"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodEmp='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub
Public Sub Emp_Usu(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .ubBDConfiguracion = True     'Para BD de Configuración.
      .usConnStrgSele = "SELECT DISTINCTROW emp.codemp, emp.razemp "
      .usConnStrgSele = .usConnStrgSele & "FROM tgemp emp, sgpms oxu "
      .usConnStrgSele = .usConnStrgSele & "WHERE oxu.codsis='CO' "
      .usConnStrgSele = .usConnStrgSele & "AND oxu.codemp=emp.codemp "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Razón Social", "Firm Name"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodEmp='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Fil_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .ubBDConfiguracion = False    'Para BD de Configuración.
      .usConnStrgSele = "SELECT codfil, " & Choose(gsIdioma, "detfil", "detfilx") & " AS detfil "
      .usConnStrgSele = .usConnStrgSele & "FROM cofil "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "codfil='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Fjo_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT CodFjo, " & Choose(gsIdioma, "DetFjo", "DetFjox") & " AS DetFjo "
      .usConnStrgSele = .usConnStrgSele & "FROM CoFjo "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(800, 6500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodFjo='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 6530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub TDc_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT CodTDc, " & Choose(gsIdioma, "DetTDc", "DetTDcx") & " AS DetTDc, AbvTDc "
      .usConnStrgSele = .usConnStrgSele & "FROM TGTDc "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"), Choose(gsIdioma, "Abreviado", "Abbreviated"))
      .uaAncho = Array(800, 4500, 800)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "", "")
      .uaOrden = Array("", "", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodTDc='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 830 + 4530 + 830 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 3
      
      .Show vbModal
   End With
End Sub

Public Sub Mdl_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .ubBDConfiguracion = True     'Para BD de Configuración.
      .usConnStrgSele = "SELECT CodMdl, " & Choose(gsIdioma, "DetMdl", "DetMdlx") & " AS DetMdl "
      .usConnStrgSele = .usConnStrgSele & "FROM SGMdl "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "WHERE " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(1500, 4800)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodMdl='" & pvDato1Previo & "'"

      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1530 + 4830 + 640)

      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2

      .Show vbModal
   End With
End Sub

Public Sub Usr_Cod(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .ubBDConfiguracion = True     'Para BD de Configuración.
      .usConnStrgSele = "SELECT CodUsr, NomUsr "
      .usConnStrgSele = .usConnStrgSele & "FROM SGUsr "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "WHERE " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY 1"
      .uaTitulos = Array(Choose(gsIdioma, "Código", "Code"), Choose(gsIdioma, "Descripción", "Description"))
      .uaAncho = Array(2500, 4200)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "CodUsr='" & pvDato1Previo & "'"

      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 2530 + 4230 + 640)

      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2

      .Show vbModal
   End With
End Sub

Public Sub Pdo_Cpr(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT concat(a.coddpe,a.pdocpr) as pdocpr, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo "
      .usConnStrgSele = .usConnStrgSele & "FROM copdocpr a "
      .usConnStrgSele = .usConnStrgSele & "WHERE a.codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & "AND a.pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND a.mespvs='" & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY concat(a.coddpe,a.pdocpr) "
      .uaTitulos = Array(Choose(gsIdioma, "Nº Pedido", "Nº Order"), Choose(gsIdioma, "Detalle", "Detail"))
      .uaAncho = Array(1200, 3000)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "pdocpr='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1200 + 3000 + 690)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 4
      
      .Show vbModal
   End With
End Sub
Public Sub Pdo_Rpt(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT concat(a.coddpe,a.pdocpr) as pdocpr, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo "
      .usConnStrgSele = .usConnStrgSele & "FROM copdocpr a "
      .usConnStrgSele = .usConnStrgSele & "WHERE a.codemp='" & gsCodEmp & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY concat(a.coddpe,a.pdocpr) "
      .uaTitulos = Array(Choose(gsIdioma, "Nº Pedido", "Nº Order"), Choose(gsIdioma, "Detalle", "Detail"))
      .uaAncho = Array(1200, 3000)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "pdocpr='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1200 + 3000 + 690)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 4
      
      .Show vbModal
   End With
End Sub
Public Sub Pdo_Sal(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
  With frmOAyuBus
    .usConnStrgSele = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.coddpe,a.pdocpr)", "(a.coddpe+a.pdocpr)") & " AS pdocpr, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo, a.tpomon, "
    If ps_Plataforma = pSrvMySql Then
      .usConnStrgSele = .usConnStrgSele & "ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN (a.impmn-"
      .usConnStrgSele = .usConnStrgSele & "IFNULL(ROUND(SUM("
      .usConnStrgSele = .usConnStrgSele & "((IFNULL(b.impogr_mn, 0)+IFNULL(b.impogn_mn, 0)+IFNULL(b.impong_mn, 0)+IFNULL(b.impexo_mn, 0))*"
      .usConnStrgSele = .usConnStrgSele & "(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(IFNULL(d.impbru_mn, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      .usConnStrgSele = .usConnStrgSele & "), 2), 0)) "
      .usConnStrgSele = .usConnStrgSele & "ELSE (a.impme-"
      .usConnStrgSele = .usConnStrgSele & "IFNULL(ROUND(SUM("
      .usConnStrgSele = .usConnStrgSele & "((IFNULL(b.impogr_me, 0)+IFNULL(b.impogn_me, 0)+IFNULL(b.impong_me, 0)+IFNULL(b.impexo_me, 0))*"
      .usConnStrgSele = .usConnStrgSele & "(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(IFNULL(d.impbru_me, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      .usConnStrgSele = .usConnStrgSele & "), 2), 0)) END, 2) AS cImpSaldo, "
    Else
      .usConnStrgSele = .usConnStrgSele & "ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN (a.impmn-"
      .usConnStrgSele = .usConnStrgSele & "ISNULL(ROUND(SUM("
      .usConnStrgSele = .usConnStrgSele & "((ISNULL(b.impogr_mn, 0)+ISNULL(b.impogn_mn, 0)+ISNULL(b.impong_mn, 0)+ISNULL(b.impexo_mn, 0))*"
      .usConnStrgSele = .usConnStrgSele & "(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(d.impbru_mn, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      .usConnStrgSele = .usConnStrgSele & "), 2), 0)) "
      .usConnStrgSele = .usConnStrgSele & "ELSE (a.impme-"
      .usConnStrgSele = .usConnStrgSele & "ISNULL(ROUND(SUM("
      .usConnStrgSele = .usConnStrgSele & "((ISNULL(b.impogr_me, 0)+ISNULL(b.impogn_me, 0)+ISNULL(b.impong_me, 0)+ISNULL(b.impexo_me, 0))*"
      .usConnStrgSele = .usConnStrgSele & "(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(d.impbru_me, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      .usConnStrgSele = .usConnStrgSele & "), 2), 0)) END, 2) AS cImpSaldo, "
    End If
    .usConnStrgSele = .usConnStrgSele & "a.fehpdo "
    .usConnStrgSele = .usConnStrgSele & "FROM ((((copdocpr a "
    .usConnStrgSele = .usConnStrgSele & "LEFT JOIN cocprdoc b ON a.codemp=b.codemp AND a.codaux=b.codaux AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.coddpe, a.pdocpr)", "(a.coddpe+a.pdocpr)") & "=b.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      .usConnStrgSele = .usConnStrgSele & "AND IFNULL(b.pdocpr, '')<>'' "
      .usConnStrgSele = .usConnStrgSele & "AND Concat(b.pdoano, b.mespvs)<='" & gsAnoAct & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND b.feedoc<=" & Right(Trim(psWhere), 12) & ") "
    ElseIf ps_Plataforma = pSrvSql Then
      .usConnStrgSele = .usConnStrgSele & "AND ISNULL(b.pdocpr, '')<>'' "
      .usConnStrgSele = .usConnStrgSele & "AND (b.pdoano+b.mespvs)<='" & gsAnoAct & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND b.feedoc<= CONVERT(smalldatetime, " & Right(Trim(psWhere), 12) & ", 103)) "
    End If
    .usConnStrgSele = .usConnStrgSele & "LEFT JOIN TGTDc c ON b.codemp=c.codemp AND b.CodTDc=c.CodTDc) "
    .usConnStrgSele = .usConnStrgSele & "LEFT JOIN cohprdoc d ON a.codemp=d.codemp AND a.codaux=d.codaux AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.coddpe, a.pdocpr)", "(a.coddpe+a.pdocpr)") & "=d.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      .usConnStrgSele = .usConnStrgSele & "AND IFNULL(d.pdocpr, '')<>'' "
      .usConnStrgSele = .usConnStrgSele & "AND Concat(d.pdoano, d.mespvs)<='" & gsAnoAct & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND d.feedoc<=" & Right(Trim(psWhere), 12) & ") "
    ElseIf ps_Plataforma = pSrvSql Then
      .usConnStrgSele = .usConnStrgSele & "AND ISNULL(d.pdocpr, '')<>'' "
      .usConnStrgSele = .usConnStrgSele & "AND (d.pdoano+d.mespvs)<='" & gsAnoAct & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND d.feedoc<= CONVERT(smalldatetime, " & Right(Trim(psWhere), 12) & ", 103)) "
    End If
    .usConnStrgSele = .usConnStrgSele & "LEFT JOIN TGTDc e ON d.codemp=e.codemp AND c.CodTDc='" & CODTDC_HPR & "') "
    .usConnStrgSele = .usConnStrgSele & "WHERE a.codemp='" & gsCodEmp & "' "
    .usConnStrgSele = .usConnStrgSele & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
    .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
    .usConnStrgSele = .usConnStrgSele & "GROUP BY a.codemp, a.codaux, a.coddpe, a.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      .usConnStrgSele = .usConnStrgSele & "HAVING cImpSaldo<>0.00 "
    Else
      .usConnStrgSele = .usConnStrgSele & "HAVING (ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN (a.impmn-"
      .usConnStrgSele = .usConnStrgSele & "ISNULL(ROUND(SUM((ISNULL(b.impogr_mn, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(b.impogn_mn, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(b.impong_mn, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(b.impexo_mn, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))), 2), 0)) ELSE "
      .usConnStrgSele = .usConnStrgSele & "(a.impme-ISNULL("
      .usConnStrgSele = .usConnStrgSele & "ROUND(SUM((ISNULL(b.impogr_me, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(b.impogn_me, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(b.impong_me, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      .usConnStrgSele = .usConnStrgSele & "(ISNULL(b.impexo_me, 0)*(CASE c.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))), 2), 0)) END, 2))<>0.00 "
    End If
    .usConnStrgOrde = "ORDER BY a.pdocpr"
    .uaTitulos = Array(Choose(gsIdioma, "Nº Pedido", "Nº Order"), Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Mon", "Cur"), Choose(gsIdioma, "Saldo", "Rest"))
    .uaAncho = Array(1100, 2000, 250, 1200)
    .uaAlineamiento = Array(dbgGeneral, dbgGeneral, dbgCenter, dbgRight)
    .uaFormato = Array("", "", "", FORMATO_NUM_1 & " ")
    .uaOrden = Array("", "", "", "")
    .uvDato1Previo = pvDato1Previo
    .usCriterio = "pdocpr='" & pvDato1Previo & "'"
    
    .unArribaFormulario = pnArriba + 350
    .unIzquierdaFormulario = pnIzquierda + 50
    .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
    .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 3100 + 1450 + 640)
    
    .uvDato1Posicion = 0
    .uvDato2Posicion = 1
    .unElementos = 4
    
    .Show vbModal
  End With
End Sub

Public Sub Sal_Doc(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
  With frmOAyuBus
    .usConnStrgSele = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodTDc, a.SerDoc, a.NroDoc)", "(a.CodTDc + a.SerDoc + a.NroDoc)") & " AS cDocume, "
    .usConnStrgSele = .usConnStrgSele & "(CASE b.TpoMon WHEN '" & TPOMON_NAC & "' THEN a.ImpSMN ELSE a.ImpSME END) AS cImpSaldo, "
    .usConnStrgSele = .usConnStrgSele & "a.CodTDc "
    .usConnStrgSele = .usConnStrgSele & "FROM (CoDocTmp1 a "
    .usConnStrgSele = .usConnStrgSele & "LEFT JOIN CoDocTmp2 b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.codcta=b.codcta AND a.codaux=b.codaux AND a.codtdc=b.codtdc AND a.serdoc=b.serdoc AND a.nrodoc=b.nrodoc AND a.usrcre=b.usrcre) "
    .usConnStrgSele = .usConnStrgSele & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
    .usConnStrgOrde = "ORDER BY a.codcta, a.CodTDc, a.SerDoc, a.NroDoc"
    .uaTitulos = Array(Choose(gsIdioma, "Documento", "Document"), Choose(gsIdioma, "Saldo", "Rest"))
    .uaAncho = Array(2000, 1500)
    .uaAlineamiento = Array(dbgGeneral, dbgRight)
    .uaFormato = Array("", FORMATO_NUM_1 & " ")
    .uaOrden = Array("", "")
    .uvDato1Previo = pvDato1Previo
    .usCriterio = "cDocume='" & pvDato1Previo & "'"
    
    .unArribaFormulario = pnArriba + 350
    .unIzquierdaFormulario = pnIzquierda + 50
    .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 2950)
    .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 2030 + 1530 + 640)
    
    .uvDato1Posicion = 0
    .uvDato2Posicion = 1
    .unElementos = 2
    
    .Show vbModal
  End With
End Sub

Public Sub Sel_Doc(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
  With frmOSelPen
    .usConnStrgSele = "SELECT codoctmp1.codcta, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(codoctmp1.CodTDc, codoctmp1.SerDoc, codoctmp1.NroDoc)", "(codoctmp1.CodTDc + codoctmp1.SerDoc + codoctmp1.NroDoc)") & " AS cDocume, "
    .usConnStrgSele = .usConnStrgSele & "(CASE b.tpomon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cTpoMon, "
    .usConnStrgSele = .usConnStrgSele & "(CASE b.tpomon WHEN '" & TPOMON_NAC & "' THEN codoctmp1.ImpSMN ELSE codoctmp1.ImpSME END) AS cImpSaldo, "
    .usConnStrgSele = .usConnStrgSele & "codoctmp1.imppmn, codoctmp1.imppme, codoctmp1.codcco, codoctmp1.indsel, b.tpomon, codoctmp1.codtdc, "
    .usConnStrgSele = .usConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "Concat(codoctmp1.codcta,codoctmp1.codtdc,codoctmp1.serdoc,codoctmp1.nrodoc)", "(codoctmp1.codcta+codoctmp1.codtdc+a.serdoc+codoctmp1.nrodoc)") & " AS cLlave "
    .usConnStrgSele = .usConnStrgSele & "FROM (codoctmp1 "
    .usConnStrgSele = .usConnStrgSele & "LEFT JOIN codoctmp2 b ON codoctmp1.codemp=b.codemp AND codoctmp1.pdoano=b.pdoano AND codoctmp1.codcta=b.codcta AND codoctmp1.codaux=b.codaux AND codoctmp1.codtdc=b.codtdc AND codoctmp1.serdoc=b.serdoc AND codoctmp1.nrodoc=b.nrodoc AND codoctmp1.usrcre=b.usrcre AND codoctmp1.fyhcre=b.fyhcre) "
    .usConnStrgSele = .usConnStrgSele & "WHERE codoctmp1.codemp='" & gsCodEmp & "' AND codoctmp1.pdoano='" & gsAnoAct & "' "
    .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
    .usConnStrgOrde = "ORDER BY codoctmp1.codcta, codoctmp1.codtdc, codoctmp1.serdoc, codoctmp1.nrodoc"
    .uaTitulos = Array(Choose(gsIdioma, "Cuenta", "Account"), Choose(gsIdioma, "Documento", "Document"), Choose(gsIdioma, "Mon", "Curr"), Choose(gsIdioma, "Saldo", "Rest"), Choose(gsIdioma, "Importe MN", "Amount NC"), Choose(gsIdioma, "Importe ME", "Amount FC"), Choose(gsIdioma, "C.Cos.", "C.Cen"), "Sel")
    .uaAncho = Array(850, 1550, 400, 1000, 1000, 1000, 550, 300)
    .uaAlineamiento = Array(flexAlignLeftCenter, flexAlignLeftCenter, dbgCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, dbgGeneral, dbgCenter)
    .uaFormato = Array("", "", "", FORMATO_NUM_1 & " ", FORMATO_NUM_1 & " ", FORMATO_NUM_1 & " ", "", "")
    .uaOrden = Array("", "", "", "", "", "", "", "")
    .uvDato1Previo = pvDato1Previo
    .usCriterio = "cDocume='" & pvDato1Previo & "'"
    
    .unArribaFormulario = pnArriba + 350
    .unIzquierdaFormulario = pnIzquierda + 50
    .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 3000)
    .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 5000)
    
    .uvDato1Posicion = 0
    .uvDato2Posicion = 1
    .unElementos = 8
    
    .Show vbModal
  End With
End Sub

Public Sub Doc_Cpr(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc,'-', NroDoc)", "(SerDoc + '-' +  NroDoc)") & " AS cDocumento, "
      .usConnStrgSele = .usConnStrgSele & "GloDoc, SerDoc, NroDoc "
      .usConnStrgSele = .usConnStrgSele & "FROM CoCprDoc "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND MesPvs='" & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY SerDoc, NroDoc"
      .uaTitulos = Array(Choose(gsIdioma, "Documento", "Document"), Choose(gsIdioma, "Glosa", "Gloss"))
      .uaAncho = Array(1500, 2500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "cDocumento='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 3650)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1530 + 2530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Doc_Vta(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc,'-', NroDoc)", "(SerDoc + '-' +  NroDoc)") & " AS cDocumento, "
      .usConnStrgSele = .usConnStrgSele & "GloDoc, SerDoc, NroDoc "
      .usConnStrgSele = .usConnStrgSele & "FROM CoVtaDoc "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND MesPvs='" & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY SerDoc, NroDoc"
      .uaTitulos = Array(Choose(gsIdioma, "Documento", "Document"), Choose(gsIdioma, "Glosa", "Gloss"))
      .uaAncho = Array(1500, 2500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "cDocumento='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 3650)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1530 + 2530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub

Public Sub Doc_Hpr(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc,'-', NroDoc)", "(SerDoc + '-' +  NroDoc)") & " AS cDocumento, "
      .usConnStrgSele = .usConnStrgSele & "GloDoc, SerDoc, NroDoc "
      .usConnStrgSele = .usConnStrgSele & "FROM CoHprDoc "
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND MesPvs='" & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY SerDoc, NroDoc"
      .uaTitulos = Array(Choose(gsIdioma, "Documento", "Document"), Choose(gsIdioma, "Glosa", "Gloss"))
      .uaAncho = Array(1500, 2500)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "cDocumento='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 3650)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1530 + 2530 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub
Public Sub Cpb_Dro(psWhere As String, pvDato1Previo As Variant, pnAlto, pnAncho, pnArriba As Integer, pnIzquierda As Integer)
   With frmOAyuBus
      .usConnStrgSele = "SELECT NroCpb, " & Choose(gsIdioma, "GloCpb", "GloCpbx") & " AS GloCpb "
      .usConnStrgSele = .usConnStrgSele & "FROM CoCpbCab"
      .usConnStrgSele = .usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .usConnStrgSele = .usConnStrgSele & "AND MesPvs='" & gsMesAct & "' "
      .usConnStrgSele = .usConnStrgSele & IIf(psWhere = "", "", "AND " & psWhere & " ")
      .usConnStrgOrde = "ORDER BY NroCpb"
      .uaTitulos = Array(Choose(gsIdioma, "Comprobante", "Voucher"), Choose(gsIdioma, "Glosa", "Gloss"))
      .uaAncho = Array(1000, 3000)
      .uaAlineamiento = Array(dbgGeneral, dbgGeneral)
      .uaFormato = Array("", "")
      .uaOrden = Array("", "")
      .uvDato1Previo = pvDato1Previo
      .usCriterio = "NroCpb='" & pvDato1Previo & "'"
      
      .unArribaFormulario = pnArriba + 350
      .unIzquierdaFormulario = pnIzquierda + 50
      .unAltoFormulario = IIf(pnAlto <> 0, pnAlto, 3650)
      .unAnchoFormulario = IIf(pnAncho <> 0, pnAncho, 1030 + 3030 + 640)
      
      .uvDato1Posicion = 0
      .uvDato2Posicion = 1
      .unElementos = 2
      
      .Show vbModal
   End With
End Sub
