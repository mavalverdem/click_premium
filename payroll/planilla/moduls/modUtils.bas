Attribute VB_Name = "modUtils"
Option Explicit
'''' Ricardo Malca - 02-02-2015
''' Librerìa con mètodos y funcinoes para manejar exportaciòn de CrystalReport a FormatoPDF
''' Para esto se utilizan las librerìas de CrXi 11, CRA ActvX Rntime 11 y RprtViewer

'Public oApp As CRAXDDRT.Application
'Public oRpt As CRAXDDRT.Report
'Public oExpOpc As CRAXDDRT.ExportOptions
'
'Private crpParamDefs As CRAXDDRT.ParameterFieldDefinitions
'Private crpParamDef As CRAXDDRT.ParameterFieldDefinition

Public Sub utlNew()
    Set oApp = New CRAXDDRT.Application
    Set oRpt = New CRAXDDRT.Report
End Sub

'carga los objeto [oApp - oRpt] con el reporte deseado
Public Sub utlSetCrpToPdf(pCrpName As String, Optional pFec As Date)
On Error GoTo Err
    
    utlNew
    
    'cargamos el reporte pCrpName
    Set oRpt = oApp.OpenReport(pCrpName & ".rpt", 1)
    
    'cfg
    With oRpt
        .DiscardSavedData
        .EnableParameterPrompting = False
    End With
    
    Set oExpOpc = oRpt.ExportOptions
    'cfg - pdf
    With oExpOpc
        .DestinationType = crEDTDiskFile
        .FormatType = crEFTPortableDocFormat
        .PDFExportAllPages = True
    End With
    
    Dim s_Fecha As String
    If Not IsDate(pFec) Then pFec = Now
    s_Fecha = gfDateText(pFec) & " / " & Format(Time(), "hh:mm:ss AMPM")
    utlSetParamToCrp s_Fecha
    
    Exit Sub
Err:
    MsgBox "Error en los paràmetros"
End Sub

Private Sub utlSetParamToCrp(pFecha As String)
On Error GoTo Err
    Dim iCnt As Integer
    Set crpParamDefs = oRpt.ParameterFields
    For iCnt = 1 To crpParamDefs.Count
        Set crpParamDef = crpParamDefs(iCnt)

        Select Case crpParamDef.ParameterFieldName
            Case "mSistema": setParamValue crpParamDef, gsNomSis
            Case "mTitulo": setParamValue crpParamDef, gsRazEmp
            Case "mFeReporte": setParamValue crpParamDef, pFecha
            Case "mPeriodo": setParamValue crpParamDef, gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1)
            Case "mRucEmpresa": setParamValue crpParamDef, gsRUCEmp
        End Select
    Next
    
    Exit Sub
Err:
    MsgBox "Error en los paràmetros"
End Sub

Private Sub setParamValue(pParam As CRAXDDRT.ParameterFieldDefinition, pVal As Variant)
    pParam.ClearCurrentValueAndRange
    Select Case pParam.ValueType
        Case crDateField, crDateTimeField, crDateField
                pParam.AddCurrentValue CDate(pVal)
        Case crNumberField
                pParam.AddCurrentValue Val(pVal)
        Case Else
                pParam.AddCurrentValue pVal & ""
    End Select
End Sub

'exportar a crp -> pdf
Public Sub utlExpCrpToPdf(pRuta As String, pNombre As String, pRs As ADODB.Recordset)
On Error GoTo Err
    oRpt.DiscardSavedData
    oExpOpc.DiskFileName = pRuta & pNombre & ".pdf"
    oRpt.Database.SetDataSource pRs
    oRpt.DisplayProgressDialog = False
    oRpt.Export False
    oApp.CanClose
    Exit Sub
Err:
    MsgBox "Error al exportar"
End Sub

'para crviewer
Public Sub utlSetRpt(pRs As ADODB.Recordset)
    oRpt.DiscardSavedData
    oRpt.Database.SetDataSource pRs
End Sub

Public Property Get utlgetRpt() As CRAXDDRT.Report
    Set utlgetRpt = oRpt
End Property
