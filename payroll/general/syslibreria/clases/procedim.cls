VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Procedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                     ' Declarar variable antes de usarla

Public ps_PathImagen As String                      ' Directorio de las iconos, imagenes, etc.
Public pl_RecordSelector As Boolean                 ' Swith de selección de registros de grilla
Public ps_PathReport As String                      ' Directorio de los archivos de reportes
Public ps_ReportName As String                      ' Nombre de archivo de reporte
Public ps_ReportTitle As String                     ' Título de la ventana de reporte

Private s_Archivo As String                         ' Archivo de imagen, icono u otro tipo
Private s_ToolText As String                        ' Mensaje emergente de posición de mouse
Private s_Caption As String                         ' Caption de control
Private i As Integer                                ' Variables temporales y reutilizables
' [
Sub CentraFormulario(ByVal o_Form As Object)
  o_Form.Top = (IIf(o_Form.MDIChild, 7100, Screen.Height) - o_Form.Height) \ 2
  o_Form.Left = (Screen.Width - o_Form.Width) \ 2
End Sub
Sub DefineStyleGrilla(ByVal tdb As Object, ByVal s_TitleTable As String, ByVal n_Tipo As Integer)
    
With tdb
  ' Caracteristicas del titulo
  .Caption = s_TitleTable
  .CaptionStyle.Font.Name = "Tahoma"
  .CaptionStyle.Font.Bold = True
  .CaptionStyle.Font.Size = 8
  .CaptionStyle.BackColor = Choose(n_Tipo, 8421376, vbActiveTitleBar, 12632064)
  .CaptionStyle.ForeColor = vbWhite
  ' Color de titulo de columna
  .HeadBackColor = Choose(n_Tipo, 9088964, 8421376, &HFF0000)
  .HeadForeColor = Choose(n_Tipo, vbBlack, vbWhite, vbWhite)
  ' Texto de cabecera columna
  .HeadFont.Bold = True
  .HeadFont.Name = "Tahoma"
  ' Registros (Filas)
  .Font.Name = "Tahoma"
  .Font.Bold = False
  .Font.Size = 8
  .EvenRowStyle.BackColor = Choose(n_Tipo, 12311523, &H80000018, -2147483624)
  .EvenRowStyle.ForeColor = Choose(n_Tipo, vbBlack, vbWhite, vbBlack)
  .OddRowStyle.BackColor = Choose(n_Tipo, 13427690, &H80&, 13427690)
  .OddRowStyle.ForeColor = Choose(n_Tipo, vbBlack, vbWhite, vbBlack)
  .AlternatingRowStyle = Choose(n_Tipo, True, False, True)
  ' Color de fondo Cor celestone &H80000013
  .DeadAreaBackColor = &HC0C0C0
  ' Cursor de selección
  .HighlightRowStyle.BackColor = Choose(n_Tipo, 6908265, &H808080, 8421504)
  .HighlightRowStyle.ForeColor = vbWhite
  ' Registros seleccionados
  .SelectedBackColor = Choose(n_Tipo, 8421376, &H808080, 8421376)
  .SelectedForeColor = Choose(n_Tipo, vbBlack, vbWhite, vbBlack)
  ' División entre columnas
  .RowDividerStyle = dbgInset
  ' Presentación de barra de desplazamiento
  .MarqueeStyle = dbgHighlightCell
  .Appearance = Choose(n_Tipo, dbg3D, dbg3D, dbgTrack3D)
  ' Mensaje de Tooltips  en las celdas
  .CellTips = dbgFloating
  ' Movilizar columnas
  .AllowColMove = True
  .Splits(0).RecordSelectors = Choose(n_Tipo, pl_RecordSelector, False, pl_RecordSelector)
End With

End Sub
Sub DefineStyleGrillaDrop(tdb As Object)
    
With tdb
  ' Color de titulo de columna
  .HeadBackColor = 8421376
  .HeadForeColor = vbWhite
  ' Texto de cabecera columna
  .HeadFont.Bold = True
  .HeadFont.Name = "Tahoma"
  ' Registros (Filas)
  .Font.Name = "Tahoma"
  .Font.Bold = False
  .Font.Size = 8.25
  .BackColor = n_BackColorHelp#
  .ForeColor = vbBlack
  .EvenRowStyle.BackColor = n_BackColorHelp#
  .EvenRowStyle.ForeColor = vbWhite
  .OddRowStyle.BackColor = &H80&
  .OddRowStyle.ForeColor = vbWhite
  .AlternatingRowStyle = False
  ' Color de fondo Cor celestone &H80000013
  .DeadAreaBackColor = &H80000013
  ' Cursor de selección
  .HighlightRowStyle.BackColor = &H808080
  .HighlightRowStyle.ForeColor = vbWhite
  ' División entre columnas
  .RowDividerStyle = dbgInset
  ' Movilizar columnas
  .AllowColMove = False
End With

End Sub
Sub EditCombo(ByVal s_Key As String, ByVal o_Ctl As Object, ByVal n_Dato As Integer, ByVal s_Modo As String, ByVal l_Edit As Boolean)

' Establece el modo de edición de un control
Select Case s_Modo
 Case "A", "B", "V"
  With o_Ctl
    If .Style <> vbComboDropdownList Then .Text = ""
    .Locked = IIf(s_Modo = "A", l_Edit, True)
    .BackColor = IIf(s_Modo = "A", n_BackColorActive#, n_BackColorInactive#)
    .ForeColor = IIf(s_Modo = "A", n_ForeColorActive#, n_ForeColorInactive#)
    .ListIndex = n_Dato
  End With
 Case "C"
  With o_Ctl
    .Locked = IIf(s_Key = "PK", True, l_Edit)
    .BackColor = IIf(s_Key = "PK", n_BackColorInactive#, n_BackColorActive#)
    .ForeColor = IIf(s_Key = "PK", n_ForeColorInactive#, n_ForeColorActive#)
    .ListIndex = n_Dato
  End With
End Select

End Sub
Sub EditDTPicker(ByVal s_Key As String, ByVal o_Ctl As Object, ByVal s_Dato As String, ByVal s_Modo As String, ByVal l_Edit As Boolean, ByVal s_Formato As String, ByVal n_TipoFormat As Integer)

' Establece el modo de edición de un control
Select Case s_Modo
 Case "A", "B", "V"
    o_Ctl.Enabled = IIf(s_Modo = "A", l_Edit, False)
 Case "C"
    o_Ctl.Enabled = IIf(s_Key = "PK", False, l_Edit)
End Select
' Formato del control
With o_Ctl
    .Format = n_TipoFormat
    .CustomFormat = s_Formato
    .Value = s_Dato
    .CalendarForeColor = n_ForeColorActive#
    .CalendarTitleBackColor = &H808000
    .CalendarTrailingForeColor = &H80&
End With

End Sub
Sub EditMask(ByVal s_Key As String, ByVal o_Ctl As Object, ByVal s_Dato As String, ByVal s_Modo As String, ByVal l_Edit As Boolean, ByVal s_Mask As String)

' Establece el modo de edición de un control
Select Case s_Modo
 Case "A", "B", "V"
    With o_Ctl
        .Enabled = IIf(s_Modo = "A", l_Edit, False)
        .BackColor = IIf(s_Modo = "A", n_BackColorActive#, n_BackColorInactive#)
        .ForeColor = IIf(s_Modo = "A", n_ForeColorActive#, n_ForeColorInactive#)
    End With
 Case "C"
    With o_Ctl
        .Enabled = IIf(s_Key = "PK", False, l_Edit)
        .BackColor = IIf(s_Key = "PK", n_BackColorInactive#, n_BackColorActive#)
        .ForeColor = IIf(s_Key = "PK", n_ForeColorInactive#, n_ForeColorActive#)
    End With
End Select
' Elimino y adiciono el formato
o_Ctl.Mask = ""
o_Ctl = s_Dato
o_Ctl.Mask = s_Mask
    
End Sub
Sub EditOptionCheck(ByVal s_Key As String, ByVal o_Ctl As Object, ByVal l_Dato As Boolean, ByVal s_Modo As String, ByVal l_Edit As Boolean)

' Establece el modo de edición de un control
Select Case s_Modo
 Case "A", "B", "V"
    With o_Ctl
        .Enabled = IIf(s_Modo = "A", l_Edit, False)
        .ForeColor = IIf(s_Modo = "A", n_ForeColorActive#, n_ForeColorInactive#)
        .Value = l_Dato
    End With
 Case "C"
    With o_Ctl
        .Enabled = IIf(s_Key = "PK", False, l_Edit)
        .ForeColor = IIf(s_Key = "PK", n_ForeColorInactive#, n_ForeColorActive#)
        .Value = l_Dato
    End With
End Select

End Sub
Sub EditText(ByVal s_Key As String, ByVal o_Ctl As Object, ByVal s_Dato As String, ByVal s_Modo As String, ByVal l_Edit As Boolean, ByVal n_DefinedSize As Integer, Optional ByVal n_Alineacion As Byte, Optional ByVal s_Formato As String)

' Establece el modo de edición de un control
Select Case s_Modo
 Case "A", "B", "V"
    With o_Ctl
        .Locked = IIf(s_Modo = "A", l_Edit, True)
        .BackColor = IIf(s_Modo = "A", n_BackColorActive#, n_BackColorInactive#)
        .ForeColor = IIf(s_Modo = "A", n_ForeColorActive#, n_ForeColorInactive#)
    End With
 Case "C"
    With o_Ctl
        .Locked = IIf(s_Key = "PK", True, l_Edit)
        .BackColor = IIf(s_Key = "PK", n_BackColorInactive#, n_BackColorActive#)
        .ForeColor = IIf(s_Key = "PK", n_ForeColorInactive#, n_ForeColorActive#)
    End With
End Select
o_Ctl.MaxLength = n_DefinedSize
o_Ctl.Text = s_Dato
o_Ctl.Alignment = n_Alineacion
If s_Formato <> "" Then o_Ctl.Text = Format(o_Ctl.Text, s_Formato)

End Sub
Sub InicializaGrilla(ByVal tdb As Object, ByVal aElemento, ByVal aElementos)
Static i%, n_Limite%, n_Index%, n%, n_Limites%
Static o_Col As Variant
Static o_Split As TrueOleDBGrid80.Split

' Elimina columnas indice 0 y split con indice 1, actuales
Do While tdb.Columns.Count <> 0: tdb.Columns.Remove 0: Loop
Do While tdb.Splits.Count <> 1: tdb.Splits.Remove 1: Loop

' Crea las columnas de acuerdo a los parametros
' titulo, campo, ancho, alineación horizontal, formato, modificable, bloqueado, dropdow, alineación vertical, visible
n_Limite% = (UBound(aElemento, 1) - 1)
For i% = 0 To n_Limite%
    o_Col = tdb.Columns.Add(i%)
    tdb.Columns(i%).Visible = False
    tdb.Columns(i%).Caption = aElemento(i%, 0)
    tdb.Columns(i%).DataField = aElemento(i%, 1)
    tdb.Columns(i%).Width = aElemento(i%, 2)
    tdb.Columns(i%).HeadAlignment = aElemento(i%, 3)
    tdb.Columns(i%).Alignment = aElemento(i%, 3)
    tdb.Columns(i%).NumberFormat = aElemento(i%, 4)
    tdb.Columns(i%).Merge = aElemento(i%, 5)        ' Si las filas se combina
    tdb.Columns(i%).Locked = aElemento(i%, 6)       ' Bloquea para escritura
    tdb.Columns(i%).DropDown = aElemento(i%, 7)     ' si muestra combobox de ayuda
Next i%

' Crea los split de acuerdo a los parametros
n_Index% = (UBound(aElemento, 2) - 1)
n_Limites% = (UBound(aElementos, 1) - 1)
For i% = 0 To n_Limites%
    tdb.Splits(i%).Caption = aElementos(i%, 0)
    tdb.Splits(i%).CaptionStyle.Font.Name = "Tahoma"
    tdb.Splits(i%).CaptionStyle.Font.Size = 8
    tdb.Splits(i%).CaptionStyle.Font.Bold = True
    tdb.Splits(i%).CaptionStyle.BackColor = 5792539
    tdb.Splits(i%).CaptionStyle.ForeColor = vbWhite
    tdb.Splits(i%).CaptionStyle.Alignment = vbCenter
    tdb.Splits(i%).BackColor = aElementos(i%, 1)
    tdb.Splits(i%).ForeColor = aElementos(i%, 2)
    tdb.Splits(i%).RecordSelectors = False
    tdb.Splits(i%).Size = dbgNumberOfColumns
    ' Personalizo las columnas y caracteristicas especiales
    For n% = 0 To n_Limite%
        tdb.Splits(i%).Columns(n%).Visible = (aElemento(n%, n_Index%) = i%)
        tdb.Splits(i%).Columns(n%).Style.VerticalAlignment = aElemento(n%, n_Index% - 1)
    Next n%
    If i% = n_Limites% Then Exit For
    Set o_Split = tdb.Splits.Add(i%)
Next i%

End Sub
Private Sub InicializaParameters(ByVal o_CryReport As Object, ByVal n_IniForm As Integer, ByVal n_IniPara As Integer, ByVal n_IniFields As Integer)
Dim n_Index As Integer
    
' Limpo las Fórmulas Restantes
For n_Index = n_IniForm To 40
  o_CryReport.Formulas(n_Index) = ""
Next n_Index
' Limpio los Parámetros Restantes del Stored Procedure
For n_Index = n_IniPara To 20
  o_CryReport.StoredProcParam(n_Index) = ""
Next n_Index
' Limpio los Campos Parámetros Restantes
For n_Index = n_IniFields To 40
  o_CryReport.ParameterFields(n_Index) = ""
Next n_Index
' Limpio los Filtros Restantes
o_CryReport.GroupSelectionFormula = ""
o_CryReport.SelectionFormula = ""

End Sub
Sub LoadGrafics(ByVal o_Ctl As Object, ByVal s_IconFile As String, ByVal s_ToolTipText As String, Optional ByVal s_Caption As String)

' Instancio el objeto de funciones
Set gdl_Funcion = CreateObject("syslink.funciones")

o_Ctl.Picture = LoadPicture()
s_IconFile = ps_PathImagen & s_IconFile & ".bmp"
If gdl_Funcion.ExisteArchivo(s_IconFile) Then
  o_Ctl.Picture = LoadPicture(s_IconFile)
End If
o_Ctl.ToolTipText = s_ToolTipText
o_Ctl.Caption = s_Caption
o_Ctl.Refresh

Set gdl_Funcion = Nothing

End Sub
Sub MarcaGet(ByVal o_Ctl As Object)

o_Ctl.SelStart = 0
o_Ctl.SelLength = Len(o_Ctl)

End Sub
Sub MarcaRegistros(ByVal dca As Object, ByVal tdb As Object, ByVal v_Inicio As Variant, ByVal v_Final As Variant, ByVal n_Accion As Integer, ByVal s_Mensaje As String)

Dim v_Registro As Variant

If dca.Recordset.RecordCount = 0 Then
    Beep
    MsgBox "No Existen " & s_Mensaje & " para Seleccionar", vbExclamation
    Exit Sub
End If
        
' obtengo el registro actual
v_Registro = dca.Recordset.Bookmark
    
' verifico si se seleccionó la acción(inicio, final o inicializa)
If n_Accion = 0 Then
    If v_Registro > v_Final Then
        Beep
        MsgBox "El Inicio del Rango Seleccionado No puede ser Mayor que el Final del mismo" & _
                vbCr & "Rango de Inicio Actual : " & v_Inicio & _
                vbCr & "Rango de Fin Actual    : " & v_Final, vbExclamation
        Exit Sub
    End If
    v_Inicio = v_Registro
    GoTo MarcaRegistros
ElseIf n_Accion% = 1 Then
    If v_Registro < v_Inicio Then
        Beep
        MsgBox "El Final del Rango Seleccionado No puede ser Menor que el Inicio del mismo" & _
               vbCr & "Rango de Inicio Actual : " & v_Inicio & _
               vbCr & "Rango de Fin Actual    : " & v_Final, vbExclamation
        Exit Sub
    End If
    v_Final = v_Registro
    GoTo MarcaRegistros
ElseIf n_Accion% = 2 Then
    GoTo LimpiaRegistros
End If
    
MarcaRegistros:
    ' coloco el puntero en espera
    PunteroEnEspera
    ' inicializo el rango de selección
    While tdb.SelBookmarks.Count > 0
        tdb.SelBookmarks.Remove 0
    Wend
    ' creo el clone y lleno el rango de selección
    dca.Recordset.Bookmark = v_Inicio
    Do While dca.Recordset.Bookmark <= v_Final
        tdb.SelBookmarks.Add dca.Recordset.Bookmark
        dca.Recordset.MoveNext
        If dca.Recordset.EOF Then
            dca.Recordset.MovePrevious
            Exit Do
        End If
    Loop
    ' coloco el puntero normal
    PunteroNormal
    Exit Sub
    
LimpiaRegistros:
    ' coloco el puntero en espera
    PunteroEnEspera
    ' inicializo el rango de selección
    While tdb.SelBookmarks.Count > 0
        tdb.SelBookmarks.Remove 0
    Wend
    ' creo el clone y lleno el rango de seleccion
    dca.Recordset.MoveLast
    v_Final = dca.Recordset.Bookmark
    dca.Recordset.MoveFirst
    v_Inicio = dca.Recordset.Bookmark
    ' coloco el puntero normal
    PunteroNormal

End Sub
Sub panToolPosicion(ByVal panToolBar As Object, ByVal panTool As Object, ByVal cmd As Object, ByVal n_IndexTool As Integer, ByVal n_Index As Integer)
Dim n_NewPosition As Integer, n_EndPosition As Integer
    
For i% = 0 To panTool.Count - 1
    panTool(i%).ForeColor = &H0&
Next i%
' Setea Color de Fuente
panTool(n_Index).ForeColor = &HFF0000
    
If n_IndexTool <> n_Index Then
    n_NewPosition% = cmd(0).Left
    n_EndPosition% = panToolBar.Width
    While n_NewPosition% <= n_EndPosition%
        n_NewPosition% = n_NewPosition% + 100
        For i% = 0 To cmd.Count - 1
            cmd(i%).Left = n_NewPosition%
        Next i%
        panToolBar.Refresh
    Wend
End If

'Actualizar Propiedad Visible
If n_IndexTool <> n_Index Then
    For i% = 0 To cmd.Count - 1
        cmd(i%).Visible = (n_Index = cmd(i%).Tag)
    Next i%
End If

If n_IndexTool <> n_Index Then
    n_NewPosition% = 1
    n_EndPosition% = (panToolBar.Width - cmd(0).Width) / 2
    While n_NewPosition% <= n_EndPosition%
        n_NewPosition% = n_NewPosition% + 100
        For i% = 0 To cmd.Count - 1
            cmd(i%).Left = n_NewPosition%
        Next i%
        panToolBar.Refresh
    Wend
End If

End Sub
Sub ParametersExportar(ByVal s_Conexion As String, ByVal o_CryReport As Object, ByVal n_Destination As Integer, ByVal l_AllowDrillDown As Boolean, ByVal l_CloseBtn As Boolean, ByVal l_GroupTree As Boolean, ByVal l_SetupBtn As Boolean, ByVal l_SearchBtn As Boolean, ByVal aFormulaStore, ByVal aFilterSelection, Optional o_Data As Object, Optional s_Subreports As String, Optional s_SqlQuery As String, Optional s_RutaTempo As String, Optional n_Origen As Integer)
  Dim n_Dimension As Integer
  Dim n_Secuencia As Integer
  
  On Error GoTo ErrorPrinter
    
  ' Prepara Impresión
  
  
'      Dim poCrystalAplicacion As CRAXDRT.Application
'      Dim poCrystalReporte As CRAXDRT.Report
'
'      Set poCrystalAplicacion = New CRAXDRT.Application
'      Set poCrystalReporte = poCrystalAplicacion.OpenReport( ps_PathReport & ps_ReportName & ".rpt", 1)
'      With poCrystalReporte
'        .Database.SetDataSource porstRecordset, 3
'        With .ExportOptions
'          .DestinationType = crEDTDiskFile
'          .DiskFileName = "d:\boleta.pdf"
'          .PDFExportAllPages = True
'          .FormatType = crEFTPortableDocFormat
'          .UseReportDateFormat = True
'          .UseReportNumberFormat = True
'        End With
'        .EnableParameterPrompting = False
'        .Export False
'      End With
'      Set poCrystalAplicacion = Nothing
'      Set poCrystalReporte = Nothing
  
  
  With o_CryReport
    .Reset
    .Destination = n_Destination
    .ReportFileName = ps_PathReport & ps_ReportName & ".rpt"
    .PrinterPort = Printer.Port
    .PrinterDriver = Printer.DriverName
    .PrinterName = Printer.DeviceName
    .WindowAllowDrillDown = l_AllowDrillDown
    .WindowShowCloseBtn = l_CloseBtn
    .WindowShowGroupTree = l_GroupTree
    .WindowShowPrintSetupBtn = l_SetupBtn
    .WindowShowSearchBtn = l_SearchBtn
    .WindowTitle = ps_ReportTitle
    .BoundReportHeading = ps_ReportTitle
    .DataFiles(0) = s_RutaTempo
    .SetTablePrivateData 0, 3, o_Data
    .DiscardSavedData = False
    .SQLQuery = s_SqlQuery
    .Connect = Choose(n_Origen + 1, s_Conexion, "")
    If s_Subreports <> "" Then
      .SubreportToChange = s_Subreports
      .SetTablePrivateData 0, 3, o_Data
      .SubreportToChange = ""
    End If
  End With
  n_Dimension = UBound(aFormulaStore, 2)
  For n_Secuencia = 0 To (n_Dimension - 1)
    ' o_CryReport.StoredProcParam(n_Secuencia ) = aFormulaStore(0, n_Secuencia )
    o_CryReport.Formulas(n_Secuencia) = aFormulaStore(1, n_Secuencia)
    o_CryReport.ParameterFields(n_Secuencia) = aFormulaStore(2, n_Secuencia)
  Next n_Secuencia
  ' Inicializa las Formulas y Parametros Restantes
  InicializaParameters o_CryReport, n_Dimension, n_Dimension, n_Dimension
  ' Filtros Restantes adicionales
  o_CryReport.SelectionFormula = aFilterSelection(0)
  o_CryReport.GroupSelectionFormula = aFilterSelection(1)
  
  ' Ejecuto la Impresión
  o_CryReport.Action = 1
  Exit Sub

ErrorPrinter:
  Select Case Err.Number
   Case 20514      'Archivo ya Existe.
    MsgBox "Ya existe un archivo con este nombre en el Disco, utilice otro para el nombrarlo", vbExclamation + vbOKOnly, ps_ReportTitle
    Exit Sub
   Case 20545      'Cancelar
    MsgBox "La Impresión se Interrumpio i/o Cancela; Verifique", vbExclamation + vbOKOnly, ps_ReportTitle
    Exit Sub
   Case Else
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, ps_ReportTitle
    Exit Sub
  End Select

End Sub
Sub ParametersPrinter(ByVal s_Conexion As String, ByVal o_CryReport As Object, ByVal n_Destination As Integer, ByVal l_AllowDrillDown As Boolean, ByVal l_CloseBtn As Boolean, ByVal l_GroupTree As Boolean, ByVal l_SetupBtn As Boolean, ByVal l_SearchBtn As Boolean, ByVal aFormulaStore, ByVal aFilterSelection, Optional o_Data As Object, Optional s_Subreports As String, Optional s_SqlQuery As String, Optional s_RutaTempo As String, Optional n_Origen As Integer)
  Static n_Dimension As Integer
  
  On Error GoTo ErrorPrinter
    
  ' Prepara Impresión
  With o_CryReport
      .Reset
      .Destination = n_Destination%
      .ReportFileName = ps_PathReport & ps_ReportName & ".rpt"
      .PrinterPort = Printer.Port
      .PrinterDriver = Printer.DriverName
      .PrinterName = Printer.DeviceName
      .WindowAllowDrillDown = l_AllowDrillDown
      .WindowShowCloseBtn = l_CloseBtn
      .WindowShowGroupTree = l_GroupTree
      .WindowShowPrintSetupBtn = l_SetupBtn
      .WindowShowSearchBtn = l_SearchBtn
      .WindowTitle = ps_ReportTitle
      .BoundReportHeading = ps_ReportTitle
      .DataFiles(0) = s_RutaTempo
      .SetTablePrivateData 0, 3, o_Data
      .DiscardSavedData = False
      .SQLQuery = s_SqlQuery
      .Connect = Choose(n_Origen% + 1, s_Conexion, "")
      If s_Subreports <> "" Then
        .SubreportToChange = s_Subreports
        .SetTablePrivateData 0, 3, o_Data
        .SubreportToChange = ""
      End If
  End With
  n_Dimension% = UBound(aFormulaStore, 2)
  For i% = 0 To (n_Dimension% - 1)
    ' o_CryReport.StoredProcParam(i%) = aFormulaStore(0, i%)
    o_CryReport.Formulas(i%) = aFormulaStore(1, i%)
    o_CryReport.ParameterFields(i%) = aFormulaStore(2, i%)
  Next i%
  ' Inicializa las Formulas y Parametros Restantes
  InicializaParameters o_CryReport, n_Dimension%, n_Dimension%, n_Dimension%
  ' Filtros Restantes adicionales
  o_CryReport.SelectionFormula = aFilterSelection(0)
  o_CryReport.GroupSelectionFormula = aFilterSelection(1)
  
  ' Ejecuto la Impresión
  o_CryReport.Action = 1
  Exit Sub

ErrorPrinter:
  Select Case Err.Number
   Case 20514      'Archivo ya Existe.
    MsgBox "Ya existe un archivo con este nombre en el Disco, utilice otro para el nombrarlo", vbExclamation + vbOKOnly, ps_ReportTitle
    Exit Sub
   Case 20545      'Cancelar
    MsgBox "La Impresión se Interrumpio i/o Cancela; Verifique", vbExclamation + vbOKOnly, ps_ReportTitle
    Exit Sub
   Case Else
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, ps_ReportTitle
    Exit Sub
  End Select

End Sub
Sub PunteroEnEspera()
  Screen.MousePointer = vbHourglass
End Sub
Sub PunteroNormal()
  Screen.MousePointer = vbDefault
End Sub
Sub RefreshAdoControl(ByVal dca As Object, ByVal tdb As Object, ByVal s_Mensaje As String)
Static rs_Clone As ADODB.Recordset
Static l_ExistRecord As Boolean

' Coloco el puntero en espera
PunteroEnEspera
dca.Refresh
' Verifico si existen Registros
l_ExistRecord = (dca.Recordset.EOF Or dca.Recordset.BOF)
If Not l_ExistRecord Then
    Set rs_Clone = dca.Recordset.Clone()
    rs_Clone.MoveLast: rs_Clone.MoveFirst
    dca.Caption = "Total : " & Trim$(rs_Clone.RecordCount) & " " & s_Mensaje
    rs_Clone.Close
    Set rs_Clone = Nothing
    tdb.Refresh
Else
    dca.Caption = "Total: 0" & " " & s_Mensaje
End If
' Coloco el puntero normal
PunteroNormal

End Sub
Sub SeteaAdoControl(ByVal s_Conexion As String, ByVal o_adc As Object, ByVal tdb As Object, ByVal s_Sentencia As String, ByVal n_adCommandType As Integer, ByVal n_adLockType As Integer, Optional ByVal n_TipoGrilla As Integer)
    
  ' Coloco el Puntero en Espera
  PunteroEnEspera
  ' Captura Columna Atual
  i% = tdb.Col
  With o_adc
    ' Colores de fondo y fuentesdx
    .Font.Name = "Bookman Old Style"
    .Font.Size = 8
    .Font.Bold = True
    .Font.Italic = True
    .BackColor = vbWhite
    .ForeColor = vbBlue
    ' Datos de selección
    .ConnectionString = s_Conexion
    .CommandType = n_adCommandType
    .CursorType = adOpenStatic
    .LockType = n_adLockType
    .RecordSource = s_Sentencia
    .Refresh
    .EOFAction = 0
    .BOFAction = 0
    ' Número de registros seleccionados
    If n_TipoGrilla <> 1 Then
      .Caption = "Total : " & Trim$(.Recordset.RecordCount) & "  " & tdb.Caption
    End If
  End With
  ' Coloco el Puntero Normal
  PunteroNormal
  ' Posiciona en Columna Seleccionada
  tdb.Col = i%

End Sub
Sub ViewGrafics(ByVal o_Form As Object, cmd As Object, ByVal aElemento)
Dim n_Dimension As Integer
' Instancio el objeto de funciones
Set gdl_Funcion = CreateObject("syslink.funciones")
n_Dimension = UBound(aElemento)
' Verifico que exista y Cargo el Icono del Formulario
o_Form.Icon = LoadPicture()
s_Archivo = ps_PathImagen & aElemento(n_Dimension, 1) & ".ico"
If gdl_Funcion.ExisteArchivo(s_Archivo) Then
  o_Form.Icon = LoadPicture(s_Archivo)
End If
' Actualizo el Titulo del Formulario
o_Form.Caption = aElemento(n_Dimension, 2)
' Verifico que exista y Cargo los Graficos, Mensajes de los botones
For i% = 0 To (n_Dimension - 1)
  s_Archivo = aElemento(i%, 1): s_ToolText = aElemento(i%, 2) ':s_Caption = aElemento(i%, 3)
  s_Caption = ""
  LoadGrafics cmd(i%), s_Archivo, s_ToolText, s_Caption
Next i%
Set gdl_Funcion = Nothing

End Sub
' ]
