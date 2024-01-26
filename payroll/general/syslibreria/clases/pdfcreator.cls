VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreadorPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private WithEvents po_PDFCreator As PDFCreator.clsPDFCreator
Attribute po_PDFCreator.VB_VarHelpID = -1
Private po_PDFCreatorOption As clsPDFCreatorOptions

Private ps_DefaultPrinter As String
Private pb_Inicializa As Boolean

Public Event eReady(sOutputFilename As String)
Public Event eError(lError As Long, sDescription As String)
Private Sub Class_Terminate()
  o_cClose
End Sub
' Retorno   : <True> = Se cerró sin error
Function o_cClose() As Boolean
  ' Verifica si Incializado
  If (pb_Inicializa And Not (po_PDFCreator Is Nothing)) Then
    po_PDFCreator.cDefaultPrinter = ps_DefaultPrinter
    po_PDFCreator.cClearCache
    o_cClose = po_PDFCreator.cClose
    While po_PDFCreator.cProgramIsRunning
      DoEvents
      Sleep 3000
    Wend
  End If
  Set po_PDFCreator = Nothing
  Set po_PDFCreatorOption = Nothing
End Function
' Finalidad : Configurar comportamiento y salida.
' Retorno   : <True> si se puede imprimir
' Parámetros: <s_RutaArchivo> = Directorio donde se ubicará el PDF
'             <s_Archivo> = Nombre que se le asignará al PDF
'             <b_Formato> Tipo 0 = PDF, 1 = PNG, 2 = JPEG, 3 = BMP, 4 = PCX, 5 = TIFF, 6 = PS, 7 = EPS, 8 = TXT, 9 = PDF/A-1b, 10 = PDF/X, 11 = PSD, 12 = PCL, 13 = RAW
'             <l_AbrirArchivo> = Indica si se abrirá el PDF al finalizar
Function o_cConfigPrint(s_RutaArchivo As String, s_Archivo As String, Optional ByVal b_Formato As Byte = 0, Optional l_AbrirArchivo As Long = 0) As Boolean
  Dim s_Extension As String
  
  s_Extension = Choose(b_Formato + 1, "pdf", "png", "jpeg", "bmp", "pcx", "tiff", "ps", "eps", "txt", "pdf", "pdf", "psd", "pcl", "raw")
  s_Archivo = s_Archivo & "." & s_Extension
  ' Verifico existencia de archivo
  If Not Dir$(s_RutaArchivo & s_Archivo) = vbNullString Then
    If MsgBox("El archivo '" & s_RutaArchivo & s_Archivo & "' ya existe ¿Sobreescribir?", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
      Kill s_RutaArchivo & s_Archivo
    Else
      Exit Function
    End If
  End If
 
 With po_PDFCreatorOption
  .AutosaveDirectory = s_RutaArchivo
  .AutosaveFilename = s_Archivo
  .UseAutosave = 1
  .UseAutosaveDirectory = 1
  .AutosaveFormat = b_Formato
  .AutosaveStartStandardProgram = l_AbrirArchivo
 End With
 Set po_PDFCreator.cOptions = po_PDFCreatorOption
 po_PDFCreator.cClearCache
 po_PDFCreator.cPrinterStop = False
 Sleep 2000
 
End Function
' Finalidad : Arranque de PDFCreator
' Retorno   : <True> = No existía una instancia activa
'             <False> = Ya existía una instancia activa. Cancelar impresión.
' Parámetros: Para este uso, en principio solo se utiliza el parámetro por defecto
Function o_cStart(Optional s_Parametro As String = "/NoProcessingAtStartup") As Boolean
  
  pb_Inicializa = o_cStart
  Set po_PDFCreator = New clsPDFCreator
  With po_PDFCreator
    .cVisible = True
    o_cStart = .cStart(s_Parametro)
    If o_cStart = False Then
      If .cStart(s_Parametro, True) = False Then
        Exit Function
      End If
      .cVisible = True
    End If
    Set po_PDFCreatorOption = .cOptions
    ps_DefaultPrinter = po_PDFCreator.cDefaultPrinter
    po_PDFCreator.cDefaultPrinter = "PDFCreator"
    .cClearCache
  End With
  pb_Inicializa = o_cStart
  
End Function
' Retorno   : El estado de actividad de PDFCreator
Public Property Get o_cPrinterStop() As Boolean
  o_cPrinterStop = po_PDFCreator.cPrinterStop
End Property
' Finalidad : Establece estado de actividad de PDFCreator
Public Property Let o_cPrinterStop(ByVal bcPrinterStop As Boolean)
  po_PDFCreator.cPrinterStop = bcPrinterStop
End Property
Private Sub po_PDFCreator_eError()
  ' Muestro el mensaje de error
  MsgBox "Inicializa, Verifique . . ." & vbCrLf & vbCrLf & _
  "Numero: " & po_PDFCreator.cError.Number & vbCrLf & _
  "Descripcion: " & po_PDFCreator.cError.Description & vbCrLf, vbCritical, "ERROR"
  RaiseEvent eError(po_PDFCreator.cError.Number, po_PDFCreator.cError.Description)
End Sub
Private Sub po_PDFCreator_eReady()
  po_PDFCreator.cPrinterStop = True
  RaiseEvent eReady(po_PDFCreator.cOutputFilename)
End Sub
