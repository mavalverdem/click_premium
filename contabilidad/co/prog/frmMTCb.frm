VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMTCb 
   Caption         =   "[Entidad]"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSunat 
      Caption         =   "Sunat"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkVerificar 
      Alignment       =   1  'Right Justify
      Caption         =   "Replicar en Empresas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1485
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpfecha 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   103350273
      CurrentDate     =   37918
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   442
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3480
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   2760
         Picture         =   "frmMTCb.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1950
         Picture         =   "frmMTCb.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdRetroceder 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         Picture         =   "frmMTCb.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   360
      End
      Begin VB.CommandButton cmdAvanzar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         Picture         =   "frmMTCb.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   338
         Width           =   360
      End
      Begin VB.CommandButton cmdCorregir 
         Caption         =   "&Corregir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   480
         Picture         =   "frmMTCb.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1220
         Picture         =   "frmMTCb.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Venta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Compra:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   600
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   6840
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMTCb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean


Private Sub dtpfecha_GotFocus()
'   txtLlave(Index).SelStart = 0
'   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub dtpfecha_LostFocus()
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
'2015-07-27 boton sunar y tc borrar
    ftext_fmt
'    txtDato(0).Text = ""
'    txtDato(1).Text = ""
'2015-07-27 boton sunar y tc borrar
End Sub

Private Sub dtpfecha_Validate(Cancel As Boolean)
   On Error GoTo Err
   
   Dim dvRegistro As Variant

With frmMTCbGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "FehTCb='" & dtpfecha.Value & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
                  If dvRegistro <> -1 Then .Bookmark = dvRegistro
                     Cancel = True
                     Exit Sub
                  End If
               .Bookmark = dvRegistro
            End If
         'End If
   End With
      
      cmdGrabar.Enabled = True
      upHabilitacion True
      pbValidada = True
'   Else
'      cmdGrabar.Enabled = False
'      upHabilitacion False
'      pbValidada = False
'   End If
      
   Exit Sub
Err:
   gpErrores
End Sub
Private Sub cmdSunat_Click()
  Dim dImpTC As Double
  
  dImpTC = fSunatTCambio()
  If dImpTC = 0 Then
    MsgBox (TEXT_9023)
    txtDato(0).Text = "0.00"
    txtDato(1).Text = "0.00"
  End If
End Sub
Public Function fSunatTCambio() As Double
  On Error GoTo ErrorRs
  
  Dim xSeek As Double
  xSeek = 0 'no existe tc en sunat 1=si existe
  
  Dim IE As InternetExplorer
  Dim HTMLdoc As HTMLDocument
  Dim HTMLdoc2 As HTMLDocument '2015-07-21 creado por rcs
  
  Dim TDelements As IHTMLElementCollection
  Dim TDelement As HTMLTableCell
  Dim R As Long
  Dim Tc_Mes, tc_año As String
  Dim intFound As Integer
  
  Dim xDia As String 'rcs
  xDia = Format(Day(dtpfecha.Value), "00") 'rcs
  
  Dim url As String
  
  Dim xFecha As String
  Dim xTC As String
  'Tc_Mes = Application.WorksheetFunction.Match(Range("B5").Value, Range("meses"), 0)
  'tc_año = Range("c5").Value
  
  Tc_Mes = Format(Month(dtpfecha.Value), "00")
  tc_año = Format(Year(dtpfecha.Value), "0000")
  url = "https://e-consulta.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
  Set IE = New InternetExplorer
  
  With IE
    .navigate url
    .Visible = False
    'Esperamos que toda la web cargue
    While .Busy Or .readyState <> READYSTATE_COMPLETE: DoEvents: Wend
    Set HTMLdoc = .document
  End With
  
  With HTMLdoc.selectForm
    .mes.selectedIndex = Tc_Mes
    .anho.Value = tc_año
    .submit
  End With
  Set HTMLdoc2 = IE.document
  While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE: DoEvents: Wend
  Set TDelements = HTMLdoc2.getElementsByTagName("Td")
  
  Dim xRegNum As Integer
  Dim x2Dia As String
  xRegNum = 0
  x2Dia = ""
  For Each TDelement In TDelements
    ' xRegNum = xRegNum + 1
    Select Case TDelement.ClassName
      Case "H3"
        'Range("b65536").End(xlUp).Offset(1, 0).Value = DateSerial(tc_año, Tc_Mes, TDelement.innerText)
        xFecha = Format(DateSerial(tc_año, Tc_Mes, TDelement.innerText), "dd/mm/yyyy")
        'If xdia = Left(xFecha, 2) Then
        x2Dia = Format(TDelement.innerText, "00")
        If xDia = x2Dia Then
          xRegNum = 1
        End If
      Case "tne10"
        If xRegNum > 0 And xDia = x2Dia Then
          If TDelement.innerText = "" Then TDelement.innerText = "SIN T.C"
          xTC = TDelement.innerText
          'Debug.Print xFecha & " - " & xTC
          If xRegNum = 1 Then
            xSeek = 1 'existe tc en sunat
            txtDato(0).Text = xTC
          End If
          If xRegNum = 2 Then
            xSeek = 1 'existe tc en sunat
            txtDato(1).Text = xTC
          End If
        End If
        xRegNum = xRegNum + 1
    End Select
    If xRegNum = 3 Then
      xRegNum = 0
      'Debug.Print xFecha & " - " & xTC
    End If
    'Debug.Print xFecha & " - " & xTC
  Next
  IE.Quit
  
  '2015-07-20 error End With
  fSunatTCambio = xSeek
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Function

'permite buscar los tipo de cambio todo el mes/año
Public Function fSunatTCambio_todo_mes() As Double
  On Error GoTo ErrorRs
  
  Dim IE As InternetExplorer
  Dim HTMLdoc As HTMLDocument
  Dim HTMLdoc2 As HTMLDocument '2015-07-21 creado por rcs
  
  Dim TDelements As IHTMLElementCollection
  Dim TDelement As HTMLTableCell
  Dim R As Long
  Dim Tc_Mes, tc_año As String
  Dim intFound As Integer
  
  Dim xDia As String 'rcs
  xDia = Format(Day(dtpfecha.Value), "00") 'rcs
  
  Dim url As String
  
  Dim xFecha As String
  Dim xTC As String
  'Tc_Mes = Application.WorksheetFunction.Match(Range("B5").Value, Range("meses"), 0)
  'tc_año = Range("c5").Value
  
  Tc_Mes = Format(Month(dtpfecha.Value), "00")
  tc_año = Format(Year(dtpfecha.Value), "0000")
  url = "http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias"
  
  Set IE = New InternetExplorer
  
  With IE
    .navigate url
    .Visible = False
    'Esperamos que toda la web cargue
    While .Busy Or .readyState <> READYSTATE_COMPLETE: DoEvents: Wend
    Set HTMLdoc = .document
  End With
  With HTMLdoc.selectForm
    .mes.selectedIndex = Tc_Mes
    .anho.Value = tc_año
    .submit
  End With
  Set HTMLdoc2 = IE.document
  While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE: DoEvents: Wend
  Set TDelements = HTMLdoc2.getElementsByTagName("Td")
  
  For Each TDelement In TDelements
    Select Case TDelement.ClassName
      Case "H3"
        'Range("b65536").End(xlUp).Offset(1, 0).Value = DateSerial(tc_año, Tc_Mes, TDelement.innerText)
        xFecha = Format(DateSerial(tc_año, Tc_Mes, TDelement.innerText), "dd/mm/yyyy")
      Case "tne10"
        If TDelement.innerText = "" Then TDelement.innerText = "SIN T.C"
        xTC = TDelement.innerText
        
        '            If Range("b65536").End(xlUp).Offset(0, 1).Value = "" Then
        '            Range("c65536").End(xlUp).Offset(1, 0).Value = TDelement.innerText
        '            Else
        '            Range("c65536").End(xlUp).Offset(0, 1).Value = TDelement.innerText
        '            End If
    End Select
    'Debug.Print xFecha & " - " & xTC
    Debug.Print TDelement.ClassName & " - " & TDelement.innerText
  Next
  IE.Quit
  
  '2015-07-20 error End With
  
  '   Set pCnn = New ADODB.Connection
  '   With pCnn
  '        If pTimeout > 0 Then
  '            '2014-09-11 error time out
  '            .CommandTimeout = pTimeout 'segundos de espera
  '        End If
  '        .CursorLocation = adUseClient
  '        .ConnectionString = CONNSTRG & gsNomBDS
  '        .Open
  '    End With
  
  'Set fCnnOpen = pCnn
  fSunatTCambio_todo_mes = 0
  Exit Function

ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  
End Function



Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmMTCbGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      'mskLlave(0).MaxLength = .uorstMain!FehTcb.DefinedSize
      'txtLlave(0).MaxLength = Len(.uorstMain!FehTCb)
    ']
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstMain!ImpTCb_Cpr.DefinedSize
      txtDato(1).MaxLength = .uorstMain!ImpTCb_Vta.DefinedSize
    ']
   End With
   
  ' Visualización de atributos
  chkVerificar.Caption = Choose(gsIdioma, "Replicar en Empresas ", "Replicate in Companies ")
  chkVerificar.Visible = (gsNvlUsr = NvlUsr_Adm And pbNuevo)
   
   If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Fecha:", "Compra:", "Venta:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Date:", "Purchase:", "Sale:")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']
End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
'   If txtDato(0).Text <> "" Then ppAyuDet 0
 ']

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMTCbGrd.uorstMain.BOF And frmMTCbGrd.uorstMain.EOF) Then
   frmMTCbGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMTCbGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMTCbGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   
    '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
 
End Sub

Public Sub cmdGrabar_Click()
  Dim sExpresion As String
   On Error GoTo Err

   With frmMTCbGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
           ' !FyHMdf = Now
         End If
        .Update
        
        ' Inserta registro masivos
        If (pbNuevo And chkVerificar.Value = vbChecked And gsNvlUsr = NvlUsr_Adm) Then
          sExpresion = "INSERT INTO tgtcb (codemp, fehtcb, imptcb_cpr, imptcb_vta, usrcre, fyhcre) "
          sExpresion = sExpresion & "SELECT codemp, '" & Format(dtpfecha.Value, "yyyy-mm-dd") & "' AS fehtcb, "
          sExpresion = sExpresion & CDec(txtDato(0).Text) & " AS imptcb_cpr, " & CDec(txtDato(1).Text) & " AS imptcb_vta, "
          sExpresion = sExpresion & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
          sExpresion = sExpresion & "FROM siscfg.sgpms pms "
          sExpresion = sExpresion & "WHERE pms.codusr='" & gsCodUsr & "' "
          sExpresion = sExpresion & "AND pms.codsis='" & gsCodSis & "' AND pms.codmdl='frmMTCbGrd' "
          sExpresion = sExpresion & "AND pms.codemp<>'" & gsCodEmp & "' "
          sExpresion = sExpresion & "AND EXISTS(SELECT * FROM tgtcb tcb WHERE tcb.codemp=pms.codemp "
          sExpresion = sExpresion & "AND DATE_FORMAT(tcb.fehtcb, '%Y-%m-%d')<>'" & Format(dtpfecha.Value, "yyyy-mm-dd") & "') "
          sExpresion = sExpresion & "ORDER BY codemp"
          frmMTCbGrd.uocnnMain.Execute sExpresion
        End If
      End With
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
  
         .uorstMain.Find "FehTCb='" & dtpfecha.Value & "'"

       ']
         cmdGrabar.Enabled = False
         upHabilitacion False
   
         upDatosPredeterminados
       '[Llave con el foco al añadir.  'Cambiar.
         'txtLlave(0).SetFocus
         dtpfecha.SetFocus
         
       ']
      Else
         cmdRetroceder.Enabled = True
         cmdAvanzar.Enabled = True
         cmdCorregir.Enabled = True
         cmdGrabar.Enabled = False
         cmdDeshacer.Enabled = False
         upHabilitacion False
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
  
   frmMTCbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub
Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub
Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.

'If Chr(KeyAscii) > 0 And Chr(KeyAscii) <= 9 Then
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
'Else
'
'   KeyAscii = 0
'
'End If

']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
'   On Error GoTo Err
'
'  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 12                             'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select
'
'  'Asigna 0 a campos numéricos si están vacíos.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select
'
'  'Busca el dato en su tabla principal.
'   Select Case Index
'   Case 12                             'Cambiar (añadir índices).
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
'  Exit Sub
'Err:
'   gpErrores
End Sub
Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  
  On Error GoTo Err
  
  With frmMTCbGrd
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        .uorstMain!codemp = gsCodEmp
        .uorstMain!FehTCb = dtpfecha.Value
        '.uorstMain!FehTCb = mskLlave(0).Text
      End If
    
      'Datos.
      .uorstMain!ImpTCb_Cpr = txtDato(0).Text
      .uorstMain!ImpTCb_Vta = txtDato(1).Text
    Else
      'Llaves.
      dtpfecha.Value = Format(.uorstMain!FehTCb, "dd/mm/yyyy")
      
      'Datos.
      txtDato(0).Text = Format(IIf(IsNull(.uorstMain!ImpTCb_Cpr), "", .uorstMain!ImpTCb_Cpr), "#0.0000")
      txtDato(1).Text = Format(IIf(IsNull(.uorstMain!ImpTCb_Vta), "", .uorstMain!ImpTCb_Vta), "#0.0000")
    End If
  End With
  
  Exit Sub
Err:
  gpErrores
  
  Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
  Dim dnContador As Integer
  
  'Llaves.
  dtpfecha.Value = Date
  'Datos.
  chkVerificar.Value = vbUnchecked
  ftext_fmt

End Sub

'2015-07-27 boton sunar y tc borrar con 0.00
Public Sub ftext_fmt()
  Dim dnContador As Integer
  
  For dnContador = 0 To txtDato.Count - 1
    txtDato(dnContador).Text = Format(0, FORMATO_NUM_2)
  Next

End Sub
'2015-07-27 boton sunar y tc borrar con 0.00

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Integer
  
  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  '2015-07-27 boton sunar y tc borrar
  cmdSunat.Enabled = pbNuevo
  '2015-07-27 boton sunar y tc borrar
  
End Sub
'[Código propio del formulario.
']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
   
   'Orden: Corregir.
   zaOpciones = Array(gbPms02)
End Property
Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdCorregir.Enabled = IIf(pbNuevo, False, taOpciones(0))
End Property


