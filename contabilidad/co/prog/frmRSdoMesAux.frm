VERSION 5.00
Begin VB.Form frmRSdoMesAux 
   Caption         =   "[título]"
   ClientHeight    =   5775
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   31
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4680
      TabIndex        =   28
      Top             =   4440
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   30
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   29
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraNivelCuenta 
      Caption         =   "Nivel de Cuentas"
      ForeColor       =   &H80000002&
      Height          =   1050
      Left            =   0
      TabIndex        =   26
      Top             =   3240
      Width           =   6915
      Begin VB.OptionButton optNivCta 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "6 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   12
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "7 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   13
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "8 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   6
         Left            =   5880
         TabIndex        =   14
         Top             =   660
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Auxiliar"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   23
      Top             =   1530
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   6885
         Picture         =   "frmRSdoMesAux.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   325
         Width           =   255
      End
      Begin VB.TextBox txtDato 
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
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   315
         Width           =   1260
      End
      Begin VB.Label lblDatoDeta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1365
         TabIndex        =   25
         Top             =   315
         Width           =   5520
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   6075
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2385
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1365
      Left            =   0
      TabIndex        =   16
      Top             =   90
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRSdoMesAux.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   870
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRSdoMesAux.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   510
         Width           =   255
      End
      Begin VB.TextBox txtDato 
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
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   945
      End
      Begin VB.TextBox txtDato 
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
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   945
      End
      Begin VB.Label lblDatoDeta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1065
         TabIndex        =   20
         Top             =   855
         Width           =   5520
      End
      Begin VB.Label lblDatoDeta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1065
         TabIndex        =   19
         Top             =   510
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7290
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5235
      Width           =   7290
      Begin VB.CommandButton cmdConfig 
         Caption         =   "&Configuración de Impresora"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2355
         TabIndex        =   2
         Top             =   0
         Width           =   1125
      End
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
         Height          =   495
         Left            =   3720
         Picture         =   "frmRSdoMesAux.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Vista Preliminar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   0
         Picture         =   "frmRSdoMesAux.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1245
         Picture         =   "frmRSdoMesAux.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   5265
      TabIndex        =   18
      Top             =   2430
      Width           =   735
   End
End
Attribute VB_Name = "frmRSdoMesAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MRViewer As MRViewerObject
Attribute MRViewer.VB_VarHelpID = -1

Public udFecha As Date
Public unCopias As Integer
Public unMargenIzquierdo As Integer
Public usDEstino As String
Public usOrientacionRpt As String
Public usOrientacionOri As String
Private paOpciones As Variant
Private pocnnMain As ADODB.Connection
Private porstMRp As ADODB.Recordset

'[Propio del formulario.
Private pnNivCta    As Byte
Private porstCOCta As ADODB.Recordset
Private porstTGAux As ADODB.Recordset
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstTGAux = New ADODB.Recordset
   
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstMRp
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
   With porstCOCta
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
      .Source = .Source & "FROM CoCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TGAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With

 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      .Item(2).DataField = "CodAux"
      .Item(2).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
   End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Currency :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraAuxiliar.Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraNivelCuenta.Caption = Choose(gsIdioma, "Nivel de Cuentas", "Account Level")
  optNivCta(7).Caption = Choose(gsIdioma, "Detalle", "Detail")
  optNivCta(0).Caption = Choose(gsIdioma, "2 dígitos", "2 digits")
  optNivCta(1).Caption = Choose(gsIdioma, "3 dígitos", "3 digits")
  optNivCta(2).Caption = Choose(gsIdioma, "4 dígitos", "4 digits")
  optNivCta(3).Caption = Choose(gsIdioma, "5 dígitos", "5 digits")
  optNivCta(4).Caption = Choose(gsIdioma, "6 dígitos", "6 digits")
  optNivCta(5).Caption = Choose(gsIdioma, "7 dígitos", "7 digits")
  optNivCta(6).Caption = Choose(gsIdioma, "8 dígitos", "8 digits")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !codcta
      .MoveFirst
      txtDato(0).Text = !codcta
   End With
   
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
  
  'Otros.
    For dnContador = 1 To Len(gsNivCta)
        optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Visible = True
        Select Case dnContador
         Case Is = 1
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 120
         Case Is = 2
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 1080
         Case Is = 3
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 2040
         Case Is = 4
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 3000
         Case Is = 5
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 3960
         Case Is = 6
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 4920
         Case Is = 7
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 5880
        End Select
    Next
'    optNivCta(Val(Mid(gsNivCta, dncontador - 1, 1)) - 2).Value = True
    optNivCta(7).Value = True
    pnNivCta = 9
    fraNivelCuenta.Width = optNivCta(Val(Mid(gsNivCta, dnContador - 1, 1)) - 2).Left + 1035
    
  'Características de impresión.
   chkImpFecha.Value = vbChecked
   udFecha = Date                      'Fecha en el encabezado.
   unCopias = 1 'frmMain.rptMain.CopiesToPrinter  'Cantidad de Copias.
   unMargenIzquierdo = 240             'Margen izquierdo.
   usDEstino = PRN_DEST_MATR           'PRN_DEST_GRAF:ica _
                                        PRN_DEST_MATR:icial.
   usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                        PRN_ORIE_HORI:zontal.
 ']
   frmOPrnCfg.OrientacionPrn 0, Me
   frmOPrnCfg.lblOriPrn.Caption = Printer.Orientation
      
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub Form_Activate()
   'Orden: Vista Previa, Imprimir, Exportar.
   zaOpciones = Array(gbPms04, gbPms05, gbPms06)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Cambiar. Añadir recordsets.
   porstTGAux.Close
   porstCOCta.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim cCadReporte As String, sMoneda As String
  Dim sExpresion As String, nContador As Long
  
  ppHabilitacion False
      
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  
  cCadReporte = "SELECT a.CodCta, a.CodAux, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, c.RazAux, "
  For nContador = 0 To 12
    sExpresion = "ROUND(a.AcuD" & Format(nContador, "00") & "_" & sMoneda & "-a.AcuH" & Format(nContador, "00") & "_" & sMoneda & ", 2)"
    cCadReporte = cCadReporte & IIf(gsMesAct >= Format(nContador, "00"), sExpresion, "0.00") & " AS clmAcu" & Format(nContador, "00") & IIf(nContador <> 12, ", ", " ")
  Next nContador
  cCadReporte = cCadReporte & "FROM ((CoAuxAcu a "
  cCadReporte = cCadReporte & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
  cCadReporte = cCadReporte & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND a.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
  If pnNivCta = 2 Then
    cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
  Else
    If pnNivCta = 9 Then
      cCadReporte = cCadReporte & "AND b.TpoCta='" & TPOCTA_TRA & "' "
    Else
      cCadReporte = cCadReporte & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(b.CodCta))=" & pnNivCta & " "
      cCadReporte = cCadReporte & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(b.CodCta))<" & pnNivCta & " "
      cCadReporte = cCadReporte & "AND b.TpoCta='" & TPOCTA_TRA & "')) "
    End If
  End If
  If Trim(txtDato(2).Text) <> "" Then
    cCadReporte = cCadReporte & "AND a.CodAux='" & Trim(txtDato(2).Text) & "' "
  End If
  If ps_Plataforma = pSrvMySql Then
    cCadReporte = cCadReporte & "HAVING ("
    For nContador = 0 To CLng(gsMesAct)
      cCadReporte = cCadReporte & "clmAcu" & Format(nContador, "00") & " <> 0.00" & IIf(nContador <> CLng(gsMesAct), " OR ", ") ")
    Next nContador
  ElseIf ps_Plataforma = pSrvSql Then
    cCadReporte = cCadReporte & "AND ("
    For nContador = 0 To CLng(gsMesAct)
      sExpresion = "ROUND(a.AcuD" & Format(nContador, "00") & "_" & sMoneda & "-a.AcuH" & Format(nContador, "00") & "_" & sMoneda & ", 2)<>0.00"
      cCadReporte = cCadReporte & sExpresion & IIf(nContador <> CLng(gsMesAct), " OR ", ") ")
    Next nContador
  End If
  cCadReporte = cCadReporte & "ORDER BY a.CodCta, a.CodAux"
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = cCadReporte
    .Open
  End With
   
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '       '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRSdoMesAux.rpt"
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRSdoMesAux.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = gsAnoAct
      ']
      
      If Index = 0 Then
        .PreviewReport
      Else
        '[ARREGLAR: Revisar el uso de los tres primeros parámetros de Print.
        .Print 1, 0, 0, unCopias
        ']ARREGLAR.
      End If
      .UnLoadReport
    End With
    
    Set MRViewer = Nothing
  End If
    
   ppHabilitacion True
End Sub

Private Sub cmdConfig_Click()
   With frmOPrnCfg
      .ConfiguraPrn 0, Me
   
      .Show vbModal
    
      .ConfiguraPrn 1, Me
   End With
   
   cmdImprimir(1).SetFocus
End Sub

Private Sub cmdSalir_Click()
   frmOPrnCfg.OrientacionPrn 1, Me
   
   Unload Me
End Sub

Private Sub optNivCta_Click(Index As Integer)

    pnNivCta = Index + 2
    
End Sub

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   'Select Case Index    'Completa con ceros a la izquierda.
   'Case 0, 1                           'Cambiar (añadir índices).
   '   If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
   '      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
   '   End If
   'End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2                          'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
     Case 0, 1                           'Cambiar (añadir índices).
       modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
       txtDato(tnIndex).Text = frmOAyuBus.uvDato1
       lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 2                              'Cambiar (añadir índices).
       modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
       txtDato(tnIndex).Text = frmOAyuBus.uvDato1
       lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0, 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstCOCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
    Case 2
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstTGAux
         .MoveFirst
         .Find "CodAux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !razaux
         End If
      End With
   End Select
End Function

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

  'Controles del formulario.
'   cboTpoMon.Enabled = tbHabilitar
'   dtpFecha.Enabled = tbHabilitar
'   optTipo(0).Enabled = tbHabilitar
'   optTipo(1).Enabled = tbHabilitar
'   With txtDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With cmdDatoAyud
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With lblDatoDeta
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

