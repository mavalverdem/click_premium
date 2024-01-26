VERSION 5.00
Begin VB.Form frmRMayGrl 
   Caption         =   "[título]"
   ClientHeight    =   4785
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDivisoria 
      Caption         =   "Divisionarias"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   15
      TabIndex        =   21
      Top             =   2400
      Width           =   1980
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   31
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   28
      Top             =   3480
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   75
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
         Width           =   1020
      End
   End
   Begin VB.Frame fraNivelCuenta 
      Caption         =   "Nivel de Cuentas"
      ForeColor       =   &H80000002&
      Height          =   1050
      Left            =   0
      TabIndex        =   18
      Top             =   1350
      Width           =   6975
      Begin VB.OptionButton optNivCta 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   660
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   26
         Top             =   675
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   2
         Left            =   2025
         TabIndex        =   25
         Top             =   675
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   3
         Left            =   3015
         TabIndex        =   24
         Top             =   675
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "6 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   23
         Top             =   675
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "7 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   5
         Left            =   4950
         TabIndex        =   22
         Top             =   675
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "8 dígitos"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   6
         Left            =   5895
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Value           =   -1  'True
         Width           =   915
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
      ScaleWidth      =   6975
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4245
      Width           =   6975
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
         Picture         =   "frmRMayGrl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmRMayGrl.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmRMayGrl.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5715
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2595
      Width           =   1260
   End
   Begin VB.Frame fraAlcance 
      Caption         =   "Alcance"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   4950
      TabIndex        =   15
      Top             =   4410
      Visible         =   0   'False
      Width           =   1980
      Begin VB.OptionButton optAlcance 
         Caption         =   "al mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   780
      End
      Begin VB.OptionButton optAlcance 
         Caption         =   "del mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   990
         TabIndex        =   7
         Top             =   255
         Width           =   870
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   9
      Top             =   45
      Width           =   6975
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6570
         Picture         =   "frmRMayGrl.frx":077E
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   855
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6570
         Picture         =   "frmRMayGrl.frx":0928
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   495
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
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
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   585
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
         Left            =   1050
         TabIndex        =   13
         Top             =   495
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
         Index           =   1
         Left            =   1050
         TabIndex        =   12
         Top             =   840
         Width           =   5520
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   4965
      TabIndex        =   16
      Top             =   2640
      Width           =   675
   End
End
Attribute VB_Name = "frmRMayGrl"
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
Private pnNivCta As Byte
'[Propio del formulario.
Private porstCOCta As ADODB.Recordset
']


Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   
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
      .Source = .Source & "FROM COCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY CodCta"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
   End With

   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
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
  fraNivelCuenta.Caption = Choose(gsIdioma, "Nivel de Cuentas", "Account Level")
  optNivCta(7).Caption = Choose(gsIdioma, "Detalle", "Detail")
  optNivCta(0).Caption = Choose(gsIdioma, "2 dígitos", "2 digits")
  optNivCta(1).Caption = Choose(gsIdioma, "3 dígitos", "3 digits")
  optNivCta(2).Caption = Choose(gsIdioma, "4 dígitos", "4 digits")
  optNivCta(3).Caption = Choose(gsIdioma, "5 dígitos", "5 digits")
  optNivCta(4).Caption = Choose(gsIdioma, "6 dígitos", "6 digits")
  optNivCta(5).Caption = Choose(gsIdioma, "7 dígitos", "7 digits")
  optNivCta(6).Caption = Choose(gsIdioma, "8 dígitos", "8 digits")
  chkDivisoria.Caption = Choose(gsIdioma, "Divisionarias", "Subsidiary Accounts")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
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
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   optAlcance(0).Value = True
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
   porstCOCta.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sDebAnt As String, sHabAnt As String
  Dim sDebAct As String, sHabAct As String
  
  ppHabilitacion False
    
  If pnNivCta = 9 Then pnNivCta = Val(Right(gsNivCta, 1))
  With porstMRp
    If .State = adStateOpen Then .Close
    If gsMesAct <> "00" Then
      sDebAnt = gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2))
      sHabAnt = gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4))
      sDebAct = gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2))
      sHabAct = gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4))
      .Source = "SELECT COCtaAcu.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
      .Source = .Source & "(" & sDebAnt & ") AS cDebeAnt, "
      .Source = .Source & "(" & sHabAnt & ") AS cHaberAnt, "
      .Source = .Source & "(" & sDebAct & ") AS cDebeMes, "
      .Source = .Source & "(" & sHabAct & ") AS cHaberMes "
      .Source = .Source & "FROM (COCtaAcu "
      .Source = .Source & "LEFT JOIN CoCta b ON COCtaAcu.codemp=b.codemp AND COCtaAcu.pdoano=b.pdoano AND COCtaAcu.CodCta=b.CodCta) "
      .Source = .Source & "WHERE COCtaAcu.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND COCtaAcu.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND COCtaAcu.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    Else
      sDebAnt = "0": sHabAnt = "0"
      sDebAct = gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2))
      sHabAct = gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4))
      .Source = "SELECT COCtaAcu.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
      .Source = .Source & sDebAnt & " AS cDebeAnt, "
      .Source = .Source & sHabAnt & " AS cHaberAnt, "
      .Source = .Source & "(" & sDebAct & ") AS cDebeMes, "
      .Source = .Source & "(" & sHabAct & ") AS cHaberMes "
      .Source = .Source & "FROM (COCtaAcu "
      .Source = .Source & "LEFT JOIN CoCta b ON COCtaAcu.codemp=b.codemp AND COCtaAcu.pdoano=b.pdoano AND COCtaAcu.CodCta=b.CodCta) "
      .Source = .Source & "WHERE COCtaAcu.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND COCtaAcu.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND COCtaAcu.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    End If
    If pnNivCta = 2 Then
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(COCtaAcu.CodCta))=" & pnNivCta & " "
    Else
      If chkDivisoria.Value = 1 Then
        .Source = .Source & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(COCtaAcu.CodCta))=" & pnNivCta & " OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(COCtaAcu.CodCta))=2)) "
      Else
        .Source = .Source & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(COCtaAcu.CodCta))=" & pnNivCta & " "
        .Source = .Source & "OR  (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(COCtaAcu.CodCta))<" & pnNivCta & " AND b.TpoCta= " & TPOCTA_TRA & " )) "
      End If
    End If
    If ps_Plataforma = pSrvMySql Then
      .Source = .Source & "HAVING (cDebeAnt + cHaberAnt + cDebeMes + cHaberMes) > 0 "
    Else
      .Source = .Source & "AND (" & sDebAnt & " + " & sHabAnt & " + " & sDebAct & " + " & sHabAct & ") > 0 "
    End If
    .Source = .Source & "ORDER BY COCtaAcu.CodCta"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptRMayGrl.rpt"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .WindowState = crptMaximized
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    '& " AND Length(COCtaAcu.CodCta)=" & pnNivCta & "  "
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRMayGrl.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      '         .Parameters("pTipoFecha") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
  'Valida que para la Opcion de 2dig este Desabilitada
  If optNivCta.Item(0).Value Then
    chkDivisoria.Value = False
    chkDivisoria.Enabled = False
  Else
    chkDivisoria.Enabled = True
  End If
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
'   Select Case Index    'Completa con ceros a la izquierda.
'   Case 0, 1                           'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
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
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

