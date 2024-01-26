VERSION 5.00
Begin VB.Form frmRBceCpbCCo 
   Caption         =   "[título]"
   ClientHeight    =   5565
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6990
   Begin VB.Frame fraNivelCenCos 
      Caption         =   " Nivel de Centro de Costo "
      ForeColor       =   &H80000002&
      Height          =   840
      Left            =   0
      TabIndex        =   38
      Top             =   4050
      Width           =   4335
      Begin VB.OptionButton optNivCCo 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   3
         Left            =   3000
         TabIndex        =   43
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   41
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   2
         Left            =   2040
         TabIndex        =   40
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5520
      TabIndex        =   37
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   34
      Top             =   4050
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   36
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   35
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame fraNivelCuenta 
      Caption         =   "Nivel de Cuentas"
      ForeColor       =   &H80000002&
      Height          =   840
      Left            =   0
      TabIndex        =   28
      Top             =   3150
      Width           =   6975
      Begin VB.OptionButton optNivCta 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   7
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "8 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   6
         Left            =   5880
         TabIndex        =   17
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "7 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   5
         Left            =   4920
         TabIndex        =   16
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "6 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   4
         Left            =   3960
         TabIndex        =   15
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   3
         Left            =   3000
         TabIndex        =   14
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Frame fraAlcance 
      Caption         =   "Alcance"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   0
      TabIndex        =   27
      Top             =   2430
      Width           =   1980
      Begin VB.OptionButton optAlcance 
         Caption         =   "del mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   990
         TabIndex        =   9
         Top             =   255
         Width           =   870
      End
      Begin VB.OptionButton optAlcance 
         Caption         =   "al mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   780
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rangos"
      ForeColor       =   &H00800000&
      Height          =   2265
      Left            =   0
      TabIndex        =   19
      Top             =   90
      Width           =   6990
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRBceCpbCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1860
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRBceCpbCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1500
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
         TabIndex        =   7
         Top             =   1845
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
         TabIndex        =   6
         Top             =   1485
         Width           =   945
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   4500
         Picture         =   "frmRBceCpbCCo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   855
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
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   630
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
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   4500
         Picture         =   "frmRBceCpbCCo.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   495
         Width           =   255
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
         TabIndex        =   30
         Top             =   1845
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
         TabIndex        =   29
         Top             =   1500
         Width           =   5520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   1260
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centros de Costo"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   1215
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
         Index           =   3
         Left            =   780
         TabIndex        =   23
         Top             =   855
         Width           =   3720
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
         Left            =   780
         TabIndex        =   22
         Top             =   495
         Width           =   3720
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
      ScaleWidth      =   6990
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5025
      Width           =   6990
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
         Picture         =   "frmRBceCpbCCo.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1245
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
         Picture         =   "frmRBceCpbCCo.frx":0BDA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
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
         Picture         =   "frmRBceCpbCCo.frx":0D24
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   5070
      TabIndex        =   25
      Top             =   2565
      Width           =   600
   End
End
Attribute VB_Name = "frmRBceCpbCCo"
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
Private porstCOCta As ADODB.Recordset
Private porstCoCCo As ADODB.Recordset
Private pnNivCta As Byte
']
Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstCoCCo = New ADODB.Recordset
   
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
      .Source = "SELECT CodCta, DetCta " _
              & "FROM CoCta"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstCoCCo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCCo, DetCCo " _
              & "FROM CoCCo"
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
      For dnContador = 2 To 3
         .Item(dnContador).DataField = "CodCCo"
         .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
    With cboTpoMon
        .AddItem TPOMON_NAC_TXT_1, 0
        .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND
    
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

    For dnContador = 1 To Len(gsNivCCo)
        optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Visible = True
        Select Case dnContador
         Case Is = 1
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 120
         Case Is = 2
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 1080
         Case Is = 3
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 2040
         Case Is = 4
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 3000
        End Select
    Next
    optNivCCo(4).Value = True
    fraNivelCenCos.Width = optNivCCo(Val(Mid(gsNivCCo, dnContador - 1, 1)) - 2).Left + 1300

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !CodCta
      .MoveFirst
      txtDato(0).Text = !CodCta
   End With
   With porstCoCCo
      .MoveLast
      txtDato(3).Text = !CodCCo
      .MoveFirst
      txtDato(2).Text = !CodCCo
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
   If txtDato(3).Text <> "" Then ppAyuDet 3
   
  'Otros.
   
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
   porstCOCta.Close
   porstCoCCo.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstCoCCo = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2, 3
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    Dim dnContador   As Byte
    Dim CadCrystal   As String, s_Moneda As String
    Dim n_NivCoCCo As Integer, n_MesIni As Integer
    Dim s_SaldoDeb As String, s_SaldoHab As String
   
    ppHabilitacion False
    n_NivCoCCo = IIf(optNivCCo(0).Value, 2, IIf(optNivCCo(1).Value, 3, 5))
    n_MesIni = Val(IIf(optAlcance(0).Value, 0, gsMesAct))
            
    s_SaldoDeb = "ROUND(SUM(": s_SaldoHab = "ROUND(SUM("
    s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0)
    For dnContador = n_MesIni To Val(gsMesAct)
      s_SaldoDeb = s_SaldoDeb & "a.AcuD" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = Val(gsMesAct), "", "+")
      s_SaldoHab = s_SaldoHab & "a.AcuH" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = Val(gsMesAct), "", "+")
    Next dnContador
    s_SaldoDeb = s_SaldoDeb & "), 2) AS cSumaD, "
    s_SaldoHab = s_SaldoHab & "), 2) AS cSumaH"
    
    With porstMRp
        If .State = adStateOpen Then .Close
        .Source = "SELECT LEFT(a.CodCCo, " & n_NivCoCCo & ") AS CodCCo, b.DetCCo, a.CodCta, c.DetCta, c.TpoSdo, c.TpoCta, "
        .Source = .Source & s_SaldoDeb & s_SaldoHab
        .Source = .Source & " FROM ((COCCoAcu a" _
                          & " LEFT JOIN CoCCo b ON LEFT(a.CodCCo, " & n_NivCoCCo & ")=b.CodCCo) " _
                          & " LEFT JOIN CoCta c ON a.CodCta=c.CodCta) " _
                          & " WHERE a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' " _
                          & " AND b.CodCCo BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
        If pnNivCta = 2 Then
          .Source = .Source & " AND Length(trim(a.CodCta))=" & pnNivCta & "  "
        Else
          If pnNivCta = 9 Then
              .Source = .Source & " AND c.TpoCta = '1' "
          Else
              .Source = .Source & " AND (Length(trim(a.CodCta))=" & pnNivCta & "  " _
                                & " OR  (Length(trim(a.CodCta))<" & pnNivCta & " AND c.TpoCta= " & TPOCTA_TRA & " )) "
          End If
        End If
        .Source = .Source & " GROUP BY LEFT(a.CodCCo, " & n_NivCoCCo & "), a.CodCta"
        .Source = .Source & " ORDER BY a.CodCCo, c.CodCta"
        .Open
    End With
     
    usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
    If usDEstino = PRN_DEST_GRAF Then
        pocnnMain.Execute "DROP TABLE IF EXISTS trptbcecpbcco"
        cmdImprimir(Index).Tag = "CREATE TABLE trptbcecpbcco " & porstMRp.Source
        pocnnMain.Execute cmdImprimir(Index).Tag
        Call gpEncabezadoRpt(frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
        With frmMain.rptMain
            .ReportFileName = gsRutRpt & "rptRBceCpbCCo.rpt"
            'Fórmular propias.
            .Formulas(5) = "mPeriodo='" & IIf(optAlcance(0).Value, "Acumulado - ", "") & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
            '.Formulas(6) = "fSaldoInicial=" & gsIniMesCnt
            ']
            .WindowShowExportBtn = IIf(paOpciones(2), True, False)
            .MarginLeft = unMargenIzquierdo
            .WindowState = crptMaximized
            .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
            .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
            .Action = 1
        End With
        pocnnMain.Execute "DROP TABLE IF EXISTS trptbcecpbcco"
   Else
        Set MRViewer = New MRViewerObject
        With MRViewer
           .DataRecordSet = porstMRp
           .LoadReport gsRutRpt & "rptRBceCpbCCo.mrp"
           Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
          '[Parámetros adicionales.
          .Parameters("pPeriodoAdc") = IIf(optAlcance(0).Value, "A ", "") & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
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
   Case 0, 1, 2, 3                   'Cambiar (añadir índices).
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
   Case 2, 3                           'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
             lblDatoDeta(tnIndex).Caption = " " & !DetCta
          End If
       End With
    Case 2, 3
       If txtDato(tnIndex).Text = "" Then
          lblDatoDeta(tnIndex).Caption = ""
          Exit Function
       End If
       With porstCoCCo
          .MoveFirst
          .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
             MsgBox TEXT_8006, vbExclamation
             ppAyuDet = True
          Else
             lblDatoDeta(tnIndex).Caption = " " & !DetCCo
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



