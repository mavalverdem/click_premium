VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPCieApe 
   Caption         =   "[título]"
   ClientHeight    =   5715
   ClientLeft      =   1905
   ClientTop       =   1245
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7425
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   400
      Left            =   3600
      TabIndex        =   29
      Top             =   5250
      Width           =   1215
   End
   Begin VB.Frame frmCuadro 
      Caption         =   " Flujo de Caja "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1485
      Index           =   2
      Left            =   2760
      TabIndex        =   10
      Top             =   3090
      Width           =   4575
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   4215
         Picture         =   "frmpcieape.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1095
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   4
         Left            =   120
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1095
         Width           =   465
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   4215
         Picture         =   "frmpcieape.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   510
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   3
         Left            =   120
         MaxLength       =   4
         TabIndex        =   11
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lblTexto 
         Caption         =   "Flujo de Caja Acreedor :"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   2500
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
         Height          =   300
         Index           =   2
         Left            =   570
         TabIndex        =   27
         Top             =   1095
         Width           =   3660
      End
      Begin VB.Label lblTexto 
         Caption         =   "Flujo de Caja Deudor :"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   255
         Width           =   2500
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
         Height          =   300
         Index           =   1
         Left            =   570
         TabIndex        =   24
         Top             =   510
         Width           =   3660
      End
   End
   Begin VB.Frame frmCuadro 
      Caption         =   " Parametros "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   2
         Left            =   135
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2505
         Width           =   4365
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   1
         Left            =   135
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2175
         Width           =   4365
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Periodo"
         Top             =   450
         Width           =   2000
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   0
         Left            =   120
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1035
         Width           =   465
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   4215
         Picture         =   "frmpcieape.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1035
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtpFehCpb 
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1605
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   216727553
         CurrentDate     =   37953
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
         Height          =   300
         Index           =   0
         Left            =   570
         TabIndex        =   22
         Top             =   1035
         Width           =   3660
      End
      Begin VB.Label lblTexto 
         Caption         =   "Periodo Comprobante :"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   210
         Width           =   2505
      End
      Begin VB.Label lblTexto 
         Caption         =   "Diario de Comprobante :"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   795
         Width           =   2505
      End
      Begin VB.Label lblTexto 
         Caption         =   "Fecha de Comprobante :"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1365
         Width           =   2505
      End
      Begin VB.Label lblTexto 
         Caption         =   "Glosa de Comprobante :"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1950
         Width           =   2505
      End
   End
   Begin VB.Frame frmCuadro 
      Caption         =   " Tipo de Proceso "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1000
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optProceso 
         Caption         =   "&Cierre de Año 2003"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2000
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Apertura de Año 2004"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   460
         Width           =   2000
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Inicializa Tablas Maestras"
         ForeColor       =   &H00C00000&
         Height          =   200
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2150
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   400
      Left            =   2340
      TabIndex        =   16
      Top             =   5250
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   400
      Left            =   6120
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4875
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Información ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   4635
      Width           =   2445
   End
End
Attribute VB_Name = "frmPCieApe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCodro As ADODB.Recordset
Public porstCoFjo As ADODB.Recordset
Dim cnn As ADODB.Connection
Private sBaseDatos As String
Dim aexcel As Boolean
Private Sub cmdAceptar_Click()
  'On Error GoTo Err
    'ini 2015-07-09 control flag mayoriza
    Dim zzMes As String, zzAnio As String
    If optProceso(0).Value Then
        zzMes = "13"
        zzAnio = gsAnoAct
    Else
        zzMes = "00"
        zzAnio = Format(Str(Val(gsAnoAct) + 1), "0000")
    End If
    Dim arr_cie() As Integer
    arr_cie() = gpCieMes_arr(zzAnio, zzMes) 'verificar solo un periodos
    If arr_cie(4) = 1 Then MsgBox TEXT_9016 & " " & zzMes & "/" & zzAnio, vbCritical: Exit Sub
    'fin 2015-07-09 control flag mayoriza

  
  Dim sMensaje As String
    
  ' Valido que todos los Parametros del Voucher esten correctos
  If txtDato(0).Text = "" Then Beep: MsgBox Choose(gsIdioma, "Debe Ingresar Diario donde se Grabará el Comprobante", "You must enter Journal where will save voucher"), vbExclamation: txtDato(0).SetFocus: Exit Sub
  If txtDato(1).Text = "" Then Beep: MsgBox Choose(gsIdioma, "Debe Ingresar Glosa del Comprobante", "You must enter Gloss of voucher"), vbExclamation: txtDato(1).SetFocus: Exit Sub
    
  sMensaje = IIf(optProceso(0).Value, Choose(gsIdioma, "Cerrar el Año '", "Close the Year '") & gsAnoAct & Choose(gsIdioma, "' y Generar el Comprobante de Cierre en el Periodo '", "' and Generate Closing Voucher in Period '") & cmbPeriodo & "'", Choose(gsIdioma, "Aperturar el Año '", "Open the Year '") & Trim$(Val(gsAnoAct) + 1) & Choose(gsIdioma, "' y Generar el Comprobante de Apertura", "' and Generate Opening Voucher")) & " ? "
  Beep
  ' Genero el comprobante de cierre/apertura de año
  
  If MsgBox(Choose(gsIdioma, "Estás Seguro de ", " Are you sure of ") & sMensaje, vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    pgbProgreso.Value = 0: pgbProgreso.Min = 0
     
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
    If optProceso(0).Value Then
      ' Paso 1: Realizo la exportacion de las tablas
      pgbProgreso.Max = 15
      pgbProgreso.Value = pgbProgreso.Min
      ppGenera_Cierre
    Else
      ' Paso 1: Realizo la exportacion de las tablas
      pgbProgreso.Max = 1
      pgbProgreso.Value = pgbProgreso.Min
      aexcel = False
      ppGenera_Apertura
    End If
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
    
    MsgBox TEXT_8008, vbInformation
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    cmdSalir.SetFocus
    End If
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdexcel_Click()
  'On Error GoTo Err
  
  Dim sMensaje As String
    
  ' Valido que todos los Parametros del Voucher esten correctos
  If txtDato(0).Text = "" Then Beep: MsgBox Choose(gsIdioma, "Debe Ingresar Diario donde se Grabará el Comprobante", "You must enter Journal where will save voucher"), vbExclamation: txtDato(0).SetFocus: Exit Sub
  If txtDato(1).Text = "" Then Beep: MsgBox Choose(gsIdioma, "Debe Ingresar Glosa del Comprobante", "You must enter Gloss of voucher"), vbExclamation: txtDato(1).SetFocus: Exit Sub
    
  sMensaje = IIf(optProceso(0).Value, Choose(gsIdioma, "Cerrar el Año '", "Close the Year '") & gsAnoAct & Choose(gsIdioma, "' y Generar el Comprobante de Cierre en el Periodo '", "' and Generate Closing Voucher in Period '") & cmbPeriodo & "'", Choose(gsIdioma, "Aperturar el Año '", "Open the Year '") & Trim$(Val(gsAnoAct) + 1) & Choose(gsIdioma, "' y Generar el Comprobante de Apertura", "' and Generate Opening Voucher")) & " ? "
  Beep
  ' Genero el comprobante de cierre/apertura de año
  
  'If MsgBox(Choose(gsIdioma, "Estás Seguro de ", " Are you sure of ") & sMensaje, vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    pgbProgreso.Value = 0: pgbProgreso.Min = 0
     
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
    If optProceso(0).Value Then
      ' Paso 1: Realizo la exportacion de las tablas
      'pgbProgreso.Max = 15
      'pgbProgreso.Value = pgbProgreso.Min
      'ppGenera_Cierre
    Else
      ' Paso 1: Realizo la exportacion de las tablas
      pgbProgreso.Max = 1
      pgbProgreso.Value = pgbProgreso.Min
      aexcel = True
      
      ppGenera_Apertura
    End If
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
    
    MsgBox TEXT_8008, vbInformation
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    cmdSalir.SetFocus
    'End If
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub


Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2
      ppAyuBus AYUDAT, Choose(Index + 1, 0, 3, 4)
      txtDato(Choose(Index + 1, 0, 3, 4)).SetFocus
   End Select
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4 ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case 3, 4                          'Cambiar (añadir índices).
         modAyuBus.Fjo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodFjo)=4 ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex - 2).Caption = " " & frmOAyuBus.uvDato2
      End Select
   End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex              'Cambiar.
      Case 0
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With porstCodro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
            End If
         End With
      Case 3, 4
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex - 2).Caption = ""
            Exit Function
         End If
         With porstCoFjo
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodFjo='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex - 2).Caption = " " & !DetFjo
            End If
         End With
      End Select
   End If
End Function

Private Sub ppGenera_Apertura()

  Dim sSentencia As String, sNroComprobante As String
  Dim nImpTCb As Double, nImpMN As Double, nImpME As Double
  Dim nImpTCb_Cpr As Double, nImpTCb_Vta As Double
  Dim nRegistro As Double, nNumRegistros As Double
  Dim nContador As Integer, sTabla As String, sCodFjo As String
  Dim sSeleccion As String, sPeriodo As String, sJoin As String, sOrden As String
  Dim porstTmp As ADODB.Recordset
  Dim porstBus As ADODB.Recordset
  Dim sTpoDebHab As String
  Dim nApertura As Integer, nDetalle As Integer
  
  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  Set porstBus = New ADODB.Recordset
  With porstTmp
    .ActiveConnection = pocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  With porstBus
    .ActiveConnection = pocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  
  ' Borro los comprobantes(cabecera y detalle) de apertura de año
  sSentencia = "DELETE FROM COCpbCab "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
  sSentencia = sSentencia & "AND MesPvs='" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sSentencia = sSentencia & "AND TpoGnr='" & TPOGNR_APE & "'"
  pocnnMain.Execute sSentencia, nNumRegistros
  ' Inicializo e inserto la información de las tablas maestras
  If chkProceso.Value = vbChecked Then
    ' Borro la información de las tablas generales
    For nContador = 11 To 1 Step -1
      sTabla = Choose(nContador, "CoDro", "CoCta", "CoCCo", "CoEfe", "CoFjo", "TgCfg", "CoCfg", "CoEFi", "CoEFiLin", "CoCCoCfg", "CoCieMes")
      sSentencia = "DELETE FROM " & sTabla & " "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "'"
      pocnnMain.Execute sSentencia, nNumRegistros
    Next nContador
    ' Importo la información de las tablas generales
    For nContador = 1 To 15
      sTabla = Choose(nContador, "CoDro", "CoCta", "TgAux", "TgAuxNat", "TgTDc", "CoCCo", "CoEfe", "CoFjo", "TgTcb", "TgCfg", "CoCfg", "CoEFi", "CoEFiLin", "CoCCoCfg", "CoCieMes")
'ini 2015-04-01 adicion colib y cobco
'      sSeleccion = Choose(nContador, "codemp, coddro, detdro, detdrox, usrcre, fyhcre", _
'                                     "codemp, codcta, detcta, detctax, tpocta, natcta, tposdo, tpoanl, codcta_dst_deb, codcta_dst_hab, codcco_dst_deb, codcco_dst_hab, tpomon, tpotcb, tpoajd, codcta_ajd_deb, codcta_ajd_hab, codcco_ajd_deb, codcco_ajd_hab, indajd, codcta_crr_deu,codcta_crr_acr, codcco_def, indcco, inddoc, indmoe, indpsp, indfjo, estcta, usrcre, fyhcre, usrmdf, fyhmdf",
'fin 2015-04-01 adicion colib y cobco
      sSeleccion = Choose(nContador, "codemp, coddro, detdro, detdrox, usrcre, fyhcre, codlib ", _
                                     "codemp, codcta, detcta, detctax, tpocta, natcta, tposdo, tpoanl, codcta_dst_deb, codcta_dst_hab, codcco_dst_deb, codcco_dst_hab, tpomon, tpotcb, tpoajd, codcta_ajd_deb, codcta_ajd_hab, codcco_ajd_deb, codcco_ajd_hab, indajd, codcta_crr_deu,codcta_crr_acr, codcco_def, indcco, inddoc, indmoe, indpsp, indfjo, estcta, usrcre, fyhcre, usrmdf, fyhmdf,codbco ", _
                                     "codemp, codaux, razaux, rucaux, tpodci, diraux, rubro, indcli, indprv, indotr, tpoper, estaux, usrcre, fyhcre, usrmdf, fyhmdf", "codemp, codaux, nomaux, apepataux, apemataux, usrcre, fyhcre, usrmdf, fyhmdf", _
                                     "codemp, codtdc, dettdc, dettdcx, abvtdc, sgntdc, forimp, usrcre, fyhcre, usrmdf, fyhmdf", _
                                     "codemp, codcco, detcco, detccox, estcco, usrcre, fyhcre, usrmdf, fyhmdf", _
                                     "codemp, codefe, detefe, detefex, tpoefe, usrcre, fyhcre, usrmdf, fyhmdf", "codemp, codfjo, detfjo, detfjox, tpofjo, codefe, usrcre, fyhcre, usrmdf, fyhmdf", _
                                     "codemp, fehtcb, imptcb_cpr, imptcb_vta, usrcre, fyhcre, usrmdf, fyhmdf", _
                                     "codemp, mesatu, pctigv,pctigv1,pctigv2, pctisc, pctir4, pcties, pctrtc, pctpcp, impuit, usrmdf_igv, fyhmdf_igv", "codemp, mesatu, tpomon_fnc, tpomon_sgn_mn, tpomon_sgn_me, codcta_nv3, codcta_nv4, codcta_nv5, codcta_nv6, codcta_nv7, codcta_nv8, codtdc_pcp, codtdc_rtc, codcta_pcp, codcta_rtc, indcco, indmne, indrtc, indpcp, codcco_nv3, codcco_nv5, tpoglo_rtc, glodocr_rtc, glodocn_rtc", _
                                     "codemp, codefi, detefi, detefix, indcnv, usrcre, fyhcre, usrmdf, fyhmdf", "codemp, codefi, nrolin, detlin, detlinx, tpolin, fmllin, bsepct, grppct, imp1, pct1, imp2, pct2, indlat, indbdesup, indbdeinf, indfondet, indfondet_syd, indfonimp, usrcre, fyhcre, usrmdf, fyhmdf", _
                                     "codemp, tipofmt, numord, codcfg,detcfg,codcco, detcco, nivel, usrcre, fyhcre, usrmdf, fyhmdf", "codemp, mescie")
      sJoin = Choose(nContador, "b.CodDro=a.CodDro", "b.CodCta=a.CodCta", "b.CodAux=a.CodAux", "b.CodAux=a.CodAux", "b.CodTdc=a.CodTdc", "b.CodCCo=a.CodCCo", "b.CodEfe=a.CodEfe", "b.CodFjo=a.CodFjo", "b.FehTcb=a.FehTcb", "b.mesatu=a.mesatu", "b.mesatu=a.mesatu", "b.CodEfi=a.CodEfi", "b.CodEfi=a.CodEfi AND b.NroLin=a.NroLin", "b.TipoFmt=a.TipoFmt AND b.NumOrd=a.NumOrd AND b.CodCCo=a.CodCCo", "b.MesCie=a.MesCie")
      sOrden = Choose(nContador, "a.CodDro", "a.CodCta", "a.CodAux", "a.CodAux", "a.CodTdc", "a.CodCCo", "a.CodEfe", "a.CodFjo", "a.FehTcb", "a.mesatu", "a.mesatu", "a.CodEfi", "a.CodEfi, a.NroLin", "a.TipoFmt, a.NumOrd, a.CodCCo", "a.MesCie")
      sPeriodo = Choose(nContador, "S", "S", "N", "N", "N", "S", "S", "S", "N", "S", "S", "S", "S", "S", "S")
      ' Inserto los registros no existentes
      sSentencia = "INSERT INTO " & sTabla & " (" & sSeleccion & IIf(sPeriodo = "S", ", pdoano) ", ") ")
      sSentencia = sSentencia & "SELECT  " & sSeleccion & IIf(sPeriodo = "S", ", '" & sBaseDatos & "' AS pdoano ", " ")
      sSentencia = sSentencia & "FROM " & sTabla & " a "
      sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & IIf(sPeriodo = "S", "AND a.pdoano='" & gsAnoAct & "' ", "")
      sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
      sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
      sSentencia = sSentencia & IIf(sPeriodo = "S", "AND b.pdoano='" & sBaseDatos & "' ", "")
      sSentencia = sSentencia & IIf(sJoin = "", "", "AND " & sJoin) & ") "
      sSentencia = sSentencia & "ORDER BY " & sOrden
      pocnnMain.Execute sSentencia, nNumRegistros
    Next nContador
  End If
  ' Elimino tabla temporal de cuenta cierre inventario
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpCuentas_') DROP TABLE #tmpCuentas"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpCuentas", sSentencia), nNumRegistros
  
  ' Genero tabla temporal de cuenta cierre inventario
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmpCuentas ", "")
  sSentencia = sSentencia & "SELECT DISTINCT cta.codemp, cta.pdoano, cta.codcta_crr_deu AS codcta_crr "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpCuentas ", "")
  sSentencia = sSentencia & "FROM cocta cta "
  sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cta.codcta_crr_deu, '')<>'' "
  sSentencia = sSentencia & "AND cta.TpoSdo='" & TPOSDO_INV & "' "
  sSentencia = sSentencia & "AND cta.IndMoe='" & INDAMO_INA & "' "
  sSentencia = sSentencia & "UNION "
  sSentencia = sSentencia & "SELECT DISTINCT cta.codemp, cta.pdoano, cta.codcta_crr_acr AS codcta_crr "
  sSentencia = sSentencia & "FROM cocta cta "
  sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cta.codcta_crr_acr, '')<>'' "
  sSentencia = sSentencia & "AND cta.TpoSdo='" & TPOSDO_INV & "' "
  sSentencia = sSentencia & "AND cta.IndMoe='" & INDAMO_INA & "' "
  sSentencia = sSentencia & "ORDER BY codcta_crr"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Elimino la tabla temporal de información
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 13)='#tmpApertura_') DROP TABLE #tmpApertura"
  'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpApertura" & nContador, sSentencia), nNumRegistros
    
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpApertura", "DROP TABLE IF EXISTS tmpApertura"), nNumRegistros
    
  ' Paso 1: Genero información de saldos cuentas con documentos
  ' sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE ", "CREATE TEMPORARY TABLE ") & ps_Prefijo & "tmpApertura ( "
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE ", "CREATE TABLE ") & ps_Prefijo & "tmpApertura ( "
  
  sSentencia = sSentencia & "codcta varchar(16) Not Null, "
  sSentencia = sSentencia & "codaux varchar(11) Null, "
  sSentencia = sSentencia & "codtdc char(2) Null, "
  sSentencia = sSentencia & "serdoc char(4) Null, "
  sSentencia = sSentencia & "nrodoc varchar(10) Null, "
  sSentencia = sSentencia & "feedoc " & IIf(ps_Plataforma = pSrvMySql, "date", "smalldatetime") & " default Null, "
  sSentencia = sSentencia & "fevdoc " & IIf(ps_Plataforma = pSrvMySql, "date", "smalldatetime") & " default Null, "
  sSentencia = sSentencia & "ferdoc " & IIf(ps_Plataforma = pSrvMySql, "date", "smalldatetime") & " default Null, "
  sSentencia = sSentencia & "cdebemn decimal(12,2) Not Null default '0.00', "
  sSentencia = sSentencia & "chabermn decimal(12,2) Not Null default '0.00', "
  sSentencia = sSentencia & "cdebeme decimal(12,2) Not Null default '0.00', "
  sSentencia = sSentencia & "chaberme decimal(12,2) Not Null default '0.00', "
  sSentencia = sSentencia & "codcta_dst_deb varchar(16) Null, "
  sSentencia = sSentencia & "codcta_dst_hab varchar(16) Null, "
  sSentencia = sSentencia & "tpomon char(1) Null, "
  sSentencia = sSentencia & "tpotcb char(1) Null, "
  sSentencia = sSentencia & "indfjo smallint Not Null default '0', "
  sSentencia = sSentencia & "codemp char(3) Not Null, "
  sSentencia = sSentencia & "pdoano char(4) Not Null, "
  sSentencia = sSentencia & "mespvs char(2) Not Null, "
  sSentencia = sSentencia & "gloite varchar(60) Null, "
  sSentencia = sSentencia & "gloitex varchar(60) Null, "
  sSentencia = sSentencia & "KEY IIdocumento (codemp, pdoano, mespvs, codcta, codaux, codtdc, serdoc, nrodoc))"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  sSentencia = "INSERT INTO " & ps_Prefijo & "tmpApertura "
  sSentencia = sSentencia & "SELECT b.CodCta AS CodCta, a.CodAux AS CodAux, a.CodTDc AS CodTDc, a.SerDoc AS SerDoc, a.NroDoc AS NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, MIN(a.FeRDoc) AS FeRDoc, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cDebeMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cHaberMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cDebeME, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cHaberME, "
  sSentencia = sSentencia & "b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb, b.IndFjo, "
  sSentencia = sSentencia & "a.codemp, a.pdoano, MIN(a.mespvs) AS mespvs, "
  sSentencia = sSentencia & "Null AS gloite, Null AS gloitex "
  sSentencia = sSentencia & "FROM ((CoCPbDet a "
  sSentencia = sSentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
  sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpCuentas c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta_Crr) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<='" & gfCeros(13 + cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(c.CodCta_crr, '')='' "
  sSentencia = sSentencia & "AND b.TpoSdo='" & TPOSDO_INV & "' "
  sSentencia = sSentencia & "AND b.IndDoc='" & INDDOC_ACT & "' "
  sSentencia = sSentencia & "AND b.IndMoe='" & INDAMO_INA & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTDc, '')<>'' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.SerDoc, '')<>'' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.NroDoc, '')<>'' "
  sSentencia = sSentencia & "GROUP BY a.codemp, a.pdoano, b.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb, b.IndFjo "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (ROUND(cDebeMN - cHaberMN, 2) <> 0.00) OR (ROUND(cDebeME - cHaberME, 2) <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00) "
    sSentencia = sSentencia & "OR (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  sSentencia = sSentencia & "ORDER BY CodCta, CodAux, CodTDc, SerDoc, NroDoc"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Actualizo glosa y fecha de emisión, vencimiento de documento (provisión)
  If ps_Plataforma = pSrvMySql Then
    sSentencia = "UPDATE tmpApertura tmp, cocpbdet det "
    sSentencia = sSentencia & "SET tmp.feedoc=det.feedoc, tmp.fevdoc=det.fevdoc, tmp.ferdoc=det.ferdoc, "
    sSentencia = sSentencia & "tmp.gloite=IFNULL(det.gloite, ''), tmp.gloitex=IFNULL(det.gloitex, '') "
  ElseIf ps_Plataforma = pSrvMySql Then
    sSentencia = "UPDATE #tmpApertura "
    sSentencia = sSentencia & "SET feedoc=det.feedoc, fevdoc=det.fevdoc, ferdoc=det.ferdoc, "
    sSentencia = sSentencia & "gloite=ISNULL(det.gloite, ''), gloitex=ISNULL(det.gloitex, '') "
    sSentencia = sSentencia & "FROM #tmpApertura tmp, cocpbdet det "
  End If
  sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND det.tpopvs='" & TPOPVS_PVS & "' "
  sSentencia = sSentencia & "AND tmp.codemp=det.codemp "
  sSentencia = sSentencia & "AND tmp.pdoano=det.pdoano "
  sSentencia = sSentencia & "AND tmp.mespvs=det.mespvs "
  sSentencia = sSentencia & "AND tmp.codcta=det.codcta "
  sSentencia = sSentencia & "AND tmp.codaux=det.codaux "
  sSentencia = sSentencia & "AND tmp.codtdc=det.codtdc "
  sSentencia = sSentencia & "AND tmp.serdoc=det.serdoc "
  sSentencia = sSentencia & "AND tmp.nrodoc=det.nrodoc"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Paso 2: Genero información de saldos cuentas con auxiliar
  sSentencia = "INSERT INTO " & ps_Prefijo & "tmpApertura "
  sSentencia = sSentencia & "SELECT b.CodCta AS CodCta, a.CodAux AS CodAux, '' AS CodTDc, '' AS SerDoc, '' AS NroDoc, Null AS FeEDoc, Null AS FeVDoc, Null AS FeRDoc, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cDebeMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cHaberMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cDebeME, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cHaberME, "
  sSentencia = sSentencia & "b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb, b.IndFjo, "
  sSentencia = sSentencia & "a.codemp, a.pdoano, MIN(a.mespvs) AS mespvs, '" & Choose(gsIdioma, txtDato(1).Text, txtDato(2).Text) & "' AS gloite, '" & Choose(gsIdioma, txtDato(2).Text, txtDato(1).Text) & "' AS gloitex "
  sSentencia = sSentencia & "FROM ((CoCPbDet a "
  sSentencia = sSentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
  sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpCuentas c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta_Crr) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<='" & gfCeros(13 + cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(c.CodCta_crr, '')='' "
  sSentencia = sSentencia & "AND b.TpoSdo='" & TPOSDO_INV & "' "
  sSentencia = sSentencia & "AND b.TpoAnl='" & TPOANL_AUX & "' "
  sSentencia = sSentencia & "AND b.IndDoc='" & INDDOC_INA & "' "
  sSentencia = sSentencia & "AND b.IndMoe='" & INDAMO_INA & "' "
  sSentencia = sSentencia & "GROUP BY a.codemp, a.pdoano, b.CodCta, a.Codaux, b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb, b.IndFjo "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (ROUND(cDebeMN - cHaberMN, 2) <> 0.00) OR (ROUND(cDebeME - cHaberME, 2) <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00) "
    sSentencia = sSentencia & "OR (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  sSentencia = sSentencia & "ORDER BY CodCta, CodAux, CodTDc, SerDoc, NroDoc"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Paso 3: Genero información de saldos cuentas
  sSentencia = "INSERT INTO " & ps_Prefijo & "tmpApertura "
  sSentencia = sSentencia & "SELECT b.CodCta AS CodCta, '' AS CodAux, '' AS CodTDc, '' AS SerDoc, '' AS NroDoc, Null AS FeEDoc, Null AS FeVDoc, Null AS FeRDoc, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cDebeMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cHaberMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cDebeME, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cHaberME, "
  sSentencia = sSentencia & "b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb, b.IndFjo, "
  sSentencia = sSentencia & "a.codemp, a.pdoano, MIN(a.mespvs) AS mespvs, '" & Choose(gsIdioma, txtDato(1).Text, txtDato(2).Text) & "' AS gloite, '" & Choose(gsIdioma, txtDato(2).Text, txtDato(1).Text) & "' AS gloitex "
  sSentencia = sSentencia & "FROM ((CoCPbDet a "
  sSentencia = sSentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
  sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpCuentas c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta_Crr) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<='" & gfCeros(13 + cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(c.CodCta_crr, '')='' "
  sSentencia = sSentencia & "AND b.TpoSdo='" & TPOSDO_INV & "' "
  sSentencia = sSentencia & "AND b.TpoAnl<>'" & TPOANL_AUX & "' "
  sSentencia = sSentencia & "AND b.IndDoc='" & INDDOC_INA & "' "
  sSentencia = sSentencia & "AND b.IndMoe='" & INDAMO_INA & "' "
  sSentencia = sSentencia & "GROUP BY b.CodCta, b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb, b.IndFjo "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (ROUND(cDebeMN - cHaberMN, 2) <> 0.00) OR (ROUND(cDebeME - cHaberME, 2) <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00) "
    sSentencia = sSentencia & "OR (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  sSentencia = sSentencia & "ORDER BY CodCta, CodAux, CodTDc, SerDoc, NroDoc"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Seleciono la información de apertura
  With porstTmp
    If .State = adStateOpen Then .Close
    .Source = "SELECT * FROM " & ps_Prefijo & "tmpApertura "
    .Source = .Source & "ORDER BY CodCta, CodAux, CodTDc, SerDoc, NroDoc"
    .Open
  End With
  
  
  'Para enviar a Excel
  If aexcel = False Then
  
  
  pgbProgreso.Max = IIf(porstTmp.RecordCount > 1, porstTmp.RecordCount, 1)
  If Not (porstTmp.EOF And porstTmp.BOF) Then
    'Obtengo el número e inserto la cabecera del comprobante.
    sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
    sSentencia = sSentencia & "FROM COCpbCab "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
    sSentencia = sSentencia & "AND CodDro='" & txtDato(0).Text & "' "
    sSentencia = sSentencia & "AND MesPvs='" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "'"
    With porstBus
      If .State = adStateOpen Then .Close
      .Source = sSentencia
      .Open
      sNroComprobante = !cUltNroCpb
      .Close
    End With
    sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
    ' Grabación de cabecera de comprobante
    sSentencia = "INSERT INTO CoCpbCab(codemp, pdoano, MesPvs, CodDro, NroCpb, FehCpb, GloCpb, glocpbx, TpoGnr, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf)"
    sSentencia = sSentencia & " VALUES("
    sSentencia = sSentencia & "'" & gsCodEmp & "', "
    sSentencia = sSentencia & "'" & sBaseDatos & "', "
    sSentencia = sSentencia & "'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "', "
    sSentencia = sSentencia & "'" & txtDato(0).Text & "', "
    sSentencia = sSentencia & "'" & sNroComprobante & "', "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
    sSentencia = sSentencia & IIf(txtDato(gsIdioma).Text = "", "Null", "'" & txtDato(gsIdioma).Text & "'") & ", "
    sSentencia = sSentencia & IIf(txtDato(3 - gsIdioma).Text = "", "Null", "'" & txtDato(3 - gsIdioma).Text & "'") & ", "
    sSentencia = sSentencia & "'" & TPOGNR_APE & "', "
    sSentencia = sSentencia & "'" & INDNCU_FAL & "', "
    sSentencia = sSentencia & "'" & INDANU_FAL & "', "
    sSentencia = sSentencia & "'" & gsAbvUsr & "', "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
    sSentencia = sSentencia & "Null, Null)"
    pocnnMain.Execute sSentencia, nNumRegistros
    ' Actualizo el numero de comprobantes en los diarios
    sSentencia = "UPDATE CoDro SET Cpb" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "='" & sNroComprobante & "' "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
    sSentencia = sSentencia & "AND CodDro='" & txtDato(0).Text & "'"
    pocnnMain.Execute sSentencia, nNumRegistros
   
    ' Obtengo el tipo de cambio de la fecha
    sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ImpTCb_Cpr,1) AS cImpTCb_Cpr, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ImpTCb_Vta, 1) AS cImpTCb_Vta "
    sSentencia = sSentencia & "FROM TGTCb "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND DATE_FORMAT(FehTCb,'%d/%m/%Y')=DATE_FORMAT('" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & "', '%d/%m/%Y')"
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND CONVERT(smalldatetime, FehTCb, 103)=CONVERT(smalldatetime, '" & Format(dtpFehCpb.Value, "dd/mm/yyyy") & "', 103)"
    End If
    With porstBus
      If .State = adStateOpen Then .Close
      .Source = sSentencia
      .Open
      nImpTCb_Cpr = IIf(.EOF, 1, !cImpTCb_Cpr)
      nImpTCb_Vta = IIf(.EOF, 1, !cImpTCb_Vta)
      .Close
    End With
    porstTmp.MoveFirst
    nRegistro = 0
    While Not porstTmp.EOF
      nImpTCb = Format(Val(IIf(porstTmp!TpoTcb = TPOTCB_VTA, nImpTCb_Vta, nImpTCb_Cpr)), FORMATO_NUM_2)
      nImpMN = Format(CDec(porstTmp!cDebeMN) - CDec(porstTmp!cHaberMN), FORMATO_NUM_1)
      nImpME = Format(CDec(porstTmp!cDebeME) - CDec(porstTmp!cHaberME), FORMATO_NUM_1)
      sCodFjo = IIf(porstTmp!IndFjo = INDFJO_ACT, IIf((nImpMN > 0 Or nImpME > 0), txtDato(3), txtDato(4)), "")
      nApertura = IIf((nImpMN >= 0 And nImpME >= 0) Or (nImpMN <= 0 And nImpME <= 0), 1, 2)
      For nDetalle = 1 To nApertura
       nRegistro = nRegistro + 1
        ' Obtengo los importes
        nImpMN = Format(IIf(nApertura = 2 And nDetalle = 2, 0, (CDec(porstTmp!cDebeMN) - CDec(porstTmp!cHaberMN))), FORMATO_NUM_1)
        nImpME = Format(IIf((nApertura = 2 And nDetalle = 1), 0, (CDec(porstTmp!cDebeME) - CDec(porstTmp!cHaberME))), FORMATO_NUM_1)
        sTpoDebHab = IIf(IIf(nApertura = 2, Choose(nDetalle, (nImpMN > 0), (nImpME > 0)), (nImpMN > 0 Or nImpME > 0)), TPOCTB_DEB, TPOCTB_HAB)
       ' Grabo detalle de comprobante
       ppGrabo_Detalle sNroComprobante, nRegistro, sCodFjo, porstTmp!codtdc, porstTmp!codcta, IIf(IsNull(porstTmp!codaux), "", porstTmp!codaux), porstTmp!serdoc, porstTmp!nrodoc, Format(porstTmp!feedoc, "dd/mm/yyyy"), Format(porstTmp!fevdoc, "dd/mm/yyyy"), Format(porstTmp!ferdoc, "dd/mm/yyyy"), IIf(IsNull(porstTmp!GloIte), "", porstTmp!GloIte), IIf(IsNull(porstTmp!GloItex), "", porstTmp!GloItex), sTpoDebHab, porstTmp!tpomon, porstTmp!TpoTcb, nImpTCb, Abs(nImpMN), Abs(nImpME)
      Next nDetalle
      pgbProgreso.Value = IIf(pgbProgreso.Max >= nRegistro, nRegistro, pgbProgreso.Max)
      DoEvents
      porstTmp.MoveNext
    Wend
  End If
  
  End If
  
  ' Elimino la tabla temporal de cuentas
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpCuentas_') DROP TABLE #tmpCuentas"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpCuentas", sSentencia), nNumRegistros
  
  'Para enviar a Excel
  If aexcel = True Then
    exportar
  End If
  
  ' Elimino la tabla temporal de apertura
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 13)='#tmpApertura_') DROP TABLE #tmpApertura"
  
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpApertura" & nContador, sSentencia), nNumRegistros
  porstTmp.Close: Set porstTmp = Nothing
  Set porstBus = Nothing

End Sub

Private Sub ppGenera_Cierre()
  Dim sSentencia As String, sNroComprobante As String
  Dim nContador As Integer
  Dim sCodWhere As String, sMsgWhere As String, sMsgWherex As String
  Dim sCodCta_Crr As String, sCodCta_Crr_Deu As String, sCodCta_Crr_Acr As String
  Dim nImpTCb As Double, nImpMN As Double, nImpME As Double
  Dim nDebeMN As Double, nDebeME As Double
  Dim nHaberMN As Double, nHaberME As Double
  Dim nImpTCb_Cpr As Double, nImpTCb_Vta As Double
  Dim nRegistro As Double, nNumRegistros As Double
  Dim sTpoDebHab As String
  Dim porstTmp As ADODB.Recordset
  Dim porstBus As ADODB.Recordset
  Dim nCierre As Integer, nDetalle As Integer
  Dim sWherePeriodo As String

  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  Set porstBus = New ADODB.Recordset
  With porstTmp
    .ActiveConnection = pocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  With porstBus
    .ActiveConnection = pocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  ' Inicializo la condicion periodos
  sWherePeriodo = "AND a.pdoano='" & sBaseDatos & "' "
  sWherePeriodo = sWherePeriodo & "AND a.MesPvs<='" & Format(13 + cmbPeriodo.ListIndex, "00") & "'"
  
  ' Borro los comprobantes(cabecera y detalle) de cierre del periodo
  sSentencia = "DELETE FROM COCpbCab "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
  sSentencia = sSentencia & "AND MesPvs='" & Format(13 + cmbPeriodo.ListIndex, "00") & "' "
  sSentencia = sSentencia & "AND TpoGnr='" & TPOGNR_CIE & "'"
  pocnnMain.Execute sSentencia, nNumRegistros
  For nContador = 1 To 15
    sCodWhere = Choose(nContador, "9", "6", "7", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "59", "05")
    sMsgWhere = Choose(nContador, "CANCELACION DE CUENTAS CONTABILIDAD ANALITICA", "TRANSFERENCIA COSTO VENTAS A VARIACION EXISTENCIAS", "TRANSFERENCIA COSTO VENTAS A VARIACION EXISTENCIAS", "DETERMINACION DEL MARGEN COMERCIAL", txtDato(gsIdioma).Text, "DETERMINACION DEL VALOR AGREGADO", "DETERMINACION DEL EXCEDENTE BRUTO DE EXPLOTACION", "DETERMINACION DEL RESULTADO DE EXPLOTACION", txtDato(gsIdioma).Text, txtDato(gsIdioma).Text, "DETERMINACION DEL REI A LA RENTA", txtDato(gsIdioma).Text, "RESULTADO DEL PERIODO", "TRANSF. DEL RESULTADO A RESULTADOS ACUMULADOS", "CIERRE DE CUENTAS DE BALANCE")
    sMsgWherex = Choose(nContador, "CANCELACION DE CUENTAS CONTABILIDAD ANALITICA", "TRANSFERENCIA COSTO VENTAS A VARIACION EXISTENCIAS", "TRANSFERENCIA COSTO VENTAS A VARIACION EXISTENCIAS", "DETERMINACION DEL MARGEN COMERCIAL", txtDato(3 - gsIdioma).Text, "DETERMINACION DEL VALOR AGREGADO", "DETERMINACION DEL EXCEDENTE BRUTO DE EXPLOTACION", "DETERMINACION DEL RESULTADO DE EXPLOTACION", txtDato(3 - gsIdioma).Text, txtDato(3 - gsIdioma).Text, "DETERMINACION DEL REI A LA RENTA", txtDato(3 - gsIdioma).Text, "RESULTADO DEL PERIODO", "TRANSF. DEL RESULTADO A RESULTADOS ACUMULADOS", "CIERRE DE CUENTAS DE BALANCE")
    sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & IIf(nContador > 14, "''", "b.CodCta") & ",'') AS cCuenta, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cDebeMN, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS cHaberMN, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cDebeME, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS cHaberME, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Crr_deu, '') AS cCuenta_Crr_deu, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Crr_acr, '') AS cCuenta_Crr_acr, "
    sSentencia = sSentencia & "b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, "
    sSentencia = sSentencia & "b.TpoMon, b.TpoTCb "
    sSentencia = sSentencia & "FROM (cocpbdet a LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.codcta=b.codcta) "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & sWherePeriodo & " "
    If nContador = 1 Then
      sSentencia = sSentencia & "AND (LEFT(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta,''), " & Len(sCodWhere) & ")='" & sCodWhere & "' "
      sSentencia = sSentencia & "OR LEFT(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta,''), 2) IN('78', '79')) "
    ElseIf nContador > 14 Then
      sSentencia = sSentencia & "AND LEFT(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.codCta_crr_deu,''), 1)>='0' "
      sSentencia = sSentencia & "AND LEFT(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.codcta_crr_deu,''), 1)<='5' "
    Else
      sSentencia = sSentencia & "AND LEFT(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.codcta_crr_deu,''), " & Len(sCodWhere) & ")='" & sCodWhere & "' "
    End If
    sSentencia = sSentencia & "AND b.TpoSdo" & IIf(nContador > 14, "='", "<>'") & TPOSDO_INV & "' "
    sSentencia = sSentencia & "GROUP BY b.codcta" & IIf(nContador > 14, "_crr_deu, b.codcta_crr_acr", "") & ", "
    sSentencia = sSentencia & "b.CodCta_Dst_Deb, b.CodCta_Dst_Hab, b.TpoMon, b.TpoTCb" & IIf(nContador > 14, "", ", b.codcta_crr_deu, b.codcta_crr_acr") & " "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "HAVING (ROUND(cDebeMN - cHaberMN, 2) <> 0.00) OR (ROUND(cDebeME - cHaberME, 2) <> 0.00) "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "HAVING (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00) "
      sSentencia = sSentencia & "OR (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
    End If
    sSentencia = sSentencia & "ORDER BY cCuenta_Crr_deu, cCuenta_Crr_acr, cCuenta"
    With porstTmp
      If .State = adStateOpen Then .Close
      .Source = sSentencia
      .Open
    End With
    If Not (porstTmp.BOF And porstTmp.EOF) Then
      ' Obtengo el número e inserto la cabecera del comprobante
      sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
      sSentencia = sSentencia & "FROM COCpbCab "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
      sSentencia = sSentencia & "AND CodDro='" & txtDato(0).Text & "' "
      sSentencia = sSentencia & "AND MesPvs='" & gfCeros(13 + cmbPeriodo.ListIndex, 2, 0, "0") & "'"
      With porstBus
        If .State = adStateOpen Then .Close
        .Source = sSentencia
        .Open
        sNroComprobante = !cUltNroCpb
        .Close
      End With
      sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
      ' Grabación de cabecera de comprobante
      sSentencia = "INSERT INTO CoCpbCab(codemp, pdoano, CodDro, NroCpb, MesPvs, FehCpb, GloCpb, Glocpbx, TpoGnr, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf)"
      sSentencia = sSentencia & " VALUES("
      sSentencia = sSentencia & "'" & gsCodEmp & "', "
      sSentencia = sSentencia & "'" & sBaseDatos & "', "
      sSentencia = sSentencia & "'" & txtDato(0).Text & "', "
      sSentencia = sSentencia & "'" & sNroComprobante & "', "
      sSentencia = sSentencia & "'" & gfCeros(13 + cmbPeriodo.ListIndex, 2, 0, "0") & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(sMsgWhere = "", "Null", "'" & sMsgWhere & "'") & ", "
      sSentencia = sSentencia & IIf(sMsgWherex = "", "Null", "'" & sMsgWherex & "'") & ", "
      sSentencia = sSentencia & "'" & TPOGNR_CIE & "', "
      sSentencia = sSentencia & "'" & INDNCU_FAL & "', "
      sSentencia = sSentencia & "'" & INDANU_FAL & "', "
      sSentencia = sSentencia & "'" & gsAbvUsr & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
      sSentencia = sSentencia & "Null, Null)"
      pocnnMain.Execute sSentencia, nNumRegistros
      ' Actualizoel numero de comprobantes en los diarios
      sSentencia = "UPDATE CoDro SET Cpb" & gfCeros(13 + cmbPeriodo.ListIndex, 2, 0, "0") & "='" & sNroComprobante & "' "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
      sSentencia = sSentencia & "AND CodDro='" & txtDato(0).Text & "'"
      pocnnMain.Execute sSentencia, nNumRegistros
     
     'Obtengo el tipo de cambio de la fecha
      sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ImpTCb_Cpr,1) AS cImpTCb_Cpr, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ImpTCb_Vta, 1) AS cImpTCb_Vta "
      sSentencia = sSentencia & "FROM TGTCb "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & "AND DATE_FORMAT(FehTCb,'%d/%m/%Y')=DATE_FORMAT('" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & "', '%d/%m/%Y')"
      ElseIf ps_Plataforma = pSrvSql Then
        sSentencia = sSentencia & "AND CONVERT(smalldatetime, FehTCb, 103)=CONVERT(smalldatetime, '" & Format(dtpFehCpb.Value, "dd/mm/yyyy") & "', 103)"
      End If
      With porstBus
        If .State = adStateOpen Then .Close
        .Source = sSentencia
        .Open
        nImpTCb_Cpr = IIf(.EOF, 1, !cImpTCb_Cpr)
        nImpTCb_Vta = IIf(.EOF, 1, !cImpTCb_Vta)
        .Close
      End With
      porstTmp.MoveFirst
      nRegistro = 0
      While Not porstTmp.EOF
        sCodCta_Crr_Deu = porstTmp!ccuenta_crr_deu
        sCodCta_Crr_Acr = porstTmp!ccuenta_crr_acr
        sCodCta_Crr = porstTmp!ccuenta_crr_deu
        nDebeMN = 0: nDebeME = 0
        nHaberMN = 0: nHaberME = 0
        Do While (porstTmp!ccuenta_crr_deu = sCodCta_Crr_Deu And porstTmp!ccuenta_crr_acr = sCodCta_Crr_Acr)
          nImpTCb = Format(Val(IIf(porstTmp!TpoTcb = TPOTCB_VTA, nImpTCb_Vta, nImpTCb_Cpr)), FORMATO_NUM_2)
          nImpMN = Format(CDec(porstTmp!cDebeMN) - CDec(porstTmp!cHaberMN), FORMATO_NUM_1)
          nImpME = Format(CDec(porstTmp!cDebeME) - CDec(porstTmp!cHaberME), FORMATO_NUM_1)
          nDebeMN = nDebeMN + CDec(porstTmp!cDebeMN)
          nHaberMN = nHaberMN + CDec(porstTmp!cHaberMN)
          nDebeME = nDebeME + CDec(porstTmp!cDebeME)
          nHaberME = nHaberME + CDec(porstTmp!cHaberME)
          If porstTmp!cCuenta <> "" Then
            nCierre = IIf((nImpMN >= 0 And nImpME >= 0) Or (nImpMN <= 0 And nImpME <= 0), 1, 2)
            For nDetalle = 1 To nCierre
              nRegistro = nRegistro + 1
              ' Obtengo los importes
              nImpMN = Format(IIf(nCierre = 2 And nDetalle = 2, 0, (CDec(porstTmp!cDebeMN) - CDec(porstTmp!cHaberMN))), FORMATO_NUM_1)
              nImpME = Format(IIf((nCierre = 2 And nDetalle = 1), 0, (CDec(porstTmp!cDebeME) - CDec(porstTmp!cHaberME))), FORMATO_NUM_1)
              sTpoDebHab = IIf(IIf(nCierre = 2, Choose(nDetalle, (nImpMN < 0), (nImpME < 0)), (nImpMN < 0 Or nImpME < 0)), TPOCTB_DEB, TPOCTB_HAB)
              ' Grabo detalle de comprobante
              ppGrabo_Detalle sNroComprobante, nRegistro, "", "", porstTmp!cCuenta, "", "", "", "", "", "", "", "", sTpoDebHab, porstTmp!tpomon, porstTmp!TpoTcb, nImpTCb, Abs(nImpMN), Abs(nImpME)
            Next nDetalle
          End If
          DoEvents
          porstTmp.MoveNext
          If porstTmp.EOF Then Exit Do
        Loop
        ' Grabo detalle de comprobante cuenta de cierre
        If sCodCta_Crr <> "" Then
          nImpMN = Format(CDec(nDebeMN) - CDec(nHaberMN), FORMATO_NUM_1)
          nImpME = Format(CDec(nDebeME) - CDec(nHaberME), FORMATO_NUM_1)
          nCierre = IIf((nImpMN >= 0 And nImpME >= 0) Or (nImpMN <= 0 And nImpME <= 0), 1, 2)
          For nDetalle = 1 To nCierre
            nRegistro = nRegistro + 1
            ' Obtengo los importes
            nImpMN = Format(IIf(nCierre = 2 And nDetalle = 2, 0, (CDec(nDebeMN) - CDec(nHaberMN))), FORMATO_NUM_1)
            nImpME = Format(IIf((nCierre = 2 And nDetalle = 1), 0, (CDec(nDebeME) - CDec(nHaberME))), FORMATO_NUM_1)
            sTpoDebHab = IIf(IIf(nCierre = 2, Choose(nDetalle, (nImpMN > 0), (nImpME > 0)), (nImpMN > 0 Or nImpME > 0)), TPOCTB_DEB, TPOCTB_HAB)
            sTpoDebHab = IIf((nContador > 14), IIf(sTpoDebHab = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB), sTpoDebHab)
            ' cuenta cierre
            sCodCta_Crr = IIf(sTpoDebHab = "D", sCodCta_Crr_Deu, sCodCta_Crr_Acr)
            ' Obtengo los datos de la cuenta de cierre
            sSentencia = "SELECT TpoMon, TpoTcb "
            sSentencia = sSentencia & "FROM CoCta "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & sBaseDatos & "' "
            sSentencia = sSentencia & "AND CodCta='" & sCodCta_Crr & "'"
            With porstBus
              If .State = adStateOpen Then .Close
              .Source = sSentencia
              .Open
            End With
            nImpTCb = Format(Val(IIf(porstBus!TpoTcb = TPOTCB_VTA, nImpTCb_Vta, nImpTCb_Cpr)), FORMATO_NUM_2)
            ' Grabo detalle de comprobante
            ppGrabo_Detalle sNroComprobante, nRegistro, "", "", sCodCta_Crr, "", "", "", "", "", "", "", "", sTpoDebHab, porstBus!tpomon, porstBus!TpoTcb, nImpTCb, Abs(nImpMN), Abs(nImpME)
            porstBus.Close
          Next nDetalle
        End If
      Wend
    End If
    pgbProgreso.Value = nContador
  Next nContador
  porstTmp.Close: Set porstTmp = Nothing
  Set porstBus = Nothing

End Sub

Private Sub ppGrabo_Detalle(sNroCpb As String, nNroIte As Double, sCodFlujo As String, sCodTDc As String, sCodCta As String, sCodAux As String, sSerDoc As String, sNroDoc As String, sFecEmi As String, sFecVen As String, sFecRec As String, sGloite As String, sGloitex As String, sTpoCtb As String, sTpoMon As String, sTpoTcb As String, nImpTCb As Double, nImpMN As Double, nImpME As Double)

  Dim sSentencia As String, nNumRegistros As Double
  Dim nPeriodo As Integer
  Dim INDMASFJO_INI As Byte, INDMASFJO_MAS As Byte
  
  sGloite = IIf(sGloite = "", txtDato(gsIdioma).Text, sGloite)
  sGloitex = IIf(sGloitex = "", txtDato(3 - gsIdioma).Text, sGloitex)
  nPeriodo = IIf(optProceso(0).Value, 13, 0)
  INDMASFJO_INI = 0: INDMASFJO_MAS = 1
  ' Grabación de detalle de comprobante
  sSentencia = "INSERT INTO CoCpbDet(codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, SerDoc, NroDoc, FeEDoc, FeVDoc, "
  sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, TpoCtb, TpoPvs, TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, IndFjo_Det, UsrCre, FyHCre, UsrMdf, FyHMdf) "
  sSentencia = sSentencia & "VALUES("
  sSentencia = sSentencia & "'" & gsCodEmp & "', "
  sSentencia = sSentencia & "'" & sBaseDatos & "', "
  sSentencia = sSentencia & "'" & txtDato(0).Text & "', "
  sSentencia = sSentencia & "'" & sNroCpb & "', "
  sSentencia = sSentencia & "'" & nNroIte & "', "
  sSentencia = sSentencia & "'" & gfCeros(nPeriodo + cmbPeriodo.ListIndex, 2, 0, "0") & "', "
  sSentencia = sSentencia & "'" & nNroIte & "', "
  sSentencia = sSentencia & IIf(sCodTDc = "", "Null", "'" & sCodTDc & "'") & ", "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "'" & sCodCta & "', "
  sSentencia = sSentencia & "Null, "
  sSentencia = sSentencia & IIf(sCodAux = "", "Null", "'" & sCodAux & "'") & ", "
  sSentencia = sSentencia & IIf(sSerDoc = "", "Null", "'" & sSerDoc & "'") & ", "
  sSentencia = sSentencia & IIf(sNroDoc = "", "Null", "'" & sNroDoc & "'") & ", "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(IIf(sFecEmi <> "", sFecEmi, dtpFehCpb.Value), "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(IIf(sFecVen <> "", sFecVen, dtpFehCpb.Value), "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(IIf(sFecRec <> "", sFecRec, dtpFehCpb.Value), "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "Null, "
  sSentencia = sSentencia & IIf(sGloite = "", "Null", "'" & sGloite & "'") & ", "
  sSentencia = sSentencia & IIf(sGloitex = "", "Null", "'" & sGloitex & "'") & ", "
  sSentencia = sSentencia & "'" & sTpoCtb & "', "
  sSentencia = sSentencia & "'" & IIf(sSerDoc = "", TPOPVS_OTR, TPOPVS_PVS) & "', "
  sSentencia = sSentencia & "'" & sTpoMon & "', "
  sSentencia = sSentencia & "'" & sTpoTcb & "', "
  sSentencia = sSentencia & nImpTCb & ", "
  sSentencia = sSentencia & nImpMN & ", "
  sSentencia = sSentencia & nImpME & ", "
  sSentencia = sSentencia & "'" & IIf(optProceso(0).Value, TPOGNR_CIE, TPOGNR_APE) & "', "
  sSentencia = sSentencia & IIf(sCodFlujo = "", INDMASFJO_INI, INDMASFJO_MAS) & ", "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null)"
  pocnnMain.Execute sSentencia, nNumRegistros
  
  If sCodFlujo <> "" Then
    ' Grabación de detalle de flujo de caja
    sSentencia = "INSERT INTO CoCpbDetFjo(codemp, pdoano, mespvs, coddro, nrocpb, NroIte, nroord, codfjo,"
    sSentencia = sSentencia & " codcta, TpoCtb, ImpMN, ImpME, UsrCre, FyHCre, UsrMdf, FyHMdf)"
    sSentencia = sSentencia & " VALUES("
    sSentencia = sSentencia & "'" & gsCodEmp & "',"
    sSentencia = sSentencia & " '" & sBaseDatos & "',"
    sSentencia = sSentencia & " '" & gfCeros(nPeriodo + cmbPeriodo.ListIndex, 2, 0, "0") & "',"
    sSentencia = sSentencia & " '" & txtDato(0).Text & "',"
    sSentencia = sSentencia & " '" & sNroCpb & "',"
    sSentencia = sSentencia & " '" & nNroIte & "',"
    sSentencia = sSentencia & " '1',"
    sSentencia = sSentencia & " '" & sCodFlujo & "',"
    sSentencia = sSentencia & " '" & sCodCta & "',"
    sSentencia = sSentencia & " '" & sTpoCtb & "',"
    sSentencia = sSentencia & " " & nImpMN & ","
    sSentencia = sSentencia & " " & nImpME & ","
    sSentencia = sSentencia & " '" & gsAbvUsr & "',"
    sSentencia = sSentencia & " '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
    sSentencia = sSentencia & " Null, Null)"
    pocnnMain.Execute sSentencia, nNumRegistros
  End If
  
End Sub

Private Sub Form_Activate()
cmdSalir.SetFocus
End Sub

Private Sub Form_Load()

Set cnn = New ADODB.Connection
cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";connection="
cnn.CursorLocation = adUseClient
cnn.Open


  'Abrir Tablas.
   Set pocnnMain = New ADODB.Connection
   Set porstCodro = New ADODB.Recordset
   Set porstCoFjo = New ADODB.Recordset

   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstCodro
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
      .Source = .Source & "FROM CODro "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4 "
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With porstCoFjo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodFjo, DetFjo "
      .Source = .Source & "FROM COFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodFjo)=4 "
      .Source = .Source & "ORDER BY CodFjo"
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(6, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Periodo Comprobante :", "Diario de Comprobante :", "Fecha de Comprobante :", "Glosa de Comprobante - Traducción :", "Flujo de Caja Deudor :", "Flujo de Caja Acreedor :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Period Voucher :", "Journal of Voucher :", "Date of Voucher :", "Gloss of Voucher - Translation :", "Debtor Cash Flow :", "Creditor Cash Flow :")
  Next nElemento
  frmCuadro(0).Caption = Choose(gsIdioma, " Tipo de Proceso ", " Type of Process ")
  chkProceso.Caption = Choose(gsIdioma, "Inicializa Tablas Maestras", "Initialize Masters Tables")
  frmCuadro(1).Caption = Choose(gsIdioma, " Parametros ", " Parameters ")
  frmCuadro(2).Caption = Choose(gsIdioma, " Flujo de Caja ", " Cash Flow ")
  lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Información...", "Processing Information...")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
  
  ' Cambio los Mensajes de Apertura/Cierre
  optProceso(0).Caption = Choose(gsIdioma, "Cierre de Año ", "Closing year ") & gsAnoAct
  optProceso(1).Caption = Choose(gsIdioma, "Apertura de Año ", "Opening year ") & Trim$(Val(gsAnoAct) + 1)
  optProceso(0).Value = vbChecked

End Sub

Private Sub Form_Unload(Cancel As Integer)

porstCodro.Close
porstCoFjo.Close
pocnnMain.Close
Set porstCodro = Nothing
Set porstCoFjo = Nothing
Set pocnnMain = Nothing

End Sub

Private Sub optProceso_Click(Index As Integer)

  txtDato(0).Text = "": lblDatoDeta(0) = ""
  sBaseDatos = gsAnoAct
  cmbPeriodo.Clear
  
  If Index = 0 Then
     txtDato(0).Text = "2001" '2015-05-18 validacion frm
    cmdAceptar.ToolTipText = Choose(gsIdioma, "Genera Comprobante de Cierre de Año", "Generate Closing Year Voucher")
    frmCuadro(1).Caption = Choose(gsIdioma, "Datos del Comprobante de Cierre de Año", "Data of Closing Year Voucher")
    txtDato(1).Text = Choose(gsIdioma, "COMPROBANTE DE CIERRE", "CLOSING VOUCHER")
    txtDato(2).Text = Choose(gsIdioma, "CLOSING VOUCHER", "COMPROBANTE DE CIERRE")
    With dtpFehCpb
      .MinDate = CDate("01/" & gsMesCie & "/" & gsAnoAct)
      .MaxDate = gfUltDia(.MinDate)
      .Value = .MaxDate
    End With
    cmbPeriodo.AddItem Choose(gsIdioma, "Cierre 1 ", "Closing 1 ") & gsAnoAct
    cmbPeriodo.AddItem Choose(gsIdioma, "Cierre 2 ", "Closing 2 ") & gsAnoAct
  Else
    cmdAceptar.ToolTipText = Choose(gsIdioma, "Genera Comprobante de Apertura de Año", "Generate Opening Year Voucher")
    frmCuadro(1).Caption = Choose(gsIdioma, "Datos del Comprobante de Apertura de Año", "Data of Opening Year Voucher")
    txtDato(1).Text = Choose(gsIdioma, "COMPROBANTE DE APERTURA", "OPENING VOUCHER")
    txtDato(2).Text = Choose(gsIdioma, "OPENING VOUCHER", "COMPROBANTE DE APERTURA")
    With dtpFehCpb
      .MaxDate = gfUltDia("01/" & gsMesApe & "/" & Trim$(Val(gsAnoAct) + IIf(gnFrances = BSEPCT_ACT, 0, 1)))
      .MinDate = CDate("01/" & gsMesApe & "/" & Trim$(Val(gsAnoAct) + IIf(gnFrances = BSEPCT_ACT, 0, 1)))
      '2015-05-18 validacion frm  .Value = CDate("02/" & gsMesApe & "/" & Trim$(Val(gsAnoAct) + IIf(gnFrances = BSEPCT_ACT, 0, 1)))
      .Value = CDate("01/" & gsMesApe & "/" & Trim$(Val(gsAnoAct) + IIf(gnFrances = BSEPCT_ACT, 0, 1)))
    End With
    cmbPeriodo.AddItem Choose(gsIdioma, "Apertura ", "Opening ") & Trim$(Val(gsAnoAct) + 1)
    sBaseDatos = Trim$(Val(gsAnoAct) + IIf(gnFrances = BSEPCT_ACT, 0, 1))
    'ini 2015-05-18 validacion frm
    'valida si existe registro en periodo de apetura no chek en crear tablas
    '0=eof 1=si tiene registros
     txtDato(0).Text = "0101" '2015-05-18 validacion frm
    If fCpbCab_Eof() = 0 Then
        chkProceso.Value = 1
    End If
    'fin 2015-05-18 validacion frm
  End If
  
  '2015-05-18 validacion frm cmbPeriodo.ListIndex = IIf(Index = 0, 1, 0)
  cmbPeriodo.ListIndex = IIf(Index = 0, 0, 0)
  chkProceso.Enabled = (Index = 1 And gnFrances = 0)
  frmCuadro(2).Enabled = (Index = 1)
'ini 2015-06-24 control flag mayoriza
chkProceso.Enabled = False 'teo queria que este inhabilitado para siempre
'fin 2015-06-24 control flag mayoriza
End Sub
'ini 2015-05-18 validacion frm
Function fCpbCab_Eof() As Integer
    Dim mEof As Integer
    
    Dim porstCpbCab As ADODB.Recordset
   Set porstCpbCab = New ADODB.Recordset
   Dim s1PdoAno As String
   s1PdoAno = Trim(Str(Val(gsAnoAct) + 1))
   With porstCpbCab
      .ActiveConnection = pocnnMain
      .Source = "SELECT count(*) registro "
      .Source = .Source & "FROM cocpbcab "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & s1PdoAno & "' "
      .Source = .Source & "AND mespvs='01' " '2015-06-22 correccion segun teo solo valida el mes
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      'If .EOF() Then
        mEof = porstCpbCab!registro
      'Else
      '  mEof = 1
      'End If
   End With
   fRstClose porstCpbCab

   fCpbCab_Eof = mEof

End Function
'fin 2015-05-18 validacion frm

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
      ppAyuBus AYUDAT, Index
   End If
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Cancel = ppAyuDet(AYUDAT, Index)
   If Cancel Then Exit Sub
   cmdDatoAyud(0).Enabled = True
   cmdAceptar.Enabled = True
   Exit Sub
End Sub
Private Sub exportar()
Dim nhoja As String
Dim strsql As String
Dim i As Integer
Dim j As Integer
Dim cols As Integer

Dim porstTmp As ADODB.Recordset
Set porstTmp = New ADODB.Recordset
With porstTmp
    .ActiveConnection = pocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
End With
  
nhoja = "Apertura"
strsql = "select '" & Right(cmbPeriodo.Text, 4) & "' as pdoano,'" & txtDato(0).Text & "' as coddro,'000001' as nrocpb,'00' as mespvs,'" & dtpFehCpb.Value & "' as fehopr,'" & txtDato(1).Text & "' as glocpb,'" & txtDato(2).Text & "' as glocpbx,'D' as tpocpb,'xxxx' as ordit,'xxxx' as dest,codtdc,codcta,'' as codcco,codaux,serdoc,nrodoc,date_format(feedoc,'%d/%m/%Y') as feedoc,date_format(fevdoc,'%d/%m/%Y') as fevdoc,date_format(ferdoc,'%d/%m/%Y') as ferdoc,'' as refdoc,gloite,gloitex,if(cdebemn-chabermn >=0,'D','H') as tpoubi,if(codtdc <> '','P','O') as tpodlle,tpomon,tpotcb,'' as imptcb,if(cdebemn-chabermn >=0,cdebemn-chabermn,chabermn-cdebemn) as impmn,if(cdebeme-chaberme >=0,cdebeme-chaberme,chaberme-cdebeme) as impme,'' as nrocorrelativo  from tmpapertura where codemp='" & gsCodEmp & "'"
cols = 30
Dim rsexportar As New Recordset
Dim ApExcel As Variant
Set ApExcel = CreateObject("Excel.application")
ApExcel.Visible = False
ApExcel.Workbooks.Add
ApExcel.Sheets("Hoja1").Name = nhoja

ApExcel.ActiveWindow.Zoom = 75
ApExcel.Cells(1, 1).formula = "Apertura : " & cmbPeriodo.Text
ApExcel.Cells(1, 1).Font.Size = 18
ApExcel.Cells(2, 1).formula = ""
'************************************

rsexportar.Open strsql, cnn, adOpenStatic, adLockOptimistic
On Error GoTo error

rsexportar.MoveFirst
For i = 1 To rsexportar.RecordCount
If i = 1 Then
    For j = 1 To cols
    ApExcel.Cells(i + 3, j).formula = rsexportar.Fields(j - 1).Name
    Next
End If
    For j = 1 To cols
    ApExcel.Cells(i + 4, j).formula = rsexportar(j - 1)
    Next
rsexportar.MoveNext
Next
MsgBox ("Proceso de Exportacion a Excel, terminado")
ApExcel.Visible = True
error:
End Sub




