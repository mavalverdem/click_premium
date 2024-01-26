VERSION 5.00
Begin VB.Form frmRMayAuxCCo 
   Caption         =   "[título]"
   ClientHeight    =   5565
   ClientLeft      =   2460
   ClientTop       =   1875
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Exporta Excel"
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
      Index           =   2
      Left            =   3480
      Picture         =   "frmRMayAuxCCo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5040
      Width           =   1125
   End
   Begin VB.CommandButton cmdexcel 
      Caption         =   "Excel"
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
      Left            =   5760
      TabIndex        =   39
      Top             =   5030
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CheckBox chkReferencia 
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   1125
      TabIndex        =   11
      Top             =   3120
      Width           =   180
   End
   Begin VB.Frame fraReferencia 
      Caption         =   " Filtro "
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   90
      TabIndex        =   10
      Top             =   3120
      Width           =   4215
      Begin VB.OptionButton optReferencia 
         Caption         =   "Pedido"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optReferencia 
         Caption         =   "Referencia"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.TextBox txtReferencia 
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
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   14
         Top             =   225
         Width           =   1695
      End
   End
   Begin VB.CheckBox chkRango 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1440
      TabIndex        =   16
      Top             =   3840
      Width           =   180
   End
   Begin VB.Frame fraRngPeriodo 
      Caption         =   " Rango Periodos "
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   90
      TabIndex        =   15
      Top             =   3840
      Width           =   4215
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   18
         Text            =   "Año Inicio"
         Top             =   300
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   855
         TabIndex        =   21
         Text            =   "Año Final"
         Top             =   645
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   19
         Text            =   "Mes Inicio"
         Top             =   300
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   22
         Text            =   "Mes Final"
         Top             =   645
         Width           =   1710
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   20
         Top             =   690
         Width           =   765
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4680
      TabIndex        =   23
      Top             =   3960
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   24
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H00800000&
      Height          =   2985
      Left            =   0
      TabIndex        =   27
      Top             =   60
      Width           =   6990
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   4
         Left            =   5520
         Picture         =   "frmRMayAuxCCo.frx":03B2
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2400
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
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   1575
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
         Left            =   120
         TabIndex        =   7
         Top             =   1800
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
         TabIndex        =   6
         Top             =   1440
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   4560
         Picture         =   "frmRMayAuxCCo.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   4560
         Picture         =   "frmRMayAuxCCo.frx":0706
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRMayAuxCCo.frx":08B0
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Left            =   135
         TabIndex        =   4
         Top             =   495
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
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   945
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRMayAuxCCo.frx":0A5A
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   855
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
         Index           =   4
         Left            =   1680
         TabIndex        =   42
         Top             =   2400
         Width           =   3720
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Auxiliar"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Width           =   615
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
         TabIndex        =   38
         Top             =   1800
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
         TabIndex        =   37
         Top             =   1440
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
         Index           =   0
         Left            =   1065
         TabIndex        =   34
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
         Left            =   1065
         TabIndex        =   33
         Top             =   855
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   30
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centros de Costo"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5025
      Width           =   6990
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
         Left            =   4605
         Picture         =   "frmRMayAuxCCo.frx":0C04
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
         Picture         =   "frmRMayAuxCCo.frx":0D4E
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
         Picture         =   "frmRMayAuxCCo.frx":1280
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
      Index           =   2
      Left            =   4800
      TabIndex        =   29
      Top             =   3240
      Width           =   675
   End
End
Attribute VB_Name = "frmRMayAuxCCo"
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
Private porstTGAux As ADODB.Recordset
']
Private Sub chkRango_Click()
  fraRngPeriodo.Enabled = (chkRango.Value = vbChecked)
End Sub
Private Sub chkReferencia_Click()
  fraReferencia.Enabled = (chkReferencia.Value = vbChecked)
End Sub

Private Sub cmdexcel_Click()
'2016-07-08 se inhabilita esta opcion hasta nuevo aviso.
'por que sale un error. TCS
  Dim dnContador As Integer, s_Sql As String
  Dim s_Sentencia As String, s_Moneda As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim l_CreateTB As Boolean
  Dim Index As Integer
       
  Index = 0
       
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End fiscal year must be equal or more than Opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
  ppHabilitacion False
    
  s_SaldoDeb = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
  s_SaldoHab = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  s_Mes = Format(IIf(chkRango.Value = vbChecked, cmbPeriodo(2).ListIndex, gsMesAct), "00")
  If s_Mes > "00" Then
    For dnContador = 0 To (Val(s_Mes) - 1)
      s_SaldoDeb = s_SaldoDeb & "AcuD" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = (Val(s_Mes) - 1), "", "+")
      s_SaldoHab = s_SaldoHab & "AcuH" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = (Val(s_Mes) - 1), "", "+")
    Next dnContador
    s_SaldoDeb = s_SaldoDeb & ", 0)"
    s_SaldoHab = s_SaldoHab & ", 0)"
  Else
    s_SaldoDeb = "ROUND(0": s_SaldoHab = "ROUND(0"
  End If
  s_SaldoDeb = s_SaldoDeb & ", 2) AS cSalDeb, "
  s_SaldoHab = s_SaldoHab & ", 2) AS cSalHab "
    
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRMayAuxCCo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 15)='#trptRMayAuxCCo') DROP TABLE #trptRMayAuxCCo")
  ' Inserto los registros
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    s_SaldoDeb = IIf(s_Ano = s_AnoIni, s_SaldoDeb, "ROUND(0, 2) AS cSalDeb, ")
    s_SaldoHab = IIf(s_Ano = s_AnoIni, s_SaldoHab, "ROUND(0, 2) AS cSalHab ")
    s_Sentencia = "SELECT a.pdoano AS cAno, a.MesPvs, a.CodDro, a.NroCpb, a.NroIte,  a.FehOpe, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & "AS cDocume, "
    s_Sentencia = s_Sentencia & "a.RefDoc, a.PdoCpr, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, a.CodCCo, " & Choose(gsIdioma, "b.DetCCo", "b.DetCCox") & " AS DetCCo, a.CodCta, " & Choose(gsIdioma, "d.DetCta", "d.DetCtax") & " AS DetCta, "
    s_Sentencia = s_Sentencia & "a.CodAux, f.RazAux, a.ImpTcb, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & s_Moneda & " ELSE 0 END) AS cDebe, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & s_Moneda & " ELSE 0 END) AS cHaber, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & s_SaldoHab
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trptRMayAuxCCo ", "")
    s_Sentencia = s_Sentencia & "FROM (((((COCpbDet a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCCo=b.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCoAcu e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.CodCta=e.CodCta AND a.CodCCo=e.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCta=d.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux f ON a.codemp=f.codemp AND a.CodAux=f.CodAux) "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    If chkRango.Value = vbChecked Then
      s_Mes = Format(IIf(s_Ano = s_AnoIni, cmbPeriodo(2).ListIndex, "0"), "00")
      s_Sentencia = s_Sentencia & "AND a.Mespvs >='" & s_Mes & "' "
      If (s_Ano = s_AnoFin) Then
        s_Mes = Format(cmbPeriodo(3).ListIndex, "00")
        s_Sentencia = s_Sentencia & "AND a.Mespvs <='" & s_Mes & "' "
      End If
      s_Sentencia = s_Sentencia & "AND a.CodCCo BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    Else
      s_Sentencia = s_Sentencia & "AND a.Mespvs='" & gsMesAct & "' "
    End If
    If chkReferencia.Value = vbChecked Then
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & IIf(optReferencia(0).Value, "a.RefDoc", "a.pdocpr") & ", '') ='" & Trim$(txtReferencia) & "' "
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCCo, '')<>'' "
    s_Sentencia = s_Sentencia & "AND a.tpognr<>'" & TPOGNR_CIE & "' "
    s_Sentencia = s_Sentencia & "ORDER BY a.MesPvs, a.CodCCo, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
    ' Executo la sentencia
    If Not l_CreateTB Then
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptRMayAuxCCo ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trptRMayAuxCCo "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
    
  ' Recordset del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT * "
    .Source = .Source & "FROM " & ps_Prefijo & "trptRMayAuxCCo "
    .Source = .Source & "ORDER BY cAno, MesPvs, CodCCo, CodCta, CodDro, NroCpb, NroIte"
    .Open
  End With
     
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRMayAuxCCoDet.rpt"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      '.Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Destination = crptToFile
      .PrintFileType = crptExcel50
      .Action = 1
    End With
  
  ' Elimino el archivo tempooral
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRMayAuxCCo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 15)='#trptRMayAuxCCo') DROP TABLE #trptRMayAuxCCo")
  
  ppHabilitacion True

End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstCoCCo = New ADODB.Recordset
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
   With porstCoCCo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS DetCCo "
      .Source = .Source & "FROM CoCCo "
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
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Centro de Costo :", "Moneda :", "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Cost Center :", "Currency :", "Beginning :", "End :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraReferencia.Caption = Choose(gsIdioma, " Filtro ", " Filter ")
  optReferencia(0).Caption = Choose(gsIdioma, "Referencia", "Reference")
  optReferencia(1).Caption = Choose(gsIdioma, "Pedido", "Order")
  fraRngPeriodo.Caption = Choose(gsIdioma, "Rango Periodos", "Range of Periods")
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
   With porstCoCCo
      .MoveLast
      txtDato(3).Text = !codcco
      .MoveFirst
      txtDato(2).Text = !codcco
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
   ' Configuro los controles de año y mes
    For dnContador = (Val(gsAnoAct) - 9) To Val(gsAnoAct)
      cmbPeriodo(0).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador
      cmbPeriodo(1).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador
    Next dnContador
    cmbPeriodo(0).ListIndex = 9
    cmbPeriodo(1).ListIndex = 9
    
    For dnContador = 0 To 13
      If gsIdioma = NvlUsr_Sup Then
        cmbPeriodo(2).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
        cmbPeriodo(3).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
      Else
        cmbPeriodo(2).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
        cmbPeriodo(3).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
      End If
    Next dnContador
    cmbPeriodo(2).ListIndex = Val(gsMesAct)
    cmbPeriodo(3).ListIndex = Val(gsMesAct)
    fraRngPeriodo.Enabled = False
    fraReferencia.Enabled = False
  
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
   porstCoCCo.Close
   porstCOCta.Close
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
  Dim dnContador As Integer, s_Sql As String
  Dim s_Sentencia As String, s_Moneda As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim l_CreateTB As Boolean
       
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End fiscal year must be equal or more than Opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
  ppHabilitacion False
    
  s_SaldoDeb = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
  s_SaldoHab = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  s_Mes = Format(IIf(chkRango.Value = vbChecked, cmbPeriodo(2).ListIndex, gsMesAct), "00")
  If s_Mes > "00" Then
    For dnContador = 0 To (Val(s_Mes) - 1)
      s_SaldoDeb = s_SaldoDeb & "AcuD" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = (Val(s_Mes) - 1), "", "+")
      s_SaldoHab = s_SaldoHab & "AcuH" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = (Val(s_Mes) - 1), "", "+")
    Next dnContador
    s_SaldoDeb = s_SaldoDeb & ", 0)"
    s_SaldoHab = s_SaldoHab & ", 0)"
  Else
    s_SaldoDeb = "ROUND(0": s_SaldoHab = "ROUND(0"
  End If
  s_SaldoDeb = s_SaldoDeb & ", 2) AS cSalDeb, "
  s_SaldoHab = s_SaldoHab & ", 2) AS cSalHab "
    
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRMayAuxCCo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 15)='#trptRMayAuxCCo') DROP TABLE #trptRMayAuxCCo")
  ' Inserto los registros
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    s_SaldoDeb = IIf(s_Ano = s_AnoIni, s_SaldoDeb, "ROUND(0, 2) AS cSalDeb, ")
    s_SaldoHab = IIf(s_Ano = s_AnoIni, s_SaldoHab, "ROUND(0, 2) AS cSalHab ")
    s_Sentencia = "SELECT a.pdoano AS cAno, a.MesPvs, a.CodDro, a.NroCpb, a.NroIte,  a.FehOpe, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & "AS cDocume, "
    s_Sentencia = s_Sentencia & "a.RefDoc, a.PdoCpr, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, a.CodCCo, " & Choose(gsIdioma, "b.DetCCo", "b.DetCCox") & " AS DetCCo, a.CodCta, " & Choose(gsIdioma, "d.DetCta", "d.DetCtax") & " AS DetCta, "
    s_Sentencia = s_Sentencia & "a.CodAux, f.RazAux, a.ImpTcb, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & s_Moneda & " ELSE 0 END) AS cDebe, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & s_Moneda & " ELSE 0 END) AS cHaber, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & s_SaldoHab
 'ini 2016-07-08 exporte excel
    s_Sentencia = s_Sentencia & ",a.GloItex "
 'fin 2016-07-08 exporte excel
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trptRMayAuxCCo ", "")
    s_Sentencia = s_Sentencia & "FROM (((((COCpbDet a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCCo=b.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCoAcu e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.CodCta=e.CodCta AND a.CodCCo=e.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCta=d.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux f ON a.codemp=f.codemp AND a.CodAux=f.CodAux) "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCCo BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    
    If txtDato(4).Text <> "" Then
        s_Sentencia = s_Sentencia & "AND a.CodAux='" & txtDato(4).Text & "' "
    End If
    
    If chkRango.Value = vbChecked Then
        s_Mes = Format(IIf(s_Ano = s_AnoIni, cmbPeriodo(2).ListIndex, "0"), "00")
      s_Sentencia = s_Sentencia & "AND a.Mespvs >='" & s_Mes & "' "
      If (s_Ano = s_AnoFin) Then
        s_Mes = Format(cmbPeriodo(3).ListIndex, "00")
        s_Sentencia = s_Sentencia & "AND a.Mespvs <='" & s_Mes & "' "
      End If
    Else
      s_Sentencia = s_Sentencia & "AND a.Mespvs='" & gsMesAct & "' "
    End If
    
    If chkReferencia.Value = vbChecked Then
        s_Sentencia = s_Sentencia & "AND LEFT(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & IIf(optReferencia(0).Value, "a.RefDoc", "a.pdocpr") & ", ''), " & Len(Trim(txtReferencia)) & ")='" & Trim(txtReferencia) & "' "
    End If
        s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCCo, '')<>'' "
        s_Sentencia = s_Sentencia & "AND a.tpognr<>'" & TPOGNR_CIE & "' "
        s_Sentencia = s_Sentencia & "ORDER BY a.MesPvs, a.CodCCo, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
        ' Executo la sentencia
    If Not l_CreateTB Then
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptRMayAuxCCo ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trptRMayAuxCCo "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
    
  ' Recordset del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
 'ini 2016-07-08 exporte excel
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
 'fin 2016-07-08 exporte excel     'sala un error al procesar con exporta excel, para prinr ok
    
 'ini 2016-07-08 exporte excel
'    .Source = "SELECT * "
'    .Source = .Source & "FROM " & ps_Prefijo & "trptRMayAuxCCo "
'    .Source = .Source & "ORDER BY cAno, MesPvs, CodCCo, CodCta, CodDro, NroCpb, NroIte"
    If Index = 2 Then
    .Source = "SELECT "
    .Source = .Source & "cAno,   MesPvs, CodCCo, DetCCo, CodCta, "
    .Source = .Source & "DetCta, CodDro, NroCpb, NroIte, FehOpe, "
    .Source = .Source & "cDocume,    CodAux, RazAux, RefDoc, PdoCpr, "
    .Source = .Source & "ImpTcb, GloIte, GloItex,    cDebe,  cHaber, "
    .Source = .Source & "cSalDeb,    cSalHab "
    .Source = .Source & "FROM " & ps_Prefijo & "trptRMayAuxCCo "
    .Source = .Source & "ORDER BY cAno, MesPvs, CodCCo, CodCta, CodDro, NroCpb, NroIte"
    Else
    .Source = "SELECT * "
    .Source = .Source & "FROM " & ps_Prefijo & "trptRMayAuxCCo "
    .Source = .Source & "ORDER BY cAno, MesPvs, CodCCo, CodCta, CodDro, NroCpb, NroIte"
    End If
 'fin 2016-07-08 exporte excel
    .Open
  End With
  
 
'ini 2016-07-08 exporte excel
'index=2 exporta excel
If Index = 2 Then
   pExporta 1, porstMRp
Else
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRMayAuxCCo.rpt"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRMayAuxCCo.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      '         .Parameters("pPeriodoAdc") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
End If
'fin 2016-07-08 exporte excel

  ' Elimino el archivo tempooral
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRMayAuxCCo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 15)='#trptRMayAuxCCo') DROP TABLE #trptRMayAuxCCo")
  
  ppHabilitacion True
End Sub
Private Sub pExporta(TpoRpt As Integer, porstTmp As Recordset)
'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err


'    Dim xArchPeriodo As String
'    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
 
    '*Set oExcel = New Excel.Application
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True

    Set oWBook = oExcel.Workbooks.Add
     Set oSheet = oWBook.Worksheets(1)

    With oSheet

        oSheet.Select
        
        .Cells(1, 1).Value = "Mayor Aux. por C.Costo"
        
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )"
        nRowI = nRowI + 2
        Dim x1 As Integer
        .Cells(nRowI, 1).Value = "Periodo"
        .Cells(nRowI, 2).Value = "Mes"
        .Cells(nRowI, 3).Value = "C.Costos"
        .Cells(nRowI, 4).Value = "Detalle C.C."
        .Cells(nRowI, 5).Value = "Cuenta"
        .Cells(nRowI, 6).Value = "Detalle Cta."
        .Cells(nRowI, 7).Value = "Diario"
        .Cells(nRowI, 8).Value = "Comprobante"
        .Cells(nRowI, 9).Value = "Item"
        .Cells(nRowI, 10).Value = "Fecha"
        .Cells(nRowI, 11).Value = "Documento"
        .Cells(nRowI, 11 + 1).Value = "Auxiliar"
        .Cells(nRowI, 12 + 1).Value = "R.Social"
        .Cells(nRowI, 13 + 1).Value = "Referencia"
        .Cells(nRowI, 14 + 1).Value = "Pedido"
        .Cells(nRowI, 15 + 1).Value = "T/C"
        .Cells(nRowI, 16 + 1).Value = "Glosa"
        .Cells(nRowI, 17 + 1).Value = "Glosa 2"
        .Cells(nRowI, 18 + 1).Value = "Debe"
        .Cells(nRowI, 19 + 1).Value = "Haber"
        .Cells(nRowI, 20 + 1).Value = "Sdo Ini Debe"
        .Cells(nRowI, 21 + 1).Value = "Sdo Ini Haber"
    
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        .Columns.AutoFit ' ajusta el ancho de las columnas
        'Sheets(oSheet).Select
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'solo sale error en esta        Columns("M:M").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("O:O").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("P:P").Select
'        Selection.NumberFormat = "#,##0.00"
        
        
    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing
'   porstTmp.Close
'   Set porstTmp = Nothing
  Exit Sub
Err:
   MsgBox (TEXT_6001)
'   porstTmp.Close
'   Set porstTmp = Nothing
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
   Case 0, 1, 2, 3, 4                        'Cambiar (añadir índices).
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
     Case 4
       modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
         End If
      End With
    Case 4
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
            lblDatoDeta(tnIndex).Caption = " " & !razAux
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
