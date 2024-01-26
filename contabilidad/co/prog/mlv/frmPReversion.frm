VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPReversion 
   Caption         =   "[Entidad]"
   ClientHeight    =   7095
   ClientLeft      =   2190
   ClientTop       =   1485
   ClientWidth     =   7590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   7590
   Visible         =   0   'False
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
      Left            =   6615
      Picture         =   "frmPReversion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6330
      Width           =   720
   End
   Begin VB.Frame frmCuadro 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Filtrar Comprobantes Tipo Diario "
      ForeColor       =   &H00800000&
      Height          =   1020
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton cmdFiltrar 
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
         Left            =   6330
         Picture         =   "frmPReversion.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   285
         Width           =   720
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   0
         Left            =   750
         MaxLength       =   4
         TabIndex        =   2
         Top             =   270
         Width           =   465
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   5640
         Picture         =   "frmPReversion.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Left            =   750
         TabIndex        =   6
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Format          =   144572417
         CurrentDate     =   37974
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
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
         Left            =   150
         TabIndex        =   5
         Top             =   675
         Width           =   540
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diario :"
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
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   495
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
         Left            =   1200
         TabIndex        =   3
         Top             =   270
         Width           =   4455
      End
   End
   Begin VB.Frame frmCuadro 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos para la Reversión "
      ForeColor       =   &H00C00000&
      Height          =   1395
      Index           =   1
      Left            =   45
      TabIndex        =   9
      Top             =   5625
      Width           =   6210
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   1620
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   1245
      End
      Begin VB.CommandButton cmdAceptar 
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
         Left            =   5175
         Picture         =   "frmPReversion.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   690
         Width           =   720
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   5640
         Picture         =   "frmPReversion.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   1
         Left            =   750
         MaxLength       =   4
         TabIndex        =   11
         Top             =   255
         Width           =   465
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes :"
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
         Index           =   4
         Left            =   150
         TabIndex        =   16
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diario :"
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
         Left            =   150
         TabIndex        =   10
         Top             =   285
         Width           =   495
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año :"
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
         Index           =   3
         Left            =   150
         TabIndex        =   14
         Top             =   660
         Width           =   495
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
         Left            =   1200
         TabIndex        =   12
         Top             =   255
         Width           =   4455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgComprobante 
      Bindings        =   "frmPReversion.frx":1522
      Height          =   4365
      Left            =   45
      TabIndex        =   8
      Top             =   1140
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7699
      _Version        =   393216
      BackColor       =   -2147483633
      BackColorFixed  =   12632256
      ForeColorFixed  =   16711680
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      Appearance      =   0
   End
End
Attribute VB_Name = "frmPReversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private usConnStrgSele As String, usConnStrgOrde As String
Private uaTitulos As Variant, uaAncho As Variant, _
       uaFormato As Variant, uaAlineamiento As Variant, _
       uaOrden As Variant
Public uvDato1Posicion As Integer, uvDato1Previo As Variant, uvDato1 As Variant
Public uvDato2Posicion As Integer, uvDato2 As Variant
Public usCriterio As String
Private pnColumnaOrd As Integer
Private nRevSeleccion As Long
'[Propio del formulario.
Private uorstMain As ADODB.Recordset
Private porstCodro As ADODB.Recordset

Private Sub ppComprobante()
  Dim sSentencia As String
  
  sSentencia = "SELECT '0' AS indsel, nrocpb, fehcpb, " & Choose(gsIdioma, "glocpb", "glocpbx") & " AS glocpb, 0 AS nReversa "
  sSentencia = sSentencia & "FROM cocpbcab "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND coddro='" & txtDato(0).Text & "' "
  sSentencia = sSentencia & "AND tpognr NOT IN(" & TPOGNR_APE & ", " & TPOGNR_DCA & ", " & TPOGNR_CIE & ", " & TPOGNR_DST & ") "
  sSentencia = sSentencia & "ORDER BY nrocpb"
  sSentencia = "SELECT '0' AS indsel, cab.nrocpb, cab.fehcpb, " & Choose(gsIdioma, "glocpb", "glocpbx") & " AS glocpb, "
  sSentencia = sSentencia & "IFNULL((SELECT COUNT(*) "
  sSentencia = sSentencia & "FROM cocpbcab rev "
  sSentencia = sSentencia & "WHERE rev.revcodemp=cab.codemp AND rev.revpdoano=cab.pdoano AND rev.revmespvs=cab.mespvs "
  sSentencia = sSentencia & "AND rev.revcoddro=cab.coddro AND rev.revnrocpb=cab.nrocpb "
  sSentencia = sSentencia & "GROUP BY rev.revcodemp, rev.revpdoano, rev.revmespvs, rev.revcoddro, rev.revnrocpb), 0) AS nReversa "
  sSentencia = sSentencia & "FROM cocpbcab cab "
  sSentencia = sSentencia & "WHERE cab.codemp='" & gsCodEmp & "' AND cab.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND cab.mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND cab.coddro='" & txtDato(0).Text & "' "
  sSentencia = sSentencia & "AND cab.tpognr NOT IN(" & TPOGNR_APE & ", " & TPOGNR_DCA & ", " & TPOGNR_CIE & ", " & TPOGNR_DST & ") "
  sSentencia = sSentencia & "ORDER BY nrocpb"

  With uorstMain
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
  End With

End Sub
Private Sub ppInicializaGrilla()
  Dim n_Index As Integer
    
  With mfgComprobante
    .cols = 5
    .FixedCols = 0
    .Rows = .Rows
    .FixedRows = 1
    .GridColor = vbBlack
    .GridColorFixed = vbBlue
    .Gridlines = flexGridFlat
    .GridLinesFixed = flexGridInset
    .GridLineWidth = 1
    .SelectionMode = flexSelectionByRow
    .BackColor = &H80000018
    .BackColorBkg = &H8000000F
    .BackColorFixed = &HC0C0C0
    .BackColorSel = &HE0E0E0
    .ForeColor = vbBlack
    .ForeColorFixed = vbBlack
    .TextStyleFixed = flexTextFlat
    .ForeColorSel = vbBlue
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .Font.Bold = False
  End With
  
  For n_Index = 0 To (mfgComprobante.cols - 1)
    mfgComprobante.Col = n_Index
    mfgComprobante.TextMatrix(0, n_Index) = uaTitulos(n_Index)
    mfgComprobante.ColAlignment(n_Index) = uaAlineamiento(n_Index)
    mfgComprobante.ColWidth(n_Index) = uaAncho(n_Index)
  Next n_Index
  ' Incializo la altura de la fila inicial
  mfgComprobante.RowHeight(1) = 0
  nRevSeleccion = 0
  
End Sub
Private Sub ppRegistrosGrilla()
  Dim n_Index As Integer
  
  mfgComprobante.Redraw = False
  ' Elimino y configuro la grilla
  mfgComprobante.Clear
  ppInicializaGrilla
  n_Index = 1
  uorstMain.Requery
  If uorstMain.RecordCount > 0 Then
    uorstMain.MoveFirst
    n_Index = 2
    Do While Not uorstMain.EOF
      ' Obtengo los importes iniciales
      With mfgComprobante
        .Rows = n_Index
        .TextMatrix(n_Index - 1, 0) = IIf(uorstMain!indsel = INDPREGEN_ACT, "Si", "No")
        .TextMatrix(n_Index - 1, 1) = IIf(IsNull(uorstMain!nrocpb), "", uorstMain!nrocpb)
        .TextMatrix(n_Index - 1, 2) = IIf(IsNull(uorstMain!FehCpb), "", uorstMain!FehCpb)
        .TextMatrix(n_Index - 1, 3) = IIf(IsNull(uorstMain!glocpb), "", uorstMain!glocpb)
        .TextMatrix(n_Index - 1, 4) = FormatNumber(uorstMain!nReversa, 0)
      End With
      ' Incremento las filas
      n_Index = n_Index + 1
      uorstMain.MoveNext
    Loop
  End If
  mfgComprobante.Redraw = True
  
End Sub
Private Sub cmbPeriodo_LostFocus(Index As Integer)
  Dim sFecha As String
  
  If Index = 1 Then
    sFecha = gfUltDia("01/" & IIf(Right(cmbPeriodo(1).Text, 2) = "13", "12", Right(cmbPeriodo(1).Text, 2)) & "/" & cmbPeriodo(0).Text)
    With DTPfecha
      .MaxDate = CDate("31/12/" & gsAnoAct)
      .MinDate = CDate("01/" & Right(cmbPeriodo(1).Text, 2) & "/" & gsAnoAct)
      .Value = .MinDate
    End With
  End If
End Sub
Private Sub cmdAceptar_Click()
  Dim sSentencia As String, sExpresion As String
  Dim sComprobante As String
  Dim nFilaSecuencia As Long

  If txtDato(0).Text = "" Then MsgBox Choose(gsIdioma, "Debe ingresar Diario de Comprobantes", "You Must enter the Journal of Vouchers"), vbCritical: txtDato(0).SetFocus: Exit Sub
  If nRevSeleccion <= 0 Then MsgBox Choose(gsIdioma, "Seleccione comprobantes a reversar", "Select vouchers to reverse"), vbCritical: txtDato(0).SetFocus: Exit Sub
 ' If mfgComprobante.Rows <= 0 Then MsgBox Choose(gsIdioma, "Seleccione comprobantes a reversar", "Select vouchers to reverse"), vbCritical: txtDato(0).SetFocus: Exit Sub
  If MsgBox(Choose(gsIdioma, "Estás Seguro de Procesar Reversión", " Are you sure of Process Reverse"), vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  sExpresion = gfNumComprobante(gsAnoAct, Right(cmbPeriodo(1).Text, 2), txtDato(1).Text)
  For nFilaSecuencia = 1 To mfgComprobante.Rows - 1
    ' Registros seleccionado
    If mfgComprobante.TextMatrix(nFilaSecuencia, 0) = "Si" Then
      sComprobante = mfgComprobante.TextMatrix(nFilaSecuencia, 1)
      ' cabecera de comprobante
      sSentencia = "INSERT INTO cocpbcab (codemp, pdoano, mespvs, coddro, nrocpb, fehcpb, glocpb, glocpbx, tpognr, indncu, indanu, usrcre, fyhcre, "
      sSentencia = sSentencia & "indSelInv, revcodemp, revpdoano, revmespvs, revcoddro, revnrocpb, Rectificar, Adicionar) "
      sSentencia = sSentencia & "SELECT codemp, pdoano, "
      sSentencia = sSentencia & "'" & Right(cmbPeriodo(1).Text, 2) & "' AS mespvs, "
      sSentencia = sSentencia & "'" & txtDato(1).Text & "' AS coddro, "
      sSentencia = sSentencia & "'" & sExpresion & "' AS nrocpb, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(DTPfecha.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & ") AS fehcpb, "
      sSentencia = sSentencia & "CONCAT('" & Choose(gsIdioma, "EXTORNO ", "BACK ") & "', glocpb) AS glocpb, "
      sSentencia = sSentencia & "glocpbx, "
      sSentencia = sSentencia & "'" & TPOGNR_DRO & "' AS tpognr, indncu, indanu, "
      sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ") AS fyhcre, "
      sSentencia = sSentencia & "indSelInv, codemp AS revcodemp, pdoano AS revpdoano, mespvs AS revmespvs, "
      sSentencia = sSentencia & "coddro AS revcoddro, nrocpb AS revnrocpb, Null AS Rectificar, Null AS Adicionar "
      sSentencia = sSentencia & "FROM cocpbcab "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
      sSentencia = sSentencia & "AND coddro='" & txtDato(0).Text & "' "
      sSentencia = sSentencia & "AND nrocpb='" & sComprobante & "'"
      frmTCpbGrd.uocnnMain.Execute sSentencia
      
      ' detalle de comprobante
      sSentencia = "INSERT INTO cocpbdet (codemp, pdoano, mespvs, coddro, nrocpb, nroite, blqite, codtdc, fehope, codcta, codcco, codaux, "
      sSentencia = sSentencia & "serdoc, nrodoc, feedoc, fevdoc, ferdoc, refdoc, pdocpr, gloite, gloitex, tpoctb, tpopvs, tpomon, tpotcb, imptcb, impmn,"
      sSentencia = sSentencia & "impme, tpognr, indfjo_det, indgnr_rp, tpodoc, codaux_psn, codcon, codmon, usrcre, fyhcre, Rectificar, Adicionar, flagGastoDeducible) "
      sSentencia = sSentencia & "SELECT codemp, pdoano, "
      sSentencia = sSentencia & "'" & Right(cmbPeriodo(1).Text, 2) & "' AS mespvs, "
      sSentencia = sSentencia & "'" & txtDato(1).Text & "' AS coddro, "
      sSentencia = sSentencia & "'" & sExpresion & "' AS nrocpb, "
      sSentencia = sSentencia & "nroite, blqite, codtdc, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(DTPfecha.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & ") AS fehope, "
      sSentencia = sSentencia & "codcta, codcco, codaux, serdoc, nrodoc, feedoc, fevdoc, ferdoc, refdoc, pdocpr, "
      sSentencia = sSentencia & "CONCAT('" & Choose(gsIdioma, "EXTORNO ", "BACK ") & "', gloite) AS gloite, gloitex, "
      sSentencia = sSentencia & "(CASE WHEN tpoctb='" & TPOCTB_DEB & "' THEN '" & TPOCTB_HAB & "' ELSE '" & TPOCTB_DEB & "' END) AS tpoctb, "
      sSentencia = sSentencia & "'" & TPOPVS_OTR & "' AS tpopvs, tpomon, tpotcb, imptcb, impmn, impme, '" & TPOGNR_DRO & "' AS tpognr, "
      sSentencia = sSentencia & "indfjo_det, indgnr_rp, tpodoc, codaux_psn, codcon, codmon, "
      sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ") AS fyhcre, "
      sSentencia = sSentencia & "Rectificar, Adicionar, flagGastoDeducible "
      sSentencia = sSentencia & "FROM cocpbdet "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
      sSentencia = sSentencia & "AND coddro='" & txtDato(0).Text & "' "
      sSentencia = sSentencia & "AND nrocpb='" & sComprobante & "' "
      sSentencia = sSentencia & "AND tpognr<>" & TPOGNR_DST
      frmTCpbGrd.uocnnMain.Execute sSentencia
    
      ' detalle de comprobante - flujo
      sSentencia = "INSERT INTO cocpbdetfjo (codemp, pdoano, mespvs, coddro, nrocpb, nroite, nroord, codfjo, codcta, tpoctb, impmn, impme, usrcre, fyhcre) "
      sSentencia = sSentencia & "SELECT codemp, pdoano, "
      sSentencia = sSentencia & "'" & Right(cmbPeriodo(1).Text, 2) & "' AS mespvs, "
      sSentencia = sSentencia & "'" & txtDato(1).Text & "' AS coddro, "
      sSentencia = sSentencia & "'" & sExpresion & "' AS nrocpb, "
      sSentencia = sSentencia & "nroite, nroord, codfjo, codcta, tpoctb, impmn, impme, "
      sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ") AS fyhcre "
      sSentencia = sSentencia & "FROM cocpbdetfjo "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
      sSentencia = sSentencia & "AND coddro='" & txtDato(0).Text & "' "
      sSentencia = sSentencia & "AND nrocpb='" & sComprobante & "'"
      frmTCpbGrd.uocnnMain.Execute sSentencia
    
      ' detalle de comprobante - retencion/percepcion
      sSentencia = "INSERT INTO cocpbdetrp (codemp, pdoano, mespvs, coddro, nrocpb, nroite, codtdc_rtcpcp, serdoc_rtcpcp, nrodoc_rtcpcp, feedoc_rtcpcp, impmn_rtcpcp, "
      sSentencia = sSentencia & "impme_rtcpcp, codaux, codcta, codtdc, serdoc, nrodoc, impmn, impme, indrtcpcp, usrcre, fyhcre) "
      sSentencia = sSentencia & "SELECT codemp, pdoano, "
      sSentencia = sSentencia & "'" & Right(cmbPeriodo(1).Text, 2) & "' AS mespvs, "
      sSentencia = sSentencia & "'" & txtDato(1).Text & "' AS coddro, "
      sSentencia = sSentencia & "'" & sExpresion & "' AS nrocpb, "
      sSentencia = sSentencia & "nroite, codtdc_rtcpcp, serdoc_rtcpcp, nrodoc_rtcpcp, feedoc_rtcpcp, impmn_rtcpcp, "
      sSentencia = sSentencia & "impme_rtcpcp, codaux, codcta, codtdc, serdoc, nrodoc, impmn, impme, indrtcpcp, "
      sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ") AS fyhcre "
      sSentencia = sSentencia & "FROM cocpbdetrp "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
      sSentencia = sSentencia & "AND coddro='" & txtDato(0).Text & "' "
      sSentencia = sSentencia & "AND nrocpb='" & sComprobante & "'"
      frmTCpbGrd.uocnnMain.Execute sSentencia
      sExpresion = gfCeros(sExpresion, 6, 1, "0")
    End If
  Next nFilaSecuencia
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  MsgBox TEXT_8008, vbInformation

End Sub
Private Sub cmdDatoAyud_Click(Index As Integer)
  txtDato(Index).SetFocus
  ppAyuBus Index
End Sub
Private Sub cmdFiltrar_Click()
  ppComprobante
  ppRegistrosGrilla
End Sub
Private Sub Form_Load()
  Dim nContador As Long
  
  '[Recordsets.                         'Cambiar.
  Set uorstMain = New ADODB.Recordset
  Set porstCodro = New ADODB.Recordset

  With porstCodro
    .ActiveConnection = frmTCpbGrd.uocnnMain
    .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
    .Source = .Source & "FROM CODro "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4 "
    .Source = .Source & "ORDER BY CodDro"
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
 ']

 '[Parámetros.                         'Cambiar.
  With txtDato
    For nContador = 0 To 1
      .Item(nContador).DataField = "CodDro"
      .Item(nContador).MaxLength = porstCodro.Fields(.Item(nContador).DataField).DefinedSize
    Next
  End With
 ']
  usConnStrgSele = "SELECT '0' AS indsel, nrocpb, fehcpb, " & Choose(gsIdioma, "glocpb", "glocpbx") & " AS glocpb, 0 AS nReversa "
  usConnStrgSele = usConnStrgSele & "FROM cocpbcab "
  usConnStrgSele = usConnStrgSele & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
  usConnStrgSele = usConnStrgSele & "AND mespvs='" & gsMesAct & "' "
  usConnStrgSele = usConnStrgSele & "AND coddro='" & txtDato(0).Text & "' "
  usConnStrgSele = usConnStrgSele & "ORDER BY nrocpb"
  With uorstMain
    .ActiveConnection = frmTCpbGrd.uocnnMain
    .Source = usConnStrgSele
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diario :", "Fecha :", "Diario :", "Año :", "Mes :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journal:", "Date :", "Journal :", "Year :", "Mounth :")
  Next nElemento
  frmCuadro(0).Caption = Choose(gsIdioma, " Filtrar Comprobantes Tipo Diario ", " Filter Vouchers Type Daily ")
  frmCuadro(1).Caption = Choose(gsIdioma, " Datos para la Reversión ", " Data for the Reverse ")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
 
  If gsIdioma = NvlUsr_Sup Then
    uaTitulos = Array("Revertir", "Comprob.", "Fecha", "Glosa", "Cant. Rev.")
  Else
    uaTitulos = Array("Reverse", "Voucher", "Date", "Gloss", "Quan. Rev.")
  End If
  uaAlineamiento = Array(flexAlignCenterCenter, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignRightCenter)
  uaAncho = Array(750, 800, 1000, 3850, 740)
 
  With DTPfecha
    .MinDate = CDate("01/" & gsMesAct & "/" & gsAnoAct)
    .MaxDate = gfUltDia(.MinDate)
    .Value = .MinDate
  End With
  ' Configuro los controles de año y mes
  cmbPeriodo(0).AddItem gsAnoAct
  cmbPeriodo(0).ListIndex = 0
  
  For nContador = Val(gsMesAct) To 13
    If gsIdioma = NvlUsr_Sup Then
      cmbPeriodo(1).AddItem Choose(nContador + 1, "Apertura" & Space(50) & "00", "Enero" & Space(50) & "01", "Febrero" & Space(50) & "02", "Marzo" & Space(50) & "03", "Abril" & Space(50) & "04", "Mayo" & Space(50) & "05", "Junio" & Space(50) & "06", "Julio" & Space(50) & "07", "Agosto" & Space(50) & "08", "Setiembre" & Space(50) & "09", "Octubre" & Space(50) & "10", "Noviembre" & Space(50) & "11", "Diciembre" & Space(50) & "12", "Cierre" & Space(50) & "13")
    Else
      cmbPeriodo(1).AddItem Choose(nContador + 1, "Opening" & Space(50) & "00", "January" & Space(50) & "01", "February" & Space(50) & "02", "March" & Space(50) & "03", "April" & Space(50) & "04", "May" & Space(50) & "05", "June" & Space(50) & "06", "July" & Space(50) & "07", "August" & Space(50) & "08", "September" & Space(50) & "09", "October" & Space(50) & "10", "November" & Space(50) & "11", "December" & Space(50) & "12", "Closing" & Space(50) & "13")
    End If
  Next nContador
  cmbPeriodo(1).ListIndex = 0
  
  ' Inicializo grilla
  ppInicializaGrilla
  cmdfiltrar.ToolTipText = Choose(gsIdioma, "Filtrar Comprobantes ", "Filter Vouchers")
  cmdAceptar.ToolTipText = Choose(gsIdioma, "Procesar Reversión", "Process Reverse")
  txtDato(1).Enabled = False
  cmdDatoAyud(1).Enabled = False
  nRevSeleccion = 0
  cmdAceptar.Enabled = (nRevSeleccion <> 0)
  
End Sub
Private Sub Form_Activate()
  'fraBuscar.Caption = TEXT_BUSCA & mfgComprobante.TextMatrix(0, 1)
End Sub
Private Sub Form_Resize()
   On Error Resume Next
  
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   cmdAceptar.Left = frmCuadro(1).Width - 1030
   cmdSalir.Left = Me.Width - 1220

End Sub
Private Sub Form_Unload(Cancel As Integer)
   porstCodro.Close
   uorstMain.Close
   Set porstCodro = Nothing
   Set uorstMain = Nothing
End Sub
Private Sub cmdSalir_Click()
  Unload Me
End Sub
Private Sub mfgComprobante_DblClick()
  ' Verificar registros
  If mfgComprobante.Rows = 1 Or mfgComprobante.row = 1 Then Exit Sub
  If mfgComprobante.Col = 0 Then
    ' veces de reversión
    If (mfgComprobante.TextMatrix(mfgComprobante.row, mfgComprobante.Col) = "No" And Int(mfgComprobante.TextMatrix(mfgComprobante.row, 4)) >= 1) Then
      If MsgBox(Choose(gsIdioma, "Comprobante '" & mfgComprobante.TextMatrix(mfgComprobante.row, 1) & "' fue revertido [" & Int(mfgComprobante.TextMatrix(mfgComprobante.row, 4)) & "] veces.", "Voucher '" & mfgComprobante.TextMatrix(mfgComprobante.row, 1) & "' was reversed [" & Int(mfgComprobante.TextMatrix(mfgComprobante.row, 4)) & "] times.") & Chr(13) & Choose(gsIdioma, " Desea revertir comprobante?", "Do you want to reverse voucher?"), vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption) = vbNo Then Exit Sub
    End If
    mfgComprobante.TextMatrix(mfgComprobante.row, mfgComprobante.Col) = IIf(mfgComprobante.TextMatrix(mfgComprobante.row, mfgComprobante.Col) = "Si", "No", "Si")
    nRevSeleccion = nRevSeleccion + IIf(mfgComprobante.TextMatrix(mfgComprobante.row, mfgComprobante.Col) = "Si", 1, -1)
    cmdAceptar.Enabled = (nRevSeleccion <> 0)
  End If
End Sub
Private Sub mfgComprobante_KeyPress(KeyAscii As Integer)
  ' Verificar registros
  If mfgComprobante.Rows = 2 Or mfgComprobante.row = 1 Then Exit Sub
End Sub
Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property
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
  Select Case Index    'Busca el dato en su tabla principal.
    Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
  End Select
  
  If Len(Trim(txtDato(0).Text)) Then
    ppComprobante
    ppRegistrosGrilla
  End If
  
End Sub
Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
    Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + frmCuadro(tnIndex).Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      txtDato(1).Text = txtDato(tnIndex).Text
      lblDatoDeta(1).Caption = lblDatoDeta(tnIndex).Caption
  End Select
End Sub
Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
    Case 0, 1
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        txtDato(1).Text = txtDato(tnIndex).Text
        lblDatoDeta(1).Caption = lblDatoDeta(tnIndex).Caption
      Exit Function
      End If
      With porstCodro
        .MoveFirst
        .Find "CodDro='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
        End If
      End With
      txtDato(1).Text = txtDato(tnIndex).Text
      lblDatoDeta(1).Caption = lblDatoDeta(tnIndex).Caption
  End Select
End Function

