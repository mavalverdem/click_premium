VERSION 5.00
Begin VB.Form frmLCpb 
   Caption         =   "[título]"
   ClientHeight    =   3345
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrnDirec 
      Caption         =   "Prn1"
      Height          =   495
      Left            =   4440
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir2 
      Caption         =   "I&mprimir Formato"
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
      Left            =   2040
      Picture         =   "frmLCpb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdImprimir2 
      Caption         =   "Preliminar Asiento"
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
      Picture         =   "frmLCpb.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   1485
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   3240
      TabIndex        =   15
      Top             =   1320
      Width           =   2220
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   16
         Top             =   315
         Width           =   915
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   9
      Top             =   45
      Width           =   5450
      Begin VB.TextBox TxtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   675
         MaxLength       =   6
         TabIndex        =   7
         Top             =   825
         Width           =   735
      End
      Begin VB.TextBox TxtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   675
         MaxLength       =   6
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox TxtDato 
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
         Height          =   330
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   6
         Top             =   825
         Width           =   570
      End
      Begin VB.TextBox TxtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
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
         MaxLength       =   4
         TabIndex        =   4
         Top             =   480
         Width           =   570
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   5055
         Picture         =   "frmLCpb.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   495
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   5055
         Picture         =   "frmLCpb.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Comprobantes"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   18
         Top             =   255
         Width           =   1095
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
         Left            =   1380
         TabIndex        =   14
         Top             =   825
         Width           =   3675
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
         Left            =   1380
         TabIndex        =   13
         Top             =   480
         Width           =   3675
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diarios"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1140
      ScaleWidth      =   5445
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2205
      Width           =   5445
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
         Left            =   2760
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
         Left            =   4260
         Picture         =   "frmLCpb.frx":0988
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
         Picture         =   "frmLCpb.frx":0AD2
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1485
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
         Left            =   1560
         Picture         =   "frmLCpb.frx":1004
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmLCpb"
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
Private porstMRp2 As ADODB.Recordset

'[Propio del formulario.
Private porstCodro As ADODB.Recordset
']

Private Sub cmdImprimir2_Click(Index As Integer)
  Dim dsFecha As String, dsGirado As String, dsGirado2 As String, dsImporteNumeros As String, dsImporteLetras As String
  Dim dbHayAux As Boolean, dbHay104 As Boolean
  Dim sReporte As String, sTipo As String
  Dim sDesBanco As String, sCheque As String, sDia As String, sMes As String, sAno As String

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)

Dim porstMRpTmp2 As ADODB.Recordset
   udFecha = Date                      'Fecha en el encabezado.
   Set porstMRpTmp2 = New ADODB.Recordset
   With porstMRpTmp2
    .ActiveConnection = frmTCpbGrd.uocnnMain
'      .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
   End With

' Obtengo la información del comprobante
  With porstMRpTmp2
    .Source = "SELECT c.FehCpb, " & Choose(gsIdioma, "c.GloCpb", "c.GloCpbx") & " AS GloCpb, a.ImpTcb, a.TpoPvs, a.TpoMon, "
    .Source = .Source & "a.MesPvs, a.FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, a.CodAux, d.RazAux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.NroCpb,')')", "('('+a.CodDro+'-'+a.NroCpb+')')") & " AS cComprobante, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    .Source = .Source & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
    .Source = .Source & "a.ImpME, a.ImpMN, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.impME ELSE 0 END) as DebME, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) as HabME, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) as DebMN, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) as HabMN, "
    .Source = .Source & "a.fevdoc, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf "
    .Source = .Source & "FROM ((COCpbCab c "
    .Source = .Source & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.NroCpb=a.NroCpb) "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
    .Source = .Source & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs = '" & gsMesAct & "' "
'2014-06-24 fmt impresion    .Source = .Source & "AND a.CodDro = '" & TxtDato(0).Text & "' AND a.NroCpb = '" & TxtDato(2).Text & "' "

'2016-03-14 ini error fmt impre
'.Source = .Source & "AND ((a.CodDro>='" & txtDato(0) & "' AND a.NroCpb>='" & txtDato(2) & "') AND (a.CodDro<='" & txtDato(1) & "' and a.NroCpb<='" & txtDato(3) & "')) "
        Dim arrv1(1) As String
        arrv1(0) = "a.CodDro"
        arrv1(1) = "a.NroCpb"
.Source = .Source & "AND  " & fConCat(arrv1) & " BETWEEN '" & txtDato(0) & txtDato(2) & "' AND '" & txtDato(1) & txtDato(3) & "' "
'2016-03-14 fin error fmt impre

    .Source = .Source & "ORDER BY a.CodDro, a.NroCpb, a.NroIte"
    .Open
    sDesBanco = "": sCheque = "": sDia = "": sMes = "": sAno = ""
    dsGirado = "": dsGirado2 = ""
    dbHayAux = False
    dbHay104 = False
    Do While Not .EOF
      If Len(Trim(!codaux)) <> 0 And Not dbHayAux Then
        dbHayAux = True
        dsGirado = !razAux
      End If
      If Left(!CodCta, 3) = "104" And Not dbHay104 Then
        dbHay104 = True
        dsFecha = Format(!fehope, "d mmmm yyyy")
        dsImporteNumeros = "********" & Format(CDec(IIf(!tpomon = TPOMON_NAC, !ImpMN, !ImpME)), FORMATO_NUM_1)
        dsImporteLetras = IIf(!tpomon = TPOMON_NAC, gfNumLet(!ImpMN, "0"), gfNumLet(!ImpME, "0")) & "********"
        dsGirado2 = !GloIte
        sDesBanco = !detcta
        sCheque = IIf(IsNull(!RefDoc), "", !RefDoc)
        sDia = Format(!fehope, "dd")
        sMes = Format(!fehope, "mm")
        sAno = Format(!fehope, "yyyy")
      End If
      If dbHayAux And dbHay104 Then Exit Do
      .MoveNext
    Loop
    .MoveFirst
    dsGirado = IIf(dbHayAux = True, dsGirado, dsGirado2) & "********"
  End With
  
  ' Verifico el tipo de impresion
  sReporte = "rptEComPro": sTipo = "C"
'ini 2014-06-24 fmt impresion
''  If MsgBox(Choose(gsIdioma, " Imprimir Cheque Voucher?", "Print Cheque Voucher"), vbQuestion + vbYesNo + vbDefaultButton1, "Consulta") = vbYes Then
''    sReporte = "rptECheVou"
''    sTipo = "V"
''    If Not dbHay104 Then
''      MsgBox Choose(gsIdioma, "El comprobante no tiene alguna cuenta 104.", "The voucher doesn't have any account 104."), vbInformation
''      porstMRpTmp2.Close
''      Set porstMRpTmp2 = Nothing
''      Exit Sub
''    End If
''  End If
'fin 2014-06-24 fmt impresion
    sReporte = "rptECheVou_lst"
  '2015-07-24 error acceso a rmp repo If Index = 0 Then
  If usDEstino = PRN_DEST_GRAF Then
    '2015-07-24 error acceso a rmp repo  sReporte = "rptECheVou_lst"
    'sReporte = "rptechevou_lst_2014_09_18_mh"
    ' Genero el reporte
    gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "LISTADO DE COMPROBANTES", "LISTING OF VOUCHERS"), udFecha, True, True, porstMRpTmp2
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & sReporte & ".rpt"
      '         .WindowShowGroupTree = True
      '[ Formulas adicionales
      .Formulas(6) = "tGirado='" & dsGirado & "'"
      .Formulas(7) = "tFecha='" & dsFecha & "'"
      .Formulas(8) = "tImporteNumeros='" & dsImporteNumeros & "'"
      .Formulas(9) = "tImporteLetras='" & dsImporteLetras & "'"
      .Formulas(10) = "cUsuario='" & gfEnmasc(IIf(IsNull(porstMRpTmp2!UsrMdf), porstMRpTmp2!UsrCre, porstMRpTmp2!UsrMdf)) & "'"
      .Formulas(11) = "cTipo='" & sTipo & "'"
      .WindowState = crptMaximized
      .MarginLeft = unMargenIzquierdo
      .Destination = 0
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRpTmp2
      .LoadReport gsRutRpt & sReporte & ".mrp"
      gpEncabezadoMRp MRViewer, Choose(gsIdioma, "LISTADO DE COMPROBANTES", "LISTING OF VOUCHERS"), udFecha, True, chkImpFecha.Value
      '[Parámetros adicionales.
      .Parameters("tGirado") = dsGirado
      .Parameters("tFecha") = dsFecha
      .Parameters("tImporteNumeros") = dsImporteNumeros
      .Parameters("tImporteLetras") = dsImporteLetras
      .Parameters("cUsuario") = gfEnmasc(IIf(IsNull(porstMRpTmp2!UsrMdf), porstMRpTmp2!UsrCre, porstMRpTmp2!UsrMdf))
      .Parameters("cDesBanco") = sDesBanco
      .Parameters("cCheque") = sCheque
      .Parameters("cDia") = sDia
      .Parameters("cMes") = sMes
      .Parameters("cAno") = sAno
      ']
      .PreviewReport
      .UnLoadReport
    End With
    Set MRViewer = Nothing
  End If
  porstMRpTmp2.Close
  Set porstMRpTmp2 = Nothing

End Sub


Private Sub cmdImprimir2_Click_2014_06_25(Index As Integer)
  Dim dsFecha As String, dsGirado As String, dsGirado2 As String, dsImporteNumeros As String, dsImporteLetras As String
  Dim dbHayAux As Boolean, dbHay104 As Boolean
  Dim sReporte As String, sTipo As String
  Dim sDesBanco As String, sCheque As String, sDia As String, sMes As String, sAno As String

   udFecha = Date                      'Fecha en el encabezado.
   Set porstMRp = New ADODB.Recordset
   With porstMRp
    .ActiveConnection = frmTCpbGrd.uocnnMain
'      .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
   End With
'*****************************************************

   Set porstMRp2 = New ADODB.Recordset
   With porstMRp2
    .ActiveConnection = frmTCpbGrd.uocnnMain
'      .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
   End With



  ' Elimino y genero el archivo del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpECheVou_lst", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#trpECheVou_lst') DROP TABLE #trpECheVou_lst")
Dim CadCrystal As String
  'CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE trpECheVou_lst (", "CREATE TABLE " & ps_Prefijo & "trpECheVou_lst (")
   CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE trpECheVou_lst ", "CREATE TABLE " & ps_Prefijo & "trpECheVou_lst (")
   CadCrystal = CadCrystal & " SELECT c.FehCpb, " & Choose(gsIdioma, "c.GloCpb", "c.GloCpbx") & " AS GloCpb, a.ImpTcb, a.TpoPvs, a.TpoMon, "
    CadCrystal = CadCrystal & "a.MesPvs, a.FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, a.CodAux, d.RazAux, "
    CadCrystal = CadCrystal & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.NroCpb,')')", "('('+a.CodDro+'-'+a.NroCpb+')')") & " AS cComprobante, "
    CadCrystal = CadCrystal & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    CadCrystal = CadCrystal & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
    CadCrystal = CadCrystal & "a.ImpME, a.ImpMN, "
    CadCrystal = CadCrystal & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.impME ELSE 0 END) as DebME, "
    CadCrystal = CadCrystal & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) as HabME, "
    CadCrystal = CadCrystal & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) as DebMN, "
    CadCrystal = CadCrystal & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) as HabMN, "
    CadCrystal = CadCrystal & "a.fevdoc, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf,a.NroIte "
    CadCrystal = CadCrystal & "FROM ((COCpbCab c "
    CadCrystal = CadCrystal & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.NroCpb=a.NroCpb) "
    CadCrystal = CadCrystal & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    CadCrystal = CadCrystal & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
    CadCrystal = CadCrystal & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
    CadCrystal = CadCrystal & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    CadCrystal = CadCrystal & "AND a.MesPvs = '" & gsMesAct & "' "
'2014-06-24 fmt impresion    .Source = .Source & "AND a.CodDro = '" & TxtDato(0).Text & "' AND a.NroCpb = '" & TxtDato(2).Text & "' "
    CadCrystal = CadCrystal & "AND ((a.CodDro>='" & txtDato(0) & "' AND a.NroCpb>='" & txtDato(2) & "') AND (a.CodDro<='" & txtDato(1) & "' and a.NroCpb<='" & txtDato(3) & "')) "
    CadCrystal = CadCrystal & "ORDER BY a.CodDro, a.NroCpb, a.NroIte"
  pocnnMain.Execute CadCrystal
Dim xSql2 As String
  
  With porstMRp2
    .Source = "SELECT * FROM  trpECheVou_lst ORDER BY MesPvs,cComprobante,NroIte"
    .Open
    Dim xcComprobante As String
    Dim mq1, mq2 As String
    mq1 = ""
    mq2 = ""
    Dim xIte As Integer
    Dim n As Integer
    xIte = 0
'************************************
Dim xtotlin As Integer
'xtotlin = 24
'xtotlin = 22
'xtotlin = 25
xtotlin = 23

If xIte = 0 Then
    Do While Not .EOF
        If mq1 <> !cComprobante Then
            xIte = 0
            mq1 = !cComprobante
        End If
        xIte = !NroIte
        .MoveNext
        'xIte = xIte + 1
       If Not .EOF Then
        If mq1 <> !cComprobante Then
        
             If xIte < xtotlin Then
         xIte = xIte + 1
                For n = xIte To xtotlin - xIte
                    xSql2 = ""
                    xSql2 = xSql2 & "INSERT trpECheVou_lst ("
                    xSql2 = xSql2 & "    FehCpb,"
                    xSql2 = xSql2 & "    GloCpb,"
                    xSql2 = xSql2 & "    TpoMon,"
                    xSql2 = xSql2 & "    MesPvs,"
                    xSql2 = xSql2 & "    FehOpe,"
                    xSql2 = xSql2 & "    CodCta,"
                    xSql2 = xSql2 & "    DetCta,"
                    xSql2 = xSql2 & "    CodCCo,"
                    xSql2 = xSql2 & "    CodAux,"
                    xSql2 = xSql2 & "    RazAux,"
                    xSql2 = xSql2 & "    cComprobante,"
                    xSql2 = xSql2 & "    cDocumento,"
                    xSql2 = xSql2 & "    CodTDc,"
                    xSql2 = xSql2 & "    SerDoc,"
                    xSql2 = xSql2 & "    NroDoc,"
                    xSql2 = xSql2 & "    RefDoc,"
                    xSql2 = xSql2 & "    GloIte,"
                    xSql2 = xSql2 & "    DebME,"
                    xSql2 = xSql2 & "    HabME,"
                    xSql2 = xSql2 & "    DebMN,"
                    xSql2 = xSql2 & "    HabMN,"
                    xSql2 = xSql2 & "    FeVDoc,"
                    xSql2 = xSql2 & "    UsrCre,"
                    xSql2 = xSql2 & "    FyHCre,"
                    xSql2 = xSql2 & "    NroIte"
'NroIte
                    xSql2 = xSql2 & ") VALUE ("
                    xSql2 = xSql2 & "    '2014-04-01',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '" & gsMesAct & "',"
                    xSql2 = xSql2 & "    '2014-04-01',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '" & mq1 & "',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    '2014-04-01',"
                    xSql2 = xSql2 & "    'rcastro',"
                    xSql2 = xSql2 & "    '2014-04-04 08:56:20',"
                    xSql2 = xSql2 & "    '" & Str(n) & "' "
                    xSql2 = xSql2 & ")"
                  pocnnMain.Execute xSql2
                  'pocnnMain.Execute "INSERT trpECheVou_lst (MesPvs,cComprobante) VALUE ('" & gsMesAct & "','" & mq1 & "')"
                 Next
             End If
         End If
       Else
       n = n
'-------------------
             If xIte < xtotlin Then
        xIte = xIte + 1
             
                 For n = xIte To xtotlin - xIte
                    xSql2 = ""
                    xSql2 = xSql2 & "INSERT trpECheVou_lst ("
                    xSql2 = xSql2 & "    FehCpb,"
                    xSql2 = xSql2 & "    GloCpb,"
                    xSql2 = xSql2 & "    TpoMon,"
                    xSql2 = xSql2 & "    MesPvs,"
                    xSql2 = xSql2 & "    FehOpe,"
                    xSql2 = xSql2 & "    CodCta,"
                    xSql2 = xSql2 & "    DetCta,"
                    xSql2 = xSql2 & "    CodCCo,"
                    xSql2 = xSql2 & "    CodAux,"
                    xSql2 = xSql2 & "    RazAux,"
                    xSql2 = xSql2 & "    cComprobante,"
                    xSql2 = xSql2 & "    cDocumento,"
                    xSql2 = xSql2 & "    CodTDc,"
                    xSql2 = xSql2 & "    SerDoc,"
                    xSql2 = xSql2 & "    NroDoc,"
                    xSql2 = xSql2 & "    RefDoc,"
                    xSql2 = xSql2 & "    GloIte,"
                    xSql2 = xSql2 & "    DebME,"
                    xSql2 = xSql2 & "    HabME,"
                    xSql2 = xSql2 & "    DebMN,"
                    xSql2 = xSql2 & "    HabMN,"
                    xSql2 = xSql2 & "    FeVDoc,"
                    xSql2 = xSql2 & "    UsrCre,"
                    xSql2 = xSql2 & "    FyHCre,"
                    xSql2 = xSql2 & "    NroIte"
                    xSql2 = xSql2 & ") VALUE ("
                    xSql2 = xSql2 & "    '2014-04-01',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '" & gsMesAct & "',"
                    xSql2 = xSql2 & "    '2014-04-01',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '" & mq1 & "',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    '',"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    0,"
                    xSql2 = xSql2 & "    '2014-04-01',"
                    xSql2 = xSql2 & "    'rcastro',"
                    xSql2 = xSql2 & "    '2014-04-04 08:56:20',"
                    xSql2 = xSql2 & "    '" & Str(n) & "' "
                    
                    xSql2 = xSql2 & ")"
                   pocnnMain.Execute xSql2
                
                  'pocnnMain.Execute "INSERT trpECheVou_lst (MesPvs,cComprobante) VALUE ('" & gsMesAct & "','" & mq1 & "')"
                 Next
             End If
'-------------------
       End If
    Loop
End If
'********************************

  End With


'*****************************************************
' Obtengo la información del comprobante
  With porstMRp
    .Source = "SELECT * FROM  trpECheVou_lst ORDER BY MesPvs,cComprobante,NroIte"
  
'    .Source = "SELECT c.FehCpb, " & Choose(gsIdioma, "c.GloCpb", "c.GloCpbx") & " AS GloCpb, a.ImpTcb, a.TpoPvs, a.TpoMon, "
'    .Source = .Source & "a.MesPvs, a.FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, a.CodAux, d.RazAux, "
'    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.NroCpb,')')", "('('+a.CodDro+'-'+a.NroCpb+')')") & " AS cComprobante, "
'    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
'    .Source = .Source & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
'    .Source = .Source & "a.ImpME, a.ImpMN, "
'    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.impME ELSE 0 END) as DebME, "
'    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) as HabME, "
'    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) as DebMN, "
'    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) as HabMN, "
'    .Source = .Source & "a.fevdoc, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf "
'    .Source = .Source & "FROM ((COCpbCab c "
'    .Source = .Source & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.NroCpb=a.NroCpb) "
'    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
'    .Source = .Source & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
'    .Source = .Source & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
'    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
'    .Source = .Source & "AND a.MesPvs = '" & gsMesAct & "' "
''2014-06-24 fmt impresion    .Source = .Source & "AND a.CodDro = '" & TxtDato(0).Text & "' AND a.NroCpb = '" & TxtDato(2).Text & "' "
'    .Source = .Source & "AND ((a.CodDro>='" & TxtDato(0) & "' AND a.NroCpb>='" & TxtDato(2) & "') AND (a.CodDro<='" & TxtDato(1) & "' and a.NroCpb<='" & TxtDato(3) & "')) "
'
'    .Source = .Source & "ORDER BY a.CodDro, a.NroCpb, a.NroIte"

    .Open
    sDesBanco = "": sCheque = "": sDia = "": sMes = "": sAno = ""
    dsGirado = "": dsGirado2 = ""
    dbHayAux = False
    dbHay104 = False
'cortar ini *********************************
''    Do While Not .EOF
''      If Len(Trim(!codaux)) <> 0 And Not dbHayAux Then
''        dbHayAux = True
''        dsGirado = !razAux
''      End If
''      If Left(!codcta, 3) = "104" And Not dbHay104 Then
''        dbHay104 = True
''        dsFecha = Format(!fehope, "d mmmm yyyy")
''        dsImporteNumeros = "********" & Format(CDec(IIf(!tpomon = TPOMON_NAC, !ImpMN, !ImpME)), FORMATO_NUM_1)
''        dsImporteLetras = IIf(!tpomon = TPOMON_NAC, gfNumLet(!ImpMN, "0"), gfNumLet(!ImpME, "0")) & "********"
''        dsGirado2 = !GloIte
''        sDesBanco = !detcta
''        sCheque = IIf(IsNull(!refdoc), "", !refdoc)
''        sDia = Format(!fehope, "dd")
''        sMes = Format(!fehope, "mm")
''        sAno = Format(!fehope, "yyyy")
''      End If
''      If dbHayAux And dbHay104 Then Exit Do
''      .MoveNext
''    Loop
''    .MoveFirst
''    dsGirado = IIf(dbHayAux = True, dsGirado, dsGirado2) & "********"
'cortar ini ***********************************
  End With
  
  ' Verifico el tipo de impresion
  sReporte = "rptEComPro": sTipo = "C"
'ini 2014-06-24 fmt impresion
''  If MsgBox(Choose(gsIdioma, " Imprimir Cheque Voucher?", "Print Cheque Voucher"), vbQuestion + vbYesNo + vbDefaultButton1, "Consulta") = vbYes Then
''    sReporte = "rptECheVou"
''    sTipo = "V"
''    If Not dbHay104 Then
''      MsgBox Choose(gsIdioma, "El comprobante no tiene alguna cuenta 104.", "The voucher doesn't have any account 104."), vbInformation
''      porstMRp.Close
''      Set porstMRp = Nothing
''      Exit Sub
''    End If
''  End If
'fin 2014-06-24 fmt impresion
  If Index = 0 Then
    sReporte = "rptECheVou_lst"
    ' Genero el reporte
    gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "LISTADO DE COMPROBANTES", "LISTING OF VOUCHERS"), udFecha, True, True, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & sReporte & ".rpt"
      '         .WindowShowGroupTree = True
      '[ Formulas adicionales
'***********************
''      .Formulas(6) = "tGirado='" & dsGirado & "'"
''      .Formulas(7) = "tFecha='" & dsFecha & "'"
''      .Formulas(8) = "tImporteNumeros='" & dsImporteNumeros & "'"
''      .Formulas(9) = "tImporteLetras='" & dsImporteLetras & "'"
''      .Formulas(10) = "cUsuario='" & gfEnmasc(IIf(IsNull(porstMRp!UsrMdf), porstMRp!UsrCre, porstMRp!UsrMdf)) & "'"
''      .Formulas(11) = "cTipo='" & sTipo & "'"
'************************
'      .Formulas(6) = "tGirado='r'"
'      .Formulas(7) = "tFecha='26/06/2014'"
'      .Formulas(8) = "tImporteNumeros='1500'"
'      .Formulas(9) = "tImporteLetras='MIL QUINIENTOS'"
'      .Formulas(10) = "cUsuario='rcastro'"
'      .Formulas(11) = "cTipo='" & sTipo & "'"

      .Formulas(6) = "tGirado='r'"
      .Formulas(7) = "tFecha=''"
      .Formulas(8) = "tImporteNumeros=''"
      .Formulas(9) = "tImporteLetras=''"
      .Formulas(10) = "cUsuario=''"
      .Formulas(11) = "cTipo='" & sTipo & "'"

      .WindowState = crptMaximized
      .MarginLeft = unMargenIzquierdo
      .Destination = 0
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & sReporte & ".mrp"
      gpEncabezadoMRp MRViewer, Choose(gsIdioma, "LISTADO DE COMPROBANTES", "LISTING OF VOUCHERS"), udFecha, True
      '[Parámetros adicionales.
      .Parameters("tGirado") = dsGirado
      .Parameters("tFecha") = dsFecha
      .Parameters("tImporteNumeros") = dsImporteNumeros
      .Parameters("tImporteLetras") = dsImporteLetras
      .Parameters("cUsuario") = gfEnmasc(IIf(IsNull(porstMRp!UsrMdf), porstMRp!UsrCre, porstMRp!UsrMdf))
      .Parameters("cDesBanco") = sDesBanco
      .Parameters("cCheque") = sCheque
      .Parameters("cDia") = sDia
      .Parameters("cMes") = sMes
      .Parameters("cAno") = sAno
      ']
      .PreviewReport
      .UnLoadReport
    End With
    Set MRViewer = Nothing
  End If
  porstMRp.Close
  Set porstMRp = Nothing

End Sub


Private Sub cmdPrnDirec_Click()
'http://www.forosdelweb.com/f69/como-imprimir-factura-desde-visual-basic-114701/
''Printer.ScaleMode = 4 'define que vas a imprimir en formato caracteres
''Printer.Orientation = 1 'define la orientacion del papel (normal o apaisado)
''Printer.CurrentX = 0 ' la coordenada x en que se va a imprimir
''Printer.CurrentY = 0 ' la coordenada y en la que se va a imprimir
''Dim Imagen As Picture
''Set Imagen = LoadPicture("negrochico.bmp")
''Printer.PaintPicture Imagen, 0, 0 ', 4300, 600 ' que imprima imagen en x = 0 en Y = 0, ancho = 4300, alto = 600
''Set con = base.OpenRecordset("select * from ccomanda1 where impreso='NO' and comanda.mesa='" & mesa & "' and tipodoc = 0 and numdoc = 0 order by tipo", dbOpenDynaset)
''old = Printer.FontSize
''Printer.Font = "Trebuchet MS" ' define el tipo de letra
''Printer.FontSize = 14 ' el tamaño de la letra
''Printer.CurrentX = 0
''Printer.CurrentY = 8
''Printer.Print " COMANDA MESA " & con("comanda.mesa") ' el printer.print imprime en la impresora, segun las coordenadas currentx(columna) y currenty(fila) que le diste
''Printer.CurrentX = 0
''Printer.CurrentY = 10
''Printer.Print "Garzón " & con("nombre")
''Printer.CurrentX = 0
''Printer.CurrentY = 11
''Printer.Print "Sector " & con("descripcionu")
''Printer.EndDoc ' finaliza el documento y lo envia a impresora

''Dim xSql2 As String
''xSql2 = ""
'''********************************
''    xSql2 = "SELECT c.FehCpb, " & Choose(gsIdioma, "c.GloCpb", "c.GloCpbx") & " AS GloCpb, a.ImpTcb, a.TpoPvs, a.TpoMon, "
''    xSql2 = xSql2 & "a.MesPvs, a.FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, a.CodAux, d.RazAux, "
''    xSql2 = xSql2 & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.NroCpb,')')", "('('+a.CodDro+'-'+a.NroCpb+')')") & " AS cComprobante, "
''    xSql2 = xSql2 & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
''    xSql2 = xSql2 & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
''    xSql2 = xSql2 & "a.ImpME, a.ImpMN, "
''    xSql2 = xSql2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.impME ELSE 0 END) as DebME, "
''    xSql2 = xSql2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) as HabME, "
''    xSql2 = xSql2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) as DebMN, "
''    xSql2 = xSql2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) as HabMN, "
''    xSql2 = xSql2 & "a.fevdoc, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf "
''    xSql2 = xSql2 & "FROM ((COCpbCab c "
''    xSql2 = xSql2 & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.NroCpb=a.NroCpb) "
''    xSql2 = xSql2 & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
''    xSql2 = xSql2 & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
''    xSql2 = xSql2 & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
''    xSql2 = xSql2 & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
''    xSql2 = xSql2 & "AND a.MesPvs = '" & gsMesAct & "' "
'''2014-06-24 fmt impresion    xSql2 = xSql2 & "AND a.CodDro = '" & TxtDato(0).Text & "' AND a.NroCpb = '" & TxtDato(2).Text & "' "
''    xSql2 = xSql2 & "AND ((a.CodDro>='" & TxtDato(0) & "' AND a.NroCpb>='" & TxtDato(2) & "') AND (a.CodDro<='" & TxtDato(1) & "' and a.NroCpb<='" & TxtDato(3) & "')) "
''    xSql2 = xSql2 & "ORDER BY a.CodDro, a.NroCpb, a.NroIte"
'********************************
   Set porstMRp = New ADODB.Recordset
   With porstMRp
    .ActiveConnection = frmTCpbGrd.uocnnMain
'      .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
   End With
  With porstMRp
    .Source = "SELECT c.FehCpb, " & Choose(gsIdioma, "c.GloCpb", "c.GloCpbx") & " AS GloCpb, a.ImpTcb, a.TpoPvs, a.TpoMon, "
    .Source = .Source & "a.MesPvs, a.FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, a.CodAux, d.RazAux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.NroCpb,')')", "('('+a.CodDro+'-'+a.NroCpb+')')") & " AS cComprobante, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    .Source = .Source & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
    .Source = .Source & "a.ImpME, a.ImpMN, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.impME ELSE 0 END) as DebME, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) as HabME, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) as DebMN, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) as HabMN, "
    .Source = .Source & "a.fevdoc, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf "
    .Source = .Source & "FROM ((COCpbCab c "
    .Source = .Source & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.NroCpb=a.NroCpb) "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
    .Source = .Source & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs = '" & gsMesAct & "' "
'2014-06-24 fmt impresion    .Source = .Source & "AND a.CodDro = '" & TxtDato(0).Text & "' AND a.NroCpb = '" & TxtDato(2).Text & "' "
    .Source = .Source & "AND ((a.CodDro>='" & txtDato(0) & "' AND a.NroCpb>='" & txtDato(2) & "') AND (a.CodDro<='" & txtDato(1) & "' and a.NroCpb<='" & txtDato(3) & "')) "
    
    .Source = .Source & "ORDER BY a.CodDro, a.NroCpb, a.NroIte"
    .Open

Printer.ScaleMode = 4 'define que vas a imprimir en formato caracteres
Printer.Orientation = 1 'define la orientacion del papel (normal o apaisado)
Printer.CurrentX = 0 ' la coordenada x en que se va a imprimir
Printer.CurrentY = 0 ' la coordenada y en la que se va a imprimir
'Dim Imagen As Picture
'Set Imagen = LoadPicture("negrochico.bmp")
'Printer.PaintPicture Imagen, 0, 0 ', 4300, 600 ' que imprima imagen en x = 0 en Y = 0, ancho = 4300, alto = 600
'Set con = base.OpenRecordset(xSql2, dbOpenDynaset)
'old = Printer.FontSize
Printer.Font = "Trebuchet MS" ' define el tipo de letra
Printer.FontSize = 14 ' el tamaño de la letra
Printer.CurrentX = 0
Printer.CurrentY = 8
Printer.Print " COMANDA MESA " & !RefDoc ' el printer.print imprime en la impresora, segun las coordenadas currentx(columna) y currenty(fila) que le diste
Printer.CurrentX = 0
Printer.CurrentY = 10
Printer.Print "Garzón " & !glocpb
Printer.CurrentX = 0
Printer.CurrentY = 11
Printer.Print "Sector " & !CodCta
Printer.EndDoc ' finaliza el documento y lo envia a impresora
   End With

End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
   
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
   With porstCodro
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodDro, " & Choose(gsIdioma, " DetDro", " DetDrox") & " AS DetDro "
      .Source = .Source & "FROM CODro "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodDro"
         .Item(dnContador).MaxLength = porstCodro.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diarios", "Comprobantes")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journals", "Vouchers")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCodro
      .MoveLast
      txtDato(1).Text = !coddro
      .MoveFirst
      txtDato(0).Text = !coddro
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
   txtDato(2).Text = "000000"
   txtDato(3).Text = "999999"
   
  'Características de impresión.
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
   porstCodro.Close
   pocnnMain.Close
   Set porstCodro = Nothing
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
Sub xxx()

End Sub
Private Sub cmdImprimir_Click(Index As Integer)
  
  ppHabilitacion False

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.CodDro, a.NroCpb, "
    .Source = .Source & "a.FehCpb, " & Choose(gsIdioma, "a.GloCpb", "a.GloCpbx") & " AS GloCpb, "
    .Source = .Source & "b.NroIte, b.FehOpe, "
    .Source = .Source & "b.CodCta, b.CodCCo, "
    .Source = .Source & "b.CodAux, b.CodTDc, "
    .Source = .Source & "b.TpoPvs, b.SerDoc, "
    .Source = .Source & "b.NroDoc, b.RefDoc, "
    .Source = .Source & "b.FeEDoc, b.FeVDoc, "
    .Source = .Source & "b.FeRDoc, " & Choose(gsIdioma, "b.GloIte", "b.GloItex") & " AS GloIte, "
    .Source = .Source & "(CASE a.TpoGnr WHEN " & TPOGNR_DRO & " THEN '" & TPOGNR_DRO_TXT & "' WHEN " & TPOGNR_CPR & " THEN '" & TPOGNR_CPR_TXT & "' WHEN " & TPOGNR_VTA & " THEN '" & TPOGNR_VTA_TXT & "' WHEN " & TPOGNR_HPR & " THEN '" & TPOGNR_HPR_TXT & "' WHEN " & TPOGNR_DST & " THEN '" & TPOGNR_DST_TXT & "' WHEN " & TPOGNR_DCA & " THEN '" & TPOGNR_DCA_TXT & "' WHEN " & TPOGNR_APE & " THEN '" & TPOGNR_APE_TXT & "' ELSE '" & TPOGNR_CIE_TXT & "' END) AS ccTpoGnr, "
    .Source = .Source & "b.TpotCb,b.TpoCtb, b.ImpME, b.ImpMN "
    .Source = .Source & "FROM cocpbcab a "
    .Source = .Source & "LEFT JOIN cocpbdet b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb = b.NroCpb "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs='" & gsMesAct & "' "
    .Source = .Source & "AND ((a.CodDro>='" & txtDato(0) & "' AND a.NroCpb>='" & txtDato(2) & "') AND (a.CodDro<='" & txtDato(1) & "' and a.NroCpb<='" & txtDato(3) & "')) "
    .Source = .Source & "ORDER BY b.CodDro, b.NroCpb, b.NroIte"
    .Open
  End With

  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLCpb.rpt"
      .SelectionFormula = "{trptLCpb.CodDro} IN '" & txtDato(0).Text & "' TO '" & txtDato(1).Text & "' "
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
      .LoadReport gsRutRpt & "rptLCpb.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption, udFecha, True)
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
If KeyAscii = 13 Then
   If Index = 2 Or Index = 3 Then
       txtDato(Index) = gfCeros(txtDato(Index), txtDato(Index).MaxLength, 0, "0")
   End If
End If
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index    'Completa con ceros a la izquierda.
   Case 2, 3                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
