VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPFil 
   Caption         =   "[título]"
   ClientHeight    =   4035
   ClientLeft      =   2220
   ClientTop       =   2205
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6315
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   3135
      Width           =   1335
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2715
      Width           =   1260
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   6315
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3495
      Width           =   6315
      Begin VB.CommandButton cmdExporta 
         Caption         =   "&Genera Archivo"
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
         Left            =   3510
         Picture         =   "frmpfil.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   1170
         Picture         =   "frmpfil.frx":0102
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
         Picture         =   "frmpfil.frx":0204
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
         Left            =   5130
         Picture         =   "frmpfil.frx":0736
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
         Left            =   2340
         TabIndex        =   2
         Top             =   0
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   2490
      Left            =   45
      TabIndex        =   7
      Top             =   90
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   4392
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Archivos de Información"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlUbicacion 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   75
      TabIndex        =   10
      Top             =   3030
      Width           =   2310
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   4290
      TabIndex        =   6
      Top             =   2760
      Width           =   630
   End
End
Attribute VB_Name = "frmPFil"
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
Private porstMRpRs  As ADODB.Recordset
Private porstCoFil  As ADODB.Recordset
Private porstCoCfg As ADODB.Recordset
Private psAnoOld As String

Private Sub cmdExporta_Click()
  Dim s_MesIni As String, s_MesFin As String
  Dim s_SalAno As String
  Dim sArchivo As String, sCadena As String
  Dim sCaracter As String, sMoneda As String, sRegistro As String
  Dim sSentencia As String, s_SaldoDeb As String, s_SaldoHab As String
  Dim nImporte As Double, nRegistro As Long
  Dim nContador As Integer, nColFinal As Integer
  Dim sCenCosto As String
  
  s_SalAno = IIf(gsMesApe > gsMesAct, gsAnoAct - 1, gsAnoAct)
  s_MesIni = gsMesApe
  s_MesFin = gsMesAct
  
  ppHabilitacion False
  lblProgreso.Caption = Choose(gsIdioma, "Procesando Información...", "Processing Information...")
 
  '[ Inicializo variables y nombre de archivo
  sArchivo = gsRUCEmp & gsAnoAct & gsMesAct & ".ema"
  sCaracter = IIf(IsNull(porstCoFil!sepcar), ";", porstCoFil!sepcar)
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
  
  ' Inicializo el archivo temporal
  pocnnMain.Execute "DELETE FROM cotmprpt WHERE codemp='" & gsCodEmp & "' AND usrcre='" & gsCodUsr & "' AND nomrpt='rptpfil'"
  
  ' Obtengo niveles de cuenta
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT DISTINCTROW " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(codcta) AS nivel "
    .Source = .Source & "FROM cofildet "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND codfil='" & porstCoFil!codfil & "' "
    .Source = .Source & "ORDER BY 1"
    .Open
  End With
  ' Verifico si existe registros
  If porstMRp.RecordCount > 0 Then
    porstMRp.MoveFirst
    Do While Not porstMRp.EOF
      ' Inserto la informacion
      sSentencia = "INSERT INTO cotmprpt (codcta, detcta, codcco, detcco, tposdo, coddro, numcol1, numcol2, numcol3,  "
      sSentencia = sSentencia & "numcol4, numcol5, numcol6, numcol7, numcol8, numcol11, numcol12, codemp, usrcre, nomrpt) "
      ' No incluye centro de costos
      sSentencia = sSentencia & "SELECT DISTINCTROW cta.codcta, " & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, "
      sSentencia = sSentencia & "Null AS codcco, Null AS detcco, "
      sSentencia = sSentencia & "cta.natcta, fil.nrolin, fil.coldet1, fil.coldet2, fil.coldet3, "
      sSentencia = sSentencia & "fil.coldet4, fil.coldet5, fil.coldet6, fil.coldet7, fil.coldet8, "
      sSentencia = sSentencia & "0.00 AS nDebe, "
      sSentencia = sSentencia & "0.00 AS nHaber, "
      sSentencia = sSentencia & "cta.codemp, '" & gsCodUsr & "', 'rptpfil' "
      sSentencia = sSentencia & "FROM cocta cta "
      sSentencia = sSentencia & "INNER JOIN cofildet fil ON cta.codemp=fil.codemp AND Left(cta.codcta, " & porstMRp!nivel & ")=fil.codcta "
      sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(cta.codcta)=" & Right(gsNivCta, 1) & " "
      sSentencia = sSentencia & "AND cta.indcco=" & INDCCO_INA & " "
      sSentencia = sSentencia & "AND fil.codfil='" & porstCoFil!codfil & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(fil.codcta)=" & porstMRp!nivel & " "
      ' Continene Centro de costos
      sSentencia = sSentencia & "UNION "
      sSentencia = sSentencia & "SELECT DISTINCTROW det.codcta, " & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, "
      sSentencia = sSentencia & "det.codcco, " & Choose(gsIdioma, "cco.detcco", "cco.detccox") & " AS detcco, "
      sSentencia = sSentencia & "cta.natcta, fil.nrolin, fil.coldet1, fil.coldet2, fil.coldet3, "
      sSentencia = sSentencia & "fil.coldet4, fil.coldet5, fil.coldet6, fil.coldet7, fil.coldet8, "
      sSentencia = sSentencia & "0.00 AS nDebe, "
      sSentencia = sSentencia & "0.00 AS nHaber, "
      sSentencia = sSentencia & "det.codemp, '" & gsCodUsr & "', 'rptpfil' "
      sSentencia = sSentencia & "FROM cofildet fil "
      sSentencia = sSentencia & "LEFT JOIN cocpbdet det ON fil.codemp=det.codemp AND fil.codcta=Left(det.codcta, " & porstMRp!nivel & ") "
      sSentencia = sSentencia & "LEFT JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.codcta=cta.codcta "
      sSentencia = sSentencia & "LEFT JOIN cocco cco ON det.codemp=cco.codemp AND det.pdoano=cco.pdoano AND det.codcco=cco.codcco "
      sSentencia = sSentencia & "WHERE fil.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND fil.codfil='" & porstCoFil!codfil & "' "
      sSentencia = sSentencia & "AND fil.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(fil.codcta)=" & porstMRp!nivel & " "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & ">='" & s_SalAno & s_MesIni & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & "<='" & gsAnoAct & s_MesFin & "' "
      sSentencia = sSentencia & "AND det.mespvs>'00' "
      sSentencia = sSentencia & "AND det.mespvs<'13' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.codcco, '')<>'' "
      sSentencia = sSentencia & "GROUP BY det.codcta, det.codcco "
      sSentencia = sSentencia & "ORDER BY codcta, nrolin"
      pocnnMain.Execute sSentencia, nRegistro
      porstMRp.MoveNext
    Loop
  End If
  ' Genero temporal de acumulado de movimientos
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsaldo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 9)='#tmpsaldo') DROP TABLE #tmpsaldo")
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE #") & "tmpsaldo ("
  sSentencia = sSentencia & "codemp char(3) NOT NULL default '', "
  sSentencia = sSentencia & "codcta varchar(16) NOT NULL default '', "
  sSentencia = sSentencia & "codcco varchar(5) default NULL , "
  sSentencia = sSentencia & "cargo decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "abono decimal(12,2) NOT NULL default '0.00')"
  pocnnMain.Execute sSentencia
      
  ' Inserto acumulado de movimientos
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sSentencia = "INSERT INTO " & ps_Prefijo & "tmpsaldo (codemp, codcta, codcco, cargo, abono) "
  sSentencia = sSentencia & "SELECT DISTINCTROW det.codemp, det.codcta, det.codcco, "
  sSentencia = sSentencia & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END), 2) AS cargo, "
  sSentencia = sSentencia & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END), 2) AS abono "
  sSentencia = sSentencia & "FROM cocpbdet det "
  sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & ">='" & s_SalAno & s_MesIni & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & "<='" & gsAnoAct & s_MesFin & "' "
  sSentencia = sSentencia & "AND det.mespvs>'00' "
  sSentencia = sSentencia & "AND det.mespvs<'13' "
  sSentencia = sSentencia & "GROUP BY det.codcta, det.codcco "
  sSentencia = sSentencia & "ORDER BY det.codcta, det.codcco"
  pocnnMain.Execute sSentencia, nRegistro
  ' Actualizo acumulado movimientos
  sSentencia = "UPDATE cotmprpt tmp, " & ps_Prefijo & "tmpsaldo acu "
  sSentencia = sSentencia & "SET tmp.numcol11=acu.cargo, tmp.numcol12=acu.abono "
  sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND tmp.usrcre='" & gsCodUsr & "' "
  sSentencia = sSentencia & "AND tmp.nomrpt='rptpfil' "
  sSentencia = sSentencia & "AND acu.codemp=tmp.codemp "
  sSentencia = sSentencia & "AND acu.codcta=tmp.codcta "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "AND IFNULL(acu.codcco, '')=IFNULL(tmp.codcco, '')"
  Else
    sSentencia = sSentencia & "AND ISNULL(acu.codcco, '')=ISNULL(tmp.codcco, '')"
  End If
  pocnnMain.Execute sSentencia, nRegistro
  ' Elimino temporal de acumulado de movimientos
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsaldo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 9)='#tmpsaldo') DROP TABLE #tmpsaldo")
      
  ' Acumulación de saldos
  s_SaldoDeb = "ROUND(": s_SaldoHab = "ROUND("
  For nContador = 0 To gsMesApe - 1
    s_SaldoDeb = s_SaldoDeb & "sal.AcuD" & Format(Trim(nContador), "00") & "_" & sMoneda & IIf(nContador = gsMesApe - 1, "", "+")
    s_SaldoHab = s_SaldoHab & "sal.AcuH" & Format(Trim(nContador), "00") & "_" & sMoneda & IIf(nContador = gsMesApe - 1, "", "+")
  Next nContador
  s_SaldoDeb = s_SaldoDeb & ", 2)"
  s_SaldoHab = s_SaldoHab & ", 2)"
        
  ' Actualizo saldos inciales de cuentas de balance
  sSentencia = "UPDATE cotmprpt tmp, coctaacu sal, cocta cta "
  sSentencia = sSentencia & "SET tmp.numcol11=tmp.numcol11+" & s_SaldoDeb & ", "
  sSentencia = sSentencia & "tmp.numcol12=tmp.numcol12+" & s_SaldoHab & " "
  sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND tmp.usrcre='" & gsCodUsr & "' "
  sSentencia = sSentencia & "AND tmp.nomrpt='rptpfil' "
  sSentencia = sSentencia & "AND sal.codemp=tmp.codemp "
  sSentencia = sSentencia & "AND sal.pdoano='" & s_SalAno & "' "
  sSentencia = sSentencia & "AND sal.codcta=tmp.codcta "
  sSentencia = sSentencia & "AND cta.codemp=sal.codemp "
  sSentencia = sSentencia & "AND cta.pdoano=sal.pdoano "
  sSentencia = sSentencia & "AND cta.codcta=sal.codcta "
  sSentencia = sSentencia & "AND cta.tposdo='" & TPOSDO_INV & "'"
  pocnnMain.Execute sSentencia, nRegistro
  
  ' Obtengo la columna final
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT DISTINCTROW coldet1, coldet2, coldet3, coldet4, coldet5, coldet6, coldet7, coldet8 "
    .Source = .Source & "FROM cofildet "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND codfil='" & porstCoFil!codfil & "'"
    .Open
  End With
  nColFinal = 5
  If porstMRp.RecordCount > 0 Then
    porstMRp.MoveFirst
    Do While Not porstMRp.EOF
      For nContador = 1 To 8
        If porstMRp("coldet" & nContador) <> 0 Then
          nColFinal = IIf(nContador > nColFinal, nContador, nColFinal)
        End If
      Next nContador
      porstMRp.MoveNext
    Loop
  End If
  
  ' Obtengo la informacion del archivo
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT codcta, detcta, codcco, detcco, tposdo AS natcta, coddro AS nrolin, numcol1 AS coldet1, numcol2 AS coldet2, numcol3 AS coldet3, "
    .Source = .Source & "numcol4 AS coldet4, numcol5 AS coldet5, numcol6 AS coldet6, numcol7 AS coldet7, numcol8 AS coldet8, numcol11 AS ndebe, numcol12 AS nhaber "
    .Source = .Source & "FROM cotmprpt "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND usrcre='" & gsCodUsr & "' "
    .Source = .Source & "AND nomrpt='rptpfil' "
    .Source = .Source & "ORDER BY 1, 3, 6"
    .Open
  End With
  
  ' Verifico si existe registros
  If porstMRp.RecordCount > 0 Then
    porstMRp.MoveFirst
    Do While Not porstMRp.EOF
      If porstMRp!NatCta = NATCTA_DEU Then
        nImporte = CDec(porstMRp!nDebe - porstMRp!nHaber)
      Else
        nImporte = CDec(porstMRp!nHaber - porstMRp!nDebe)
      End If
      nImporte = nImporte * IIf(Left(Trim(IIf(IsNull(porstMRp!codcta), "", porstMRp!codcta)), 1) = "9", -1, 1)
      ' Genero la cadena si importe es diferente de cero
      If nImporte <> 0 Then
        ' Inicializo la cadena
        sCadena = "": sSentencia = ""
        ' Obtengo el centro de costos
        sCenCosto = Right(Trim(IIf(IsNull(porstMRp!codcco), "0", porstMRp!codcco)), 1)
        sCenCosto = IIf(IsNumeric(sCenCosto) And sCenCosto <= "4", sCenCosto, "5")
        ' Columna 3 - Columna final
        For nContador = 1 To nColFinal
          sRegistro = Choose(CInt(porstMRp("coldet" & nContador)) + 1, "", "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", Format(nImporte, "#0.00"))
          If sRegistro = "FUN_CCO" Then
            sRegistro = Choose(CInt(sCenCosto) + 1, "", "FUN_MGT", "FUN_ADM", "FUN_MIS", "FUN_MKT", "FUN_CCO")
          End If
          sCadena = sCadena & sRegistro & IIf(nColFinal = nContador, "", sCaracter)
          sSentencia = sSentencia & IIf(nColFinal = nContador, "", sRegistro)
        Next nContador
        ' Inicializo la sentencia
        sSentencia = IIf(sSentencia <> "", "X", "")
        ' Columna 2
        sRegistro = Trim(IIf(IsNull(porstMRp!codcco), "", porstMRp!codcco))
        sRegistro = IIf(sRegistro <> "", Val(sRegistro), "")
        sRegistro = sRegistro & IIf(sRegistro <> "", sSentencia, "")
        sCadena = sRegistro & sCaracter & sCadena
        ' Columna 1
        sRegistro = Trim(IIf(IsNull(porstMRp!codcta), "", porstMRp!codcta))
        sCadena = sRegistro & sCaracter & sCadena
        Print #1, sCadena
      End If
      porstMRp.MoveNext
    Loop
  End If
  ' Elimino los registros
  pocnnMain.Execute "DELETE FROM COTmpRpt WHERE codemp='" & gsCodEmp & "' AND UsrCre='" & gsCodUsr & "' AND nomrpt='rptpfil'"
  Close #1
  porstMRp.Close
  lblProgreso.Caption = ""
  MsgBox TEXT_8008, vbInformation
  ppHabilitacion True

End Sub

']
Private Sub Form_Load()
   
  '[ Verifico exitencia año anterior
  psAnoOld = Trim$(Val(gsAnoAct) - 1)
  ']
   
   On Error GoTo Err
  
 '[Recordsets.                         'Cambiar.
    Set pocnnMain = New ADODB.Connection
    Set porstMRp = New ADODB.Recordset
    Set porstCoFil = New ADODB.Recordset
    Set porstMRpRs = New ADODB.Recordset
    Set porstCoCfg = New ADODB.Recordset
    
    With pocnnMain
        .CursorLocation = adUseClient
        .ConnectionString = CONNSTRG & gsNomBDS
        .Open
    End With
    
    'Obtener simbolo de Moneda
    With porstCoCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT TpoMon_Sgn_MN, TpoMon_Sgn_ME "
      .Source = .Source & "FROM COCfg "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      gsTpoMon_Sgn_MN = .Fields(0)
      gsTpoMon_Sgn_ME = .Fields(1)
      .Close
    End With
    Set porstCoCfg.ActiveConnection = Nothing
    Set porstCoCfg = Nothing
    
    With porstMRp
        .ActiveConnection = pocnnMain
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
    End With
    With porstCoFil
        .ActiveConnection = pocnnMain
        .Source = "SELECT codfil, " & Choose(gsIdioma, "detfil", "detfilx") & " AS detfil, sepcar "
        .Source = .Source & "FROM cofil "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "ORDER BY codfil"
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
    End With
    With porstMRpRs
        .ActiveConnection = pocnnMain
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Source = "SELECT * "
        .Source = .Source & "FROM " & IIf(ps_Plataforma = pSrvSql, "#", "") & "trptREstFin "
        .Source = .Source & "ORDER BY NroLin"
'        .Open
    End With
 ']

 '[Parámetros.                         'Cambiar.
 ']
   
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND
    
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
 
 '[Datos predeterminados.              'Cambiar.
  chkImpFecha.Value = Checked
  lblProgreso.Caption = ""
  'Otros.
   
  'Características de impresión.
   udFecha = Date                      'Fecha en el encabezado.
'   unCopias = rptMain.CopiesToPrinter  'Cantidad de Copias.
   unMargenIzquierdo = 240             'Margen izquierdo.
   usDEstino = PRN_DEST_MATR           'PRN_DEST_GRAF:ica _
                                        PRN_DEST_MATR:icial.
   usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                        PRN_ORIE_HORI:zontal.
 ']
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = porstCoFil
   
   frmOPrnCfg.OrientacionPrn 0, Me
   frmOPrnCfg.lblOriPrn.Caption = Printer.Orientation
      
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub Form_Activate()
  'Orden: Vista Previa, Imprimir, Exportar.
  zaOpciones = Array(gbPms04, gbPms05, gbPms06)
  ppDatosGrid
   
End Sub
Private Sub Form_Resize()
   On Error Resume Next
   
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Cambiar. Añadir recordsets.
   porstCoFil.Close
   pocnnMain.Close
   Set porstCoFil = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
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

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
  cmdImprimir(0).Enabled = tbHabilitar
  cmdImprimir(1).Enabled = tbHabilitar
  cmdConfig.Enabled = tbHabilitar
  cmdExporta.Enabled = tbHabilitar
  cmdSalir.Enabled = tbHabilitar
End Sub

Public Property Get zaOpciones() As Variant
End Property

Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

Private Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   dgrMain.Caption = Choose(gsIdioma, "Archivos de Información", "Files Information")
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Código", "Code")
            .Item(dnNum).Width = 1000
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Descripción", "Description")
            .Item(dnNum).Width = 4680
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

']
']
