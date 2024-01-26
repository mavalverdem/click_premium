VERSION 5.00
Begin VB.Form frmRLbrCja 
   Caption         =   "[título]"
   ClientHeight    =   3420
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExpota 
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
      Left            =   3480
      Picture         =   "frmRLbrCja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2880
      Width           =   1245
   End
   Begin VB.CheckBox chkEquivalente 
      Caption         =   "Visualiza Equivalente"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   15
      TabIndex        =   9
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   19
      Top             =   2100
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   21
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   20
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   12
      Top             =   45
      Width           =   6975
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
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6570
         Picture         =   "frmRLbrCja.frx":03F0
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   495
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6570
         Picture         =   "frmRLbrCja.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Index           =   1
         Left            =   1050
         TabIndex        =   17
         Top             =   840
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
         Left            =   1080
         TabIndex        =   16
         Top             =   495
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame fraAlcance 
      Caption         =   "Alcance"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   2265
      Begin VB.OptionButton optAlcance 
         Caption         =   "del mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   1065
         TabIndex        =   7
         Top             =   255
         Width           =   1080
      End
      Begin VB.OptionButton optAlcance 
         Caption         =   "al mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   255
         Width           =   1080
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5745
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1725
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
      ScaleWidth      =   6975
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2880
      Width           =   6975
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
         Picture         =   "frmRLbrCja.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1245
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
         Left            =   4800
         Picture         =   "frmRLbrCja.frx":0C76
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmRLbrCja.frx":0DC0
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
      Left            =   5010
      TabIndex        =   18
      Top             =   1770
      Width           =   660
   End
End
Attribute VB_Name = "frmRLbrCja"
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
']
'ini 2016-08-10 exporta excel
Private Sub cmdExpota_Click()
    pExporta 0
End Sub

Private Sub pExporta(TpoRpt As Integer)
    Dim oProgress As New frmzProgressBar
    oProgress.Show
    oProgress.pgbProgreso.Value = 0: oProgress.pgbProgreso.Min = 0
    oProgress.pgbProgreso.Max = 4
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Min
    oProgress.Caption = "Procesando Caja y Bancos"

'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err

    Dim pocnnMain As ADODB.Connection
    Set pocnnMain = New ADODB.Connection
    With pocnnMain
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With

    Dim pocnnTmp As ADODB.Connection
    Set pocnnTmp = New ADODB.Connection
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
    
    Dim cCadReporte  As String
    Dim sTabla As String
    sTabla = "xlsVtaCab"
    

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       .Source = "SELECT * FROM " & ps_Prefijo & sTabla
'ini 2016-08-10 exporta excel

        .Source = "SELECT concat(a.pdoano,a.mespvs," & Chr(34) & "00" & Chr(34) & ") AS DPERIODO,"
        .Source = .Source & "concat(a.CodDro,'-',a.NroCpb,'-',REPEAT('0', (6-LENGTH(a.nroite))), CONVERT(a.nroite, CHAR(5))) as DNUMSIOPE,"
        .Source = .Source & "replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        .Source = .Source & "ifnull(c.DetDro ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
        .Source = .Source & "AS DDETDRO,"
        .Source = .Source & "concat(a.CodDro,'-',a.NroCpb,'-',a.Mespvs) as DCOMMES,"
        
        .Source = .Source & "a.codaux as codaux,"
        .Source = .Source & "e.RazAux as razaux,"
        .Source = .Source & "ifnull(a.codtdc,'') as codtdc,"
        .Source = .Source & "ifnull(a.serdoc,'') as serdoc,"
        .Source = .Source & "ifnull(a.nrodoc,'') as nrodoc,"
        .Source = .Source & "ifnull(a.TPODOC,'') as TPODOC,"
        .Source = .Source & "a.refdoc as refdoc,"
        .Source = .Source & "a.CodCta AS DNUMCTACON,"
        .Source = .Source & "replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        .Source = .Source & "ifnull(b.DetCta ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
        .Source = .Source & "AS DDETCta,"
        .Source = .Source & "ifnull(date_format(a.FehOpe,'%d/%m/%Y'),'') as DFECOPE,"
        .Source = .Source & "ifnull(date_format(a.Feedoc,'%d/%m/%Y'),'') as DFEEDOC,"
        .Source = .Source & "ifnull(date_format(f.Fehope,'%d/%m/%Y'),'') as DFEEDOCREGCONCOMP,"
        .Source = .Source & "concat(f.coddro,'-',f.nrocpb,'-',f.mespvs) as NROCOMPREGCONCOMP,"
        .Source = .Source & "ifnull(date_format(f.Feedoc,'%d/%m/%Y'),'') as DFEEDOCCOMPRAS,"
        .Source = .Source & "ifnull(f.nrocdt,'')as nrocdt,ifnull(date_format(f.fehcdt,'%d/%m/%Y'),'') as fehcdt,"
        .Source = .Source & "ifnull(date_format(g.Fehope,'%d/%m/%Y'),'') as DFEEDOCREGCONHOVO,"
        .Source = .Source & "concat(g.coddro,'-',g.nrocpb,'-',g.mespvs) as NROCOMPREGCONHONO,"
        .Source = .Source & "ifnull(date_format(g.Feedoc,'%d/%m/%Y'),'') as DFEEDOCHONORARIOS,"
        .Source = .Source & "ifnull(date_format(h.Fehope,'%d/%m/%Y'),'') as DFEEDOCREGCONVENT,"
        .Source = .Source & "concat(h.coddro,'-',h.nrocpb,'-',h.mespvs) as NROCOMPREGCOVENTAS,"
        .Source = .Source & "ifnull(date_format(h.Feedoc,'%d/%m/%Y'),'') as DFEEDOCVENTAS,"
        
        
        .Source = .Source & "replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        .Source = .Source & "ifnull(a.gloite ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
        .Source = .Source & "as DGLOSA,"
        .Source = .Source & "replace(FORMAT(IF(a.TpoCtb='D', a.ImpMN, 0),2),',','') AS DDEBE,"
        .Source = .Source & "replace(FORMAT(IF(a.TpoCtb='H', a.ImpMN, 0),2),',','') AS DHABER,"
        .Source = .Source & ""
        .Source = .Source & "replace(FORMAT(IF(a.TpoCtb='D', a.ImpME, 0),2),',','') AS DDEBEME,"
        .Source = .Source & "replace(FORMAT(IF(a.TpoCtb='H', a.ImpME, 0),2),',','') AS DHABERME "
        
        .Source = .Source & "FROM COCpbDet a "
        .Source = .Source & "LEFT JOIN cocta b ON a.codemp=b.codemp  and a.pdoano=b.pdoano and a.CodCta=b.CodCta "
        .Source = .Source & "LEFT JOIN CoDro c ON a.codemp=c.codemp  and a.pdoano=c.pdoano and  CONCAT(LEFT(a.CodDro, 2), '  ')=c.CodDro "
        .Source = .Source & "LEFT JOIN CoDro d ON a.codemp=d.codemp  and a.pdoano=d.pdoano and  a.CodDro=d.CodDro "
        .Source = .Source & "LEFT JOIN TGAux e ON a.codemp=e.codemp and a.CodAux=e.CodAux "
        .Source = .Source & "LEFT JOIN cocta b2 ON a.codemp=b2.codemp and a.pdoano=b2.pdoano and LEFT(a.CodCta, 2)=b2.CodCta "
        .Source = .Source & "LEFT JOIN cocprdoc f ON a.codemp=f.codemp  and a.codaux=f.codaux and a.codtdc=f.codtdc and a.serdoc=f.serdoc and a.nrodoc=f.nrodoc "
        .Source = .Source & "LEFT JOIN cohprdoc g ON a.codemp=g.codemp  and a.codaux=g.codaux and a.serdoc=g.serdoc and a.nrodoc=g.nrodoc "
        .Source = .Source & "LEFT JOIN covtadoc h ON a.codemp=h.codemp  and a.codaux=h.codaux and a.codtdc=h.codtdc and a.serdoc=h.serdoc and a.nrodoc=h.nrodoc "
        '.Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "'  and a.CodDro BETWEEN '00' AND '9999' AND (a.Mespvs >='01' and a.Mespvs <='12') and  NOT(a.tpognr='5' AND a.impMN=0.00)"
        '.Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "'  and a.CodDro BETWEEN '00' AND '9999' AND a.Mespvs <='" & gsMesAct & "' "
        .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "'  and a.CodDro BETWEEN '00' AND '9999' AND a.Mespvs <='" & gsMesAct & "' and  NOT(a.tpognr='5' AND a.impMN=0.00) "
        .Source = .Source & "and mid(a.coddro,1,2) in ('05','06') "
        .Source = .Source & "Having (ddebe + dhaber) <> 0.00 "
        .Source = .Source & "ORDER BY dperiodo, DNUMSIOPE,a.CodCta ASC "
        
'fin 2016-08-10 exporta excel
       .Open
    End With

        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1

    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
 
    'Set oSheet = oWBook.Worksheets(1)
 

    '*Set oExcel = New Excel.Application
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True

    Set oWBook = oExcel.Workbooks.Add
    '*Set oWBook = oExcel.Workbooks.Open(dlbDirectorio(0).path & xArchPeriodo, , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    '*Set oSheet = oWBook.Worksheets("Clientes")
     Set oSheet = oWBook.Worksheets(1)
    '*oExcel.Visible = True

    With oSheet
        oSheet.Select
               
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = "Libro Caja"
        nRowI = nRowI + 2
        Dim x1 As Integer
        Dim x As Integer
        x = 0
        Dim nColumna As Long
        .Cells(nRowI, 1).Value = porstTmp.Fields(0).Name
        For nColumna = 1 To porstTmp.Fields.Count - 1
            .Cells(nRowI, nColumna + 1).Value = porstTmp.Fields(nColumna).Name
        Next nColumna
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
        
'''        x = x + 1: .Cells(nRowI, x).Value = "Empresa"
'''        x = x + 1: .Cells(nRowI, x).Value = "Detalle"

        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'        Columns("L:L").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("M:M").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"

        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'        oSheet.Select
'        Cells(1, 1).Select

    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing

'fin exporta datos a excel

   porstTmp.Close
   pocnnTmp.Close
   Set porstTmp = Nothing
   Set pocnnTmp = Nothing

   pocnnMain.Close
   Set pocnnMain = Nothing

    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
    Unload oProgress          ' Unload progress bar window

  Exit Sub
Err:
    MsgBox (TEXT_6001)
  If pocnnTmp.State = adStateOpen Then
    porstTmp.Close
    pocnnTmp.Close
    Set porstTmp = Nothing
    Set pocnnTmp = Nothing
    
   pocnnMain.Close
   Set pocnnMain = Nothing
    
  End If

End Sub
'fin 2016-08-10 exporta excel


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
      .Source = .Source & "AND LEFT(CodCta, 2)='10' "
      .Source = .Source & "ORDER BY CodCta "
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
  chkEquivalente.Caption = Choose(gsIdioma, "Imprime Equivalente", "Print Equivalent")
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraAlcance.Caption = Choose(gsIdioma, "Alcance", "Scope")
  optAlcance(0).Caption = Choose(gsIdioma, "al mes", "to month")
  optAlcance(1).Caption = Choose(gsIdioma, "del mes", "from month")
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
      txtDato(1).Text = !CodCta
      .MoveFirst
      txtDato(0).Text = !CodCta
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   optAlcance(1).Value = True
   chkEquivalente.Value = vbUnchecked
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
  Dim sMoneda As String, sMonedae As String
  Dim sEquivalente As String
  ppHabilitacion False
  
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sMonedae = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_TXT, TPOMON_NAC_TXT)
  sEquivalente = IIf(chkEquivalente.Value = vbChecked, "S", "N")
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.CodCta, a.CodDro, a.NroCpb, a.FehOpe, a.MesPvs, "
    .Source = .Source & "(CASE a.MesPvs WHEN '1' THEN 'ENERO' WHEN '2' THEN 'FEBRERO' WHEN '3' THEN 'MARZO' WHEN '4' THEN 'ABRIL' WHEN '5' THEN 'MAYO' WHEN '6' THEN 'JUNIO' WHEN '7' THEN 'JULIO' WHEN '8' THEN 'AGOSTO' WHEN '9' THEN 'SETIEMBRE' WHEN '10' THEN 'OCTUBRE' WHEN '11' THEN 'NOVIEMBRE' WHEN '12' THEN 'DICIEMBRE' END) AS Tmes, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    .Source = .Source & "a.CodAux, b.RazAux, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, d.DetDro, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cCargo, "
    .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cAbono, "
    If optAlcance(0).Value = True Then
      .Source = .Source & "(" & gsIniAno(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb, "
      .Source = .Source & "(" & gsIniAno(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab, "
      .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsIniAno(IIf(cboTpoMon.ListIndex = 0, 2, 1)) & ") ELSE 0 END) AS cAntCtaCar, "
      .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsIniAno(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo "
    Else
      .Source = .Source & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb, "
      .Source = .Source & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab, "
      .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 2, 1)) & ") ELSE 0 END) AS cAntCtaCar, "
      .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo "
    End If
    .Source = .Source & "FROM ((((COCpbDet a "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    .Source = .Source & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "AND a.MesPvs " & IIf(optAlcance(0).Value, "<", "") & "='" & gsMesAct & "' "
    .Source = .Source & "ORDER BY a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(optAlcance(0).Value, "rptrLbrcjaalmes.rpt", "rptrlbrcja.rpt")
      ' Parametros adicionales
      .ParameterFields(1) = "Equivalente;" & sEquivalente & ";true"

      '         .WindowShowGroupTree = True
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
      .LoadReport gsRutRpt & IIf(optAlcance(0).Value, "rptRLbrCjaAlMes.mrp", "rptRLbrCja.mrp")
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("Equivalente") = sEquivalente
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
      modAyuBus.Cta_Cod "LEFT(CodCta, 2) = '10'", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetCta), "", !DetCta)
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

    cmdExpota.Enabled = tbHabilitar '2016-08-10 exporta excel
    
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

