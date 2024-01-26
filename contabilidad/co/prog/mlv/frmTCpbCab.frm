VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTCpbCab 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9330
   StartUpPosition =   1  'CenterOwner
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
      Height          =   300
      Index           =   1
      Left            =   1020
      TabIndex        =   10
      Top             =   1410
      Width           =   6400
   End
   Begin VB.TextBox txtLlave 
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
      Height          =   300
      Index           =   1
      Left            =   7200
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   4980
      Picture         =   "frmTCpbCab.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   300
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
      Height          =   300
      Index           =   0
      Left            =   1020
      TabIndex        =   8
      Top             =   1050
      Width           =   6400
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   9330
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5820
      Width           =   9330
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Preliminar"
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
         Index           =   0
         Left            =   5520
         Picture         =   "frmTCpbCab.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   720
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
         Left            =   0
         Picture         =   "frmTCpbCab.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   278
         Width           =   360
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
         Left            =   0
         Picture         =   "frmTCpbCab.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   360
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
         Left            =   1860
         Picture         =   "frmTCpbCab.frx":0A30
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
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
         Left            =   1140
         Picture         =   "frmTCpbCab.frx":0B32
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   720
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
         Left            =   420
         Picture         =   "frmTCpbCab.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Re&frescar"
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
         Left            =   4800
         Picture         =   "frmTCpbCab.frx":0D7E
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
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
         Height          =   560
         Left            =   8470
         Picture         =   "frmTCpbCab.frx":0EC8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   2640
         Picture         =   "frmTCpbCab.frx":1012
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
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
         Left            =   4080
         Picture         =   "frmTCpbCab.frx":1114
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   720
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
         Height          =   560
         Index           =   1
         Left            =   6240
         Picture         =   "frmTCpbCab.frx":1216
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdRevisar 
         Caption         =   "&Revisar"
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
         Left            =   3360
         Picture         =   "frmTCpbCab.frx":1318
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calc&ular"
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
         Left            =   7035
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   7830
         X2              =   7830
         Y1              =   0
         Y2              =   550
      End
   End
   Begin VB.TextBox txtLlave 
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
      Height          =   300
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   520
   End
   Begin MSDataGridLib.DataGrid dgrDetalle 
      Height          =   2595
      Left            =   0
      TabIndex        =   11
      Top             =   1830
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4577
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   -2147483630
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
   Begin VB.Frame Frame1 
      ForeColor       =   &H80000002&
      Height          =   1095
      Left            =   0
      TabIndex        =   30
      Top             =   4680
      Width           =   9280
      Begin VB.TextBox txtDeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
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
         Left            =   7440
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox txtDeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
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
         Left            =   5640
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtDeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
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
         Left            =   7440
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtDeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
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
         Left            =   5640
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Totales M.E. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   6
         Left            =   4440
         TabIndex        =   33
         Top             =   645
         Width           =   960
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Totales M.N. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   5
         Left            =   4440
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpFehCpb 
      Height          =   300
      Left            =   1020
      TabIndex        =   6
      Top             =   690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      Format          =   66519041
      CurrentDate     =   37953
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducción :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   4
      Left            =   105
      TabIndex        =   9
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "NºComprobante:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   165
      Width           =   1185
   End
   Begin VB.Label lblLlaveDeta 
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
      Left            =   1305
      TabIndex        =   2
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   3
      Left            =   105
      TabIndex        =   7
      Top             =   1110
      Width           =   510
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   2
      Left            =   105
      TabIndex        =   5
      Top             =   750
      Width           =   540
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Diario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   9240
      Y1              =   570
      Y2              =   570
   End
End
Attribute VB_Name = "frmTCpbCab"
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
Private porstMRp As ADODB.Recordset
Private pbLoad As Boolean
Private pbNuevo As Boolean
Private pbValidada As Boolean
Private pnColumnaOrd As Integer
Private pbGraba As Boolean

Private Sub cmdImprimir_Click(Index As Integer)
  
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

' Obtengo la información del comprobante
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
    .Source = .Source & "AND a.CodDro = '" & txtLlave(0).Text & "' AND a.NroCpb = '" & txtLlave(1).Text & "' "
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
  If MsgBox(Choose(gsIdioma, " Imprimir Cheque Voucher?", "Print Cheque Voucher"), vbQuestion + vbYesNo + vbDefaultButton1, "Consulta") = vbYes Then
    sReporte = "rptECheVou"
    sTipo = "V"
    If Not dbHay104 Then
      MsgBox Choose(gsIdioma, "El comprobante no tiene alguna cuenta 104.", "The voucher doesn't have any account 104."), vbInformation
      porstMRp.Close
      Set porstMRp = Nothing
      Exit Sub
    End If
  End If
  If Index = 0 Then
    sReporte = "rptECheVou"
    ' Genero el reporte
    gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "LISTADO DE COMPROBANTES", "LISTING OF VOUCHERS"), udFecha, True, True, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & sReporte & ".rpt"
      '         .WindowShowGroupTree = True
      '[ Formulas adicionales
      .Formulas(6) = "tGirado='" & dsGirado & "'"
      .Formulas(7) = "tFecha='" & dsFecha & "'"
      .Formulas(8) = "tImporteNumeros='" & dsImporteNumeros & "'"
      .Formulas(9) = "tImporteLetras='" & dsImporteLetras & "'"
      .Formulas(10) = "cUsuario='" & gfEnmasc(IIf(IsNull(porstMRp!UsrMdf), porstMRp!UsrCre, porstMRp!UsrMdf)) & "'"
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

Private Sub Form_Load()
   pbLoad = True
   pbValidada = False

   Me.KeyPreview = True
   
   With frmTCpbGrd                     'Cambiar Formulario de Grid.
    '[Llaves.                          'Cambiar
      txtLlave(0).MaxLength = .uorstMain_0!coddro.DefinedSize
      txtLlave(1).MaxLength = .uorstMain_0!NroCpb.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
   
      txtDato(gsIdioma - 1).MaxLength = .uorstMain_0!glocpb.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain_0!glocpbx.DefinedSize
    ']
   
      dgrDetalle.MarqueeStyle = dbgHighlightRow
      Set dgrDetalle.DataSource = .uorstMain_1
    '[Propio del Formulario.
      With dtpFehCpb
        .MinDate = CDate("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct)
        .MaxDate = gfUltDia(.MinDate)
        .Value = .MaxDate
      End With
      txtDeta(0).Text = Format(0, FORMATO_NUM_1)
      txtDeta(2).Text = Format(0, FORMATO_NUM_1)
      txtDeta(1).Text = Format(0, FORMATO_NUM_1)
      txtDeta(3).Text = Format(0, FORMATO_NUM_1)
      txtLlave(1).Enabled = False
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(7, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diario : ", "Nº Comprobante : ", "Fecha : ", "Glosa : ", "Traducción : ", "Totales MN : ", "Totales ME : ")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journal : ", "Nº Voucher : ", "Date : ", "Gloss : ", "Translation : ", "Totals NC : ", "Totals FC : ")
  Next nElemento
  cmdCalcular.Caption = Choose(gsIdioma, "Calc&ular", "Calc&ulate")
  CaptionBotones Me, False, False, True, True, True, True, True, True, False, True, True, True, True, aLabel
  ']
   
End Sub

Private Sub Form_Activate()
'   If pbLoad Then
    '[Busca detalle de códigos.           'Cambiar (habilitar/deshabilitar).
'      If txtDato(0).Text <> "" Then ppAyuDet AYUDAT, 0
'      If txtDato(3).Text <> "" Then ppAyuDet AYUDAT, 3
    ']
    
    '[Propio del Formulario.
'      txtDato(5).Tag = frmTFacGrd.uorstTGCli!IndEvn
'      ppCambiaClienteEventual
    ']
'   End If

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
   pbGraba = True
   
   ppbotones_tpognr0 (pbGraba)
   
   pbGraba = False
   
   If pbLoad And Not pbNuevo Then
      pbLoad = False
      frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
      frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' "
      frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.MesPvs='" & gsMesAct & "' "
      frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.CodDro='" & frmTCpbGrd.uorstMain_0!coddro & "' "
      frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "' "
      ppDatosWhere
   End If
   upDatosGrid


'[Propio del Formulario.
  If Month(dtpFehCpb.Value) <> Val(gfMesAct(gsMesAct)) Or Year(dtpFehCpb.Value) <> Val(gsAnoAct) Then
    cmdCorregir.Enabled = False
    cmdNuevo.Enabled = False
    cmdEliminar.Enabled = False
  End If
  If txtLlave(0).Text <> "" Then Call ppAyuDet("L", 0)
']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If frmTCpbGrd.uorstMain_0.RecordCount <> 0 Then
'      frmTCpbGrd.uorstMain_0.CancelUpdate   'Cambiar Formulario de Grid.
   End If
End Sub

Private Sub cmdRetroceder_Click()
'   gpTUe_Retroceder frmTFacGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
   If txtDeta(0).Text <> txtDeta(1).Text Then
      MsgBox TEXT_9011, vbExclamation
      ' Realizo el cuadre del comprobante de diario
      If frmTCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then GoTo Finalizar
      If MsgBox(Choose(gsIdioma, "Desea Forzar el Cuadre?", "Do you want to tally with necessarily ?"), vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   If txtDeta(2).Text <> txtDeta(3).Text Then
      MsgBox TEXT_9012, vbExclamation
      ' Realizo el cuadre del comprobante de diario
      If frmTCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then GoTo Finalizar
      If MsgBox(Choose(gsIdioma, "Desea Forzar el Cuadre?", "Do you want to tally with necessarily ?"), vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
Finalizar:
   gpTUe_Retroceder frmTCpbGrd.uorstMain_0, Me
   frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.MesPvs='" & gsMesAct & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.CodDro='" & frmTCpbGrd.uorstMain_0!coddro & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "' "
   ppDatosWhere
   ppbotones_tpognr0 (pbGraba)
End Sub

Private Sub cmdAvanzar_Click()
'   gpTUe_Avanzar frmTFacGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
   If txtDeta(0).Text <> txtDeta(1).Text Then
      MsgBox TEXT_9011, vbExclamation
      ' Realizo el cuadre del comprobante de diario
      If frmTCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then GoTo Finalizar
      If MsgBox(Choose(gsIdioma, "Desea Forzar el Cuadre?", "Do you want to tally with necessarily ?"), vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   If txtDeta(2).Text <> txtDeta(3).Text Then
      MsgBox TEXT_9012, vbExclamation
      ' Realizo el cuadre del comprobante de diario
      If frmTCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then GoTo Finalizar
      If MsgBox(Choose(gsIdioma, "Desea Forzar el Cuadre?", "Do you want to tally with necessarily ?"), vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If

Finalizar:
   gpTUe_Avanzar frmTCpbGrd.uorstMain_0, Me
   frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.MesPvs='" & gsMesAct & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.CodDro='" & frmTCpbGrd.uorstMain_0!coddro & "' "
   frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "' "
   ppDatosWhere
   ppbotones_tpognr0 (pbGraba)
End Sub

Public Sub cmdCorregir_Click()
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   pbNuevo = False
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
   On Error GoTo Err

   With frmTCpbGrd                     'Cambiar Formulario de Grid.
      '[ No Pertenece al Formulario - Agregado por Angel
      If pbNuevo Then
         pbGraba = True
         .uocnnMain.BeginTrans
         ' MA para conservar el numero digitado por el usuario
         .uorstCODro.Fields("Cpb" & gsMesAct).Value = txtLlave(1).Text
         .uorstCODro.Update
         .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      End If
      ']
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain_0.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain_0
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
            !FyHMdf = Now
         End If
         .Update
      End With
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      If pbNuevo Then
'      If MsgBox(TEXT_1022, vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
'      ' [Formato de emisión.           'Cambiar.
'         Emision ("{VTFacDet.SerDoc} = '" & txtLlave(0).Text & "' AND {VTFacDet.NroDoc} = '" & txtLlave(1).Text & "'")
'      ' ]
         .uorstMain_0.Requery
'         .uorstMain_0.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
          .uorstMain_0.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
'         .uorstMain_0.Find "CodDro='" & txtLlave(0).Text & "' And NroCpb='" & txtLlave(1).Text & "'"
       ']
         cmdGrabar.Enabled = False
         upHabilitacion False
         txtLlave(1).Enabled = False
   
       '[ No Pertenece al Formulario - Agregado por Angel
         cmdNuevo_Click
         cmdRetroceder.Enabled = True
         cmdAvanzar.Enabled = True
         cmdCorregir.Enabled = True
         cmdGrabar.Enabled = False
         cmdDeshacer.Enabled = False
         cmdNuevo.Enabled = True
         cmdEliminar.Enabled = True
         cmdRevisar.Enabled = True
         frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
         frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' "
         frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.MesPvs='" & gsMesAct & "' "
         frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.CodDro='" & txtLlave(0).Text & "' "
         frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.NroCpb='" & txtLlave(1).Text & "' "
         ppDatosWhere
       ']
         
'         upDatosPredeterminados
       '[Llave con el foco al añadir.  'Cambiar.
'         txtLlave(0).SetFocus
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
   
'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
fEstMayUpd
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
   
   cmdNuevo.SetFocus
   Exit Sub
Err:
   gpErrores
  
   frmTCpbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Public Sub cmdNuevo_Click()
' '[Propio del formulario.
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   With frmTCpbGrd.uorstMain_0
      If .RecordCount > 0 Then
         .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
         If .EOF Then
            MsgBox Choose(gsIdioma, "Tiene que Grabar la Cabecera del Diario para poder Registrar el Detalle", "You have to save Journal Header to register Detail"), vbInformation
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
            Exit Sub
         End If
      End If
   End With
' ']
   
'''   gpTUg_Nuevo Me, frmTFacDet          'Cambiar Formulario de Datos.
'gpTVd_Nuevo Me, frmTFacDet          'Cambiar Formulario de Datos.
   gpTVd_Nuevo Me, frmTCpbDet
   '[Agregado por Angel
  frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.MesPvs='" & gsMesAct & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.CodDro='" & txtLlave(0).Text & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.NroCpb='" & txtLlave(1).Text & "' "
  ppDatosWhere
  If frmTCpbGrd.uorstMain_1.RecordCount > 0 Then
    frmTCpbGrd.uorstMain_1.MoveLast
  End If
   ']
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
Dim dvRegistro As Variant
  'Verificación de ítemes creados.     'Cambiar Formulario de Grid.
   If frmTCpbGrd.uorstMain_1.RecordCount = 0 Then
'      MsgBox TEXT_8001, vbInformation
      MsgBox TEXT_8001, vbInformation
   Exit Sub
   End If
   
   'With frmTFacDet                     'Cambiar Formulario de Datos.
   With frmTCpbDet
      .zbNuevo = False
      .upDatosDesconectados 1
      .Caption = TEXT_MODIF & " " & Me.Caption
      
      .Show vbModal
   End With
   '[Agregado por Angel
   dvRegistro = frmTCpbGrd.uorstMain_1.Bookmark
  frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.MesPvs='" & gsMesAct & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.CodDro='" & txtLlave(0).Text & "' "
  frmTCpbGrd.usConnStrgWher_1 = frmTCpbGrd.usConnStrgWher_1 & "AND COCpbDet.NroCpb='" & txtLlave(1).Text & "' "
   ppDatosWhere
   'modificacion 14/09/2009
   'frmTCpbGrd.uorstMain_1.Bookmark = dvRegistro
   ']
'''   dgrMain.SetFocus
dgrDetalle.SetFocus
  
  Exit Sub
Err:
   gpErrores
End Sub

Public Sub cmdEliminar_Click()
   On Error GoTo Err

Dim dnBlqIte As Integer
Dim dvRegistro As Variant

   With frmTCpbGrd                     'Cambiar Formulario de Grid.
   
    'ini 2016-05-27/28 nivel=asisten no elimin datos
       If gsNvlUsr = NVLUSR_ASIS Then
          MsgBox TEXT_9026, vbCritical
          Exit Sub
       End If
    'fin 2016-05-27/28 nivel=asisten no elimin datos
   
     'Verificaciones.
      If gbCieCpb Then                 'Mes Cerrado.
         MsgBox TEXT_9016, vbCritical
         Exit Sub
      End If
      If .uorstMain_0!IndAnu = INDANU_VER Then
         MsgBox TEXT_8009, vbInformation
         Exit Sub
      ElseIf .uorstMain_1.RecordCount = 0 Then
         MsgBox TEXT_8001, vbInformation
         Exit Sub
      ElseIf .uorstMain_1.BOF Then
         .uorstMain_1.MoveNext
      ElseIf .uorstMain_1.EOF Then
         .uorstMain_1.MovePrevious
      End If
      If frmTCpbGrd.uorstMain_1!tpognr <> TPOGNR_DRO Then
         MsgBox Choose(gsIdioma, "No se Puede Eliminar este Item", "This Item can not be eliminated"), vbInformation
         Exit Sub
      End If
     'Confirmación                     'Cambiar.
      If MsgBox(TEXT_1021 & " " & Trim(dgrDetalle.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
         If .uorstMain_1!TpoCtb = TPOCTB_DEB Then
            frmTCpbCab.txtDeta(0) = Format(CDec(frmTCpbCab.txtDeta(0).Text) - .uorstMain_1!ImpMN, FORMATO_NUM_1)
            frmTCpbCab.txtDeta(2) = Format(CDec(frmTCpbCab.txtDeta(2).Text) - .uorstMain_1!ImpME, FORMATO_NUM_1)
         Else
            frmTCpbCab.txtDeta(1) = Format(CDec(frmTCpbCab.txtDeta(1).Text) - .uorstMain_1!ImpMN, FORMATO_NUM_1)
            frmTCpbCab.txtDeta(3) = Format(CDec(frmTCpbCab.txtDeta(3).Text) - .uorstMain_1!ImpME, FORMATO_NUM_1)
         End If
         .uocnnMain.BeginTrans
         With .uorstMain_1
            dvRegistro = .Bookmark
            If Not .RecordCount = 0 Then
               dnBlqIte = !blqite
               .MoveFirst
               Do
                  If !blqite = dnBlqIte Then
                     .Delete
                  End If
                  .MoveNext
               Loop Until .EOF
            End If
            If .RecordCount > 0 Then
               If dvRegistro > .RecordCount Then
                  .MoveLast
               Else
                  .Bookmark = dvRegistro
               End If
            End If
         End With
         .uocnnMain.CommitTrans
      End If
      dgrDetalle.SetFocus
   End With

   Exit Sub
Err:
   gpErrores
   
   frmTCpbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
  '[ Datos No Pertenecen al Formulario - Agregado por Angel
  frmTCpbGrd.uorstMain_1.Requery
  frmTCpbCab.upDatosGrid
  frmTCpbCab.dgrDetalle.SetFocus
  ']
End Sub

Public Sub cmdCalcular_Click()
   With frmTCpbGrd.uorstMain_1
      txtDeta(0).Text = Format(0, FORMATO_NUM_1)
      txtDeta(1).Text = Format(0, FORMATO_NUM_1)
      txtDeta(2).Text = Format(0, FORMATO_NUM_1)
      txtDeta(3).Text = Format(0, FORMATO_NUM_1)
      If frmTCpbGrd.uorstMain_1.RecordCount > 0 Then
         frmTCpbGrd.uorstMain_1.MoveFirst
         Do
            txtDeta(0).Text = Format(CDec(txtDeta(0).Text) + !cImpMN_Deb, FORMATO_NUM_1)
            txtDeta(2).Text = Format(CDec(txtDeta(2).Text) + !cImpME_Deb, FORMATO_NUM_1)
            txtDeta(1).Text = Format(CDec(txtDeta(1).Text) + !cImpMN_Hab, FORMATO_NUM_1)
            txtDeta(3).Text = Format(CDec(txtDeta(3).Text) + !cImpME_Hab, FORMATO_NUM_1)
            frmTCpbGrd.uorstMain_1.MoveNext
         Loop Until .EOF
      End If
   End With
   Set dgrDetalle.DataSource = frmTCpbGrd.uorstMain_1
   upDatosGrid
End Sub

Private Sub cmdSalir_Click()
  
  If CDec(txtDeta(0).Text) <> CDec(txtDeta(1).Text) Then
    ' Realizo el cuadre del comprobante de diario
    If frmTCpbGrd.uorstMain_0!tpognr = TPOGNR_DRO Then
      If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(CDec(txtDeta(0).Text) - CDec(txtDeta(1).Text)))) & "." & Chr(13) & Chr(13) & Choose(gsIdioma, "Desea Forzar el Cuadre?", " Do you want to force the tally with?"), vbYesNo) = 6 Then
        ppCuadreff
      Else
        cmdCorregir.Enabled = IIf(cmdCorregir.Enabled, True, True)
        cmdCorregir.SetFocus
        Exit Sub
      End If
    Else
      MsgBox TEXT_9011, vbExclamation
    End If
  End If
  If txtDeta(2).Text <> txtDeta(3).Text Then
    ' Realizo el cuadre del comprobante de diario
    If frmTCpbGrd.uorstMain_0!tpognr = TPOGNR_DRO Then
      If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(CDec(txtDeta(2).Text) - CDec(txtDeta(3).Text)))) & "." & Chr(13) & Chr(13) & Choose(gsIdioma, "Desea Forzar el Cuadre?", "Do you want to force the tally with?"), vbYesNo) = 6 Then
        ppCuadreff
      Else
        cmdCorregir.Enabled = IIf(cmdCorregir.Enabled, True, True)
        cmdCorregir.SetFocus
        Exit Sub
      End If
    Else
      MsgBox TEXT_9012, vbExclamation
    End If
  End If
  Unload Me

End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtLlave(Index).SetFocus
'   Case 2, 3
'      mskLlave(Index).SetFocus
   End Select
   ppAyuBus AYULLA, Index
End Sub

'Private Sub cmdDatoAyud_Click(Index As Integer)
'   Select Case Index                   'Cambiar. Añadir índices.
'   Case 0, 3
'      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
'   End Select
'   ppAyuBus AYUDAT, Index
'End Sub

Private Sub txtllave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYULLA, Index
   End If
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then
      If Len(txtLlave(0)) <> 4 Then
         txtLlave(0).Enabled = True
         txtLlave(0).SetFocus    'Cambiar.
         Exit Sub
      End If
      txtLlave(0).Enabled = False
      cmdLlaveAyud(0).Enabled = False
      lblLlaveDeta(0).Enabled = False
      
      dtpFehCpb.SetFocus
   End If
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
   Dim dbSalir As Boolean
   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
      Case 1
         If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
            txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
         End If
   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
      Case 0
''         Cancel = ppAyuDet(Index)
         Cancel = ppAyuDet(AYULLA, Index)
         If Cancel Then Exit Sub
         If Len(txtLlave(0)) <> 4 Then
            txtLlave(0).SetFocus
            Exit Sub
         End If
   End Select
 
  'Captura del Siguiente Número.       'Cambiar (Activar/Inactivar).
   If Index = 0 And Len(Trim(txtLlave.Item(0).Text)) <> 0 Then
      If pbNuevo Then
         With frmTCpbGrd.uorstCODro
            If .RecordCount <> 0 Then .MoveFirst
            .Find "CodDro='" & txtLlave(0).Text & "'"
            If .EOF Then
               .AddNew
               !codemp = gsCodEmp
               !coddro = txtLlave(0).Text
            End If
            txtLlave(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtLlave(0).Text)
         End With
      End If
   End If
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0 Then
      With frmTCpbGrd.uorstMain_0
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistro = .Bookmark
            .MoveFirst
             .Find "cLLave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistro <> -1 Then .Bookmark = dvRegistro
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistro
         End If
      End With
      
      cmdGrabar.Enabled = True
''      ppbotones_tpognr0 (pbGraba)
      upHabilitacion True
     '[No pertenece al Formulario - Agregado por Angel
      txtLlave(1).Enabled = True
     ']
      pbValidada = True
   Else
      cmdGrabar.Enabled = False
      upHabilitacion False
      pbValidada = False
   End If
      
   Exit Sub
Err:
   gpErrores
End Sub

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

 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYUDAT, Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 0, 3                           'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

  'Asigna 0 a campos numéricos si están vacíos.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select

  'Busca el dato en su tabla principal.
'   Select Case Index
'   Case 0, 3                           'Cambiar (añadir índices).
'      Cancel = ppAyuDet(AYUDAT, Index)
'      If Cancel Then Exit Sub
'   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrDetalle_HeadClick(ByVal ColIndex As Integer)
'   On Error GoTo Err
   
'   pnColumnaOrd = ColIndex
'   fraBuscar.Caption = TEXT_BUSCA & dgrDetalle.Columns(pnColumnaOrd).Caption
'   txtBuscar = ""

'   psConnStrgOrde = "ORDER BY "
'   Select Case pnColumnaOrd            'Cambiar.
'   Case 1, 2, 3
'      usConnStrgOrde_1 = usConnStrgOrde_1 & "1, 2, 3"
'   Case Else
'      usConnStrgOrde_1 = usConnStrgOrde_1 & pnColumnaOrd + 1
''      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
'   End Select
''   With uorstDetalle
'   With uorstMain_1
'      .Close
''      .Source = psConnStrgSele & psConnStrgWher & psConnStrgOrde
'      .Source = usConnStrgSele_1 & usConnStrgWher_1 & usConnStrgOrde_1
'      .Open
'   End With
''   Set dgrDetalle.DataSource = uorstDetalle
'   Set dgrDetalle.DataSource = uorstMain_1
'   DatosGrid

'   Exit Sub
'Err:
'  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub dgrDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
   If frmTCpbGrd.uorstMain_1.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      frmTCpbGrd.uorstMain_1.MoveFirst
   Case vbKeyEnd
      frmTCpbGrd.uorstMain_1.MoveLast
   End Select
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
'   Else
'      Select Case tnIndex
'      Case 0                              'Cambiar (añadir índices).
'         modAyuBus.Cli_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
'         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'      End Select
   End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex                 'Cambiar.
      Case 0
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTCpbGrd.uorstCODro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
            End If
         End With
      End Select
   Else
'      Select Case tnIndex              'Cambiar.
'      Case 2
'         If txtDato(tnIndex).Text = "" Then
'            lblDatoDeta(tnIndex).Caption = ""
'            Exit Function
'         End If
'         With porstTGDtt
'            If .RecordCount > 0 Then .MoveFirst
'            .Find "CodDtt='" & txtDato(tnIndex).Text & "'"
'            If .EOF Then
'               MsgBox TEXT_8006, vbExclamation
'               ppAyuDet = True
'            Else
'               lblDatoDeta(tnIndex).Caption = " " & !DetDtt
'            End If
'         End With
'      End Select
   End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
'[Propio del formulario.
   Static svValores As Variant
']
   
   On Error GoTo Err

   With frmTCpbGrd                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain_0!codemp = gsCodEmp
            .uorstMain_0!pdoano = gsAnoAct
            .uorstMain_0!coddro = txtLlave(0).Text
            .uorstMain_0!NroCpb = txtLlave(1).Text
         End If

        'Datos.
         .uorstMain_0!FehCpb = dtpFehCpb.Value
         .uorstMain_0!glocpb = txtDato(gsIdioma - 1).Text
         .uorstMain_0!glocpbx = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
         .uorstMain_0!mespvs = gsMesAct
         .uorstMain_0!tpognr = TPOGNR_DRO
         .uorstMain_0!IndAnu = INDANU_FAL
'[ARREGLAR. Eliminar la función gfRetornaValor y crear, en vez de, un recordset fijo.
         svValores = Val(gfRetornaValor(CONNSTRG & gsNomBDS, "SELECT COUNT(*) AS cRegistro FROM COCpbDet WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND mespvs='" & gsMesAct & "' AND CodDro = '" & txtLlave(0).Text & "' AND NroCpb='" & txtLlave(1).Text & "' AND CodCta = '" & CTAFZD_CTA & "'"))
']ARREGLAR.
         .uorstMain_0!IndNCu = IIf(((txtDeta(0).Text <> txtDeta(1).Text) Or (txtDeta(2).Text <> txtDeta(3).Text) Or (svValores > 0)), INDNCU_VER, INDNCU_FAL)
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain_0!coddro
         txtLlave(1).Text = .uorstMain_0!NroCpb
      
        'Datos.
        'cambio glocpb nulo= "" de los contrario es el igual al campo glosa
         txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain_0!glocpb), "", .uorstMain_0!glocpb)
         txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain_0!glocpbx), "", .uorstMain_0!glocpbx)
         dtpFehCpb.Value = .uorstMain_0!FehCpb
         ppAyuDet AYULLA, 0
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
  txtLlave(0).Text = ""
  txtLlave(1).Text = ""
  
  'Datos.
  dtpFehCpb.Value = IIf(Month(Date) = Val(gsMesAct) And Year(Date) = Val(gsAnoAct), Date, gfUltDia("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct))
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Text = ""
    Next
  End With
  
  'Ayudas.
  lblLlaveDeta(0).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
  dtpFehCpb.Enabled = tbHabilitar
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  
  'Ayudas.
  cmdLlaveAyud(0).Enabled = (pbNuevo)
  lblLlaveDeta(0).Enabled = tbHabilitar
End Sub

'[Código propio del formulario.

Private Sub ppDatosWhere()             'Cambiar.
   frmTCpbGrd.uorstMain_1.Close
   frmTCpbGrd.uorstMain_1.Source = frmTCpbGrd.usConnStrgSele_1 & frmTCpbGrd.usConnStrgWher_1 & frmTCpbGrd.usConnStrgOrde_1
   frmTCpbGrd.uorstMain_1.Open
   frmTCpbGrd.uorstMain_1.Properties("Unique Table").Value = "COCpbDet"

' '  frmTFacGrd.usConnStrgWher_1 = "WHERE a.NroDoc='" & frmTFacGrd.uorstMain_0!NroDoc & "' "
''   frmTCpbGrd.usConnStrgWher_1 = "WHERE CodDro='" & frmTCpbGrd.uorstMain_0!CodDro & "' And NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "'"
''   With frmTFacGrd.uorstMain_1
'   With frmTCpbGrd.uorstMain_1
'      .Close
''      .Source = frmTFacGrd.usConnStrgSele_1 & frmTFacGrd.usConnStrgWher_1 & frmTFacGrd.usConnStrgOrde_1
'      .Source = frmTCpbGrd.usConnStrgSele_1 & frmTCpbGrd.usConnStrgWher_1 & frmTCpbGrd.usConnStrgOrde_1
'      .Open
''      .Properties("Unique Table").Value = "VTFacDet"
'      .Properties("Unique Table").Value = "COCpbDet"
'      If .RecordCount <> 0 Then
'         txtDeta(0).Text = Format(0, FORMATO_NUM_1)
'         txtDeta(1).Text = Format(0, FORMATO_NUM_1)
'         txtDeta(2).Text = Format(0, FORMATO_NUM_1)
'         txtDeta(3).Text = Format(0, FORMATO_NUM_1)
'         Do
'            txtDeta(0).Text = Format(txtDeta(0).Text + !cImpMN_Deb, FORMATO_NUM_1)
'            txtDeta(2).Text = Format(txtDeta(2).Text + !cImpME_Deb, FORMATO_NUM_1)
'            txtDeta(1).Text = Format(txtDeta(1).Text + !cImpMN_Hab, FORMATO_NUM_1)
'            txtDeta(3).Text = Format(txtDeta(3).Text + !cImpME_Hab, FORMATO_NUM_1)
'            .MoveNext
'         Loop Until .EOF
'      End If
'
'   End With
'''   With uorstTotales
'''      .Close
'''      .Source = "SELECT SerDoc, NroDoc, PctIGV, TotVVt, TotIGV, TotPVt " _
'''              & "FROM VTFacCab " _
'''              & usConnStrgWher
'''      .Open
'''   End With
''   Set dgrDetalle.DataSource = frmTFacGrd.uorstMain_1
   cmdCalcular_Click
'  Set dgrDetalle.DataSource = frmTCpbGrd.uorstMain_1
'   upDatosGrid
End Sub

Public Sub upDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrDetalle.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = "Item"
            .Item(dnNum).Width = 400
            .Item(dnNum).Alignment = dbgLeft
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
            .Item(dnNum).Width = 850 '70 * (frmTCpbGrd.uorstMain_1.Fields("CodCta").DefinedSize + 1)
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "C.Cto.", "Cost Center")
            .Item(dnNum).Width = 90 * (frmTCpbGrd.uorstMain_1.Fields("CodCCo").DefinedSize + 1)
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 1000 '100 * (frmTCpbGrd.uorstMain_1.Fields("CodAux").DefinedSize + 1)
         Case 4
            .Item(dnNum).Caption = "Doc"
            .Item(dnNum).Width = 100 * (frmTCpbGrd.uorstMain_1.Fields("AbvTDc").DefinedSize + 2)
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Serie", "Series")
            .Item(dnNum).Width = 100 * (frmTCpbGrd.uorstMain_1.Fields("SerDoc").DefinedSize + 2)
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Numero", "Number")
            .Item(dnNum).Width = 90 * (frmTCpbGrd.uorstMain_1.Fields("NroDoc").DefinedSize + 1)
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
            .Item(dnNum).Width = 1100
         Case 8
            .Item(dnNum).Caption = Choose(gsIdioma, "Debe", "Debit")
            .Item(dnNum).Width = 1000 ' * (uorstDetalle.Fields("ImpMN").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 9
            .Item(dnNum).Caption = Choose(gsIdioma, "Haber", "Credit")
            .Item(dnNum).Width = 1000 ' * (uorstDetalle.Fields("ImpME").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 10
            .Item(dnNum).Caption = Choose(gsIdioma, "Tipo", "Type")
            .Item(dnNum).Width = 100 * (frmTCpbGrd.uorstMain_0.Fields("TpoGnr").DefinedSize + 5)
            .Item(dnNum).Alignment = dbgCenter
         Case 11
            .Item(dnNum).Caption = Choose(gsIdioma, "Debe M.E.", "Debit F.C.")
            .Item(dnNum).Width = 1000 ' * (uorstDetalle.Fields("ImpMN").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 12
            .Item(dnNum).Caption = Choose(gsIdioma, "Haber M.E.", "Credit F.C.")
            .Item(dnNum).Width = 1000 ' * (uorstDetalle.Fields("ImpME").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'Private Sub ppCambiaClienteEventual()
'   txtDato(1).Visible = frmTFacGrd.uorstTGCli!IndEvn
'   txtDato(5).Visible = frmTFacGrd.uorstTGCli!IndEvn
'   lblDatoDeta(0).Visible = Not frmTFacGrd.uorstTGCli!IndEvn
'   lblRUCCli.Visible = Not frmTFacGrd.uorstTGCli!IndEvn
'   cmdDatoAyud(0).Left = IIf(frmTFacGrd.uorstTGCli!IndEvn, 8340, 7380)
'End Sub

'Private Function pfCuadre(tnTipo As Byte) As Boolean
'   If txtDeta(0).Text <> txtDeta(1).Text Or txtDeta(2).Text <> txtDeta(3).Text Then
'[REVISAR. Activar primera opción para no permitir comprobantes descuadrados.
'      MsgBox "Los importes no cuadran." & Chr(13) & "No se puede " & IIf(tnTipo = 1, "grabar.", "salir."), vbCritical
'      MsgBox "Los importes no cuadran.", vbCritical
']REVISAR.
'   Else
'      pfCuadre = True
'   End If
'End Function

Private Sub ppCuadreff()
Dim ffImpMN_Deb, ffImpMN_Hab, ffImpME_Deb, ffImpME_Hab As Double
Dim dnNroIte As Integer
   
  With frmTCpbGrd
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      .uorstMain_1.AddNew
      .uorstMain_1!codemp = gsCodEmp
      .uorstMain_1!pdoano = gsAnoAct
      .uorstMain_1!mespvs = gsMesAct
      .uorstMain_1!coddro = txtLlave(0).Text
      .uorstMain_1!NroCpb = txtLlave(1).Text
      .uorstMain_1!NroIte = frmTCpbGrd.pfNumItemCpb(gsAnoAct, gsMesAct, frmTCpbCab.txtLlave(0).Text, frmTCpbCab.txtLlave(1).Text)
      .uorstMain_1!fehope = dtpFehCpb.Value
      .uorstMain_1!CodCta = "FF"
      .uorstMain_1!GloIte = "Item de Cuadre Forzado"
      .uorstMain_1!GloItex = "Item force tally with"
      .uorstMain_1!tpomon = TPOMON_NAC
      .uorstMain_1!feedoc = dtpFehCpb.Value                 ' Necesario fecha por defecto
      .uorstMain_1!fevdoc = dtpFehCpb.Value
      .uorstMain_1!ferdoc = dtpFehCpb.Value
    '[Modificado por Miguel Angel (17/01/2004).
      .uorstMain_1!ImpTCb = Format(1, FORMATO_NUM_2)
      If frmTCpbGrd.uorstTGTCb.RecordCount <> 0 Then
        frmTCpbGrd.uorstTGTCb.MoveFirst
        frmTCpbGrd.uorstTGTCb.Find "FehTCb = '" & .uorstMain_1!feedoc & "'"
        If Not frmTCpbGrd.uorstTGTCb.EOF Then
           .uorstMain_1!ImpTCb = Format(frmTCpbGrd.uorstTGTCb!ImpTCb_Vta, FORMATO_NUM_2)
        End If
      End If
   ']
      .uorstMain_1!tpognr = TPOGNR_DRO
      .uorstMain_1!TpoPvs = TPOPVS_OTR
      ffImpMN_Hab = 0
      ffImpMN_Deb = 0
      ffImpME_Hab = 0
      ffImpME_Deb = 0
      If CDec(txtDeta(0).Text) > CDec(txtDeta(1).Text) Then
         ffImpMN_Hab = gfRedond(Abs(CDec(txtDeta(0).Text) - CDec(txtDeta(1).Text)), 2)
         .uorstMain_1!ImpMN = CDec(IIf(ffImpMN_Hab = "" Or IsNull(ffImpMN_Hab), 0, ffImpMN_Hab))
         .uorstMain_1!TpoCtb = TPOCTB_HAB
      Else
         If CDec(txtDeta(0).Text) < CDec(txtDeta(1).Text) Then
            ffImpMN_Deb = gfRedond(Abs(CDec(txtDeta(1).Text) - CDec(txtDeta(0).Text)), 2)
            .uorstMain_1!ImpMN = CDec(IIf(ffImpMN_Deb = "" Or IsNull(ffImpMN_Deb), 0, ffImpMN_Deb))
            .uorstMain_1!TpoCtb = TPOCTB_DEB
         Else
            .uorstMain_1!ImpMN = 0
         End If
      End If
      If CDec(txtDeta(2).Text) > CDec(txtDeta(3).Text) Then
         ffImpME_Hab = gfRedond(Abs(CDec(txtDeta(2).Text) - CDec(txtDeta(3).Text)), 2)
         .uorstMain_1!ImpME = CDec(IIf(IsNull(ffImpME_Hab), 0, ffImpME_Hab))
         .uorstMain_1!TpoCtb = TPOCTB_HAB
      Else
         If CDec(txtDeta(2).Text) < CDec(txtDeta(3).Text) Then
            ffImpME_Deb = gfRedond(Abs(CDec(txtDeta(3).Text) - CDec(txtDeta(2).Text)), 2)
            .uorstMain_1!ImpME = CDec(IIf(IsNull(ffImpME_Deb), 0, ffImpME_Deb))
            .uorstMain_1!TpoCtb = TPOCTB_DEB
         Else
            .uorstMain_1!ImpME = 0
         End If
      End If
      .uorstMain_1!UsrCre = gsAbvUsr
      .uorstMain_1!FyHCre = Now
      .uorstMain_1.Update
   '''Se coloca que no cuadra, porque estos son asientos de cuadro forzado.
      .uorstMain_0!IndNCu = INDNCU_VER
      .uorstMain_0!UsrMdf = gsAbvUsr
      .uorstMain_0!FyHMdf = Now
      .uorstMain_0.Update
      .uocnnMain.CommitTrans
      dnNroIte = .uorstMain_1!NroIte
      .uorstMain_1.Requery
      frmTCpbCab.upDatosGrid
      If .uorstMain_1.RecordCount <> 0 Then
         .uorstMain_1.MoveFirst
         .uorstMain_1.Find "NroIte=" & dnNroIte
      End If
      txtDeta(0).Text = Format(txtDeta(0).Text + ffImpMN_Deb, FORMATO_NUM_1)
      txtDeta(1).Text = Format(txtDeta(1).Text + ffImpMN_Hab, FORMATO_NUM_1)
      txtDeta(2).Text = Format(txtDeta(2).Text + ffImpME_Deb, FORMATO_NUM_1)
      txtDeta(3).Text = Format(txtDeta(3).Text + ffImpME_Hab, FORMATO_NUM_1)
   End With

End Sub

Private Sub ppbotones_tpognr0(tbgraba As Boolean)
   '[ No pertenece al Formulario - Agregado por Angel
   cmdCorregir.Enabled = True
   cmdNuevo.Enabled = True
   cmdEliminar.Enabled = True
   cmdRefrescar.Enabled = True
   cmdCalcular.Enabled = True
   If Not pbNuevo Then
      If frmTCpbGrd.uorstMain_0.RecordCount > 0 And frmTCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then
         cmdCorregir.Enabled = False
         cmdLlaveAyud(0).Enabled = False
         cmdNuevo.Enabled = False
         cmdEliminar.Enabled = False
         cmdRefrescar.Enabled = False
         cmdCalcular.Enabled = False
      End If
   Else
      cmdNuevo.Enabled = tbgraba
      cmdRevisar.Enabled = tbgraba
      cmdEliminar.Enabled = tbgraba
      cmdRefrescar.Enabled = tbgraba
      cmdImprimir(0).Enabled = tbgraba
      cmdImprimir(1).Enabled = tbgraba
      cmdCalcular.Enabled = tbgraba
      cmdCorregir.Enabled = tbgraba
      cmdLlaveAyud(0).Enabled = IIf(pbNuevo, tbgraba, False)
   End If
   ']
End Sub

Private Sub dtpFehCpb_LostFocus()
''Private Sub dtpFehCpb_Validate(Cancel As Boolean)
   If Month(dtpFehCpb.Value) <> Val(gsMesAct) Or Year(dtpFehCpb.Value) <> Val(gsAnoAct) Then
      If Month(dtpFehCpb.Value) <> Val(gsMesAct) Then
         If Month(dtpFehCpb.Value) = 1 And gsMesAct = "00" Or Month(dtpFehCpb.Value) = 12 And gsMesAct = "13" Then
            Exit Sub
         End If
      End If
      MsgBox Choose(gsIdioma, "La fecha debe ser del Mes y Año que se provisiona.", "The date must correspond with Month and Year that provision."), vbExclamation
      dtpFehCpb.SetFocus
   End If
End Sub

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
   
   'Orden: Corregir.
''   zaOpciones = Array(gbPms02)
zaOpciones = Array(gbPms01, gbPms02, gbPms03, gbPms04, gbPms05)
End Property

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdCorregir.Enabled = IIf(pbNuevo, False, taOpciones(0))
   
cmdNuevo.Enabled = IIf(pbNuevo, False, taOpciones(0))
cmdRevisar.Enabled = IIf(pbNuevo, False, True)
cmdEliminar.Enabled = IIf(pbNuevo, False, taOpciones(2))
cmdRefrescar.Enabled = IIf(pbNuevo, False, True)
cmdImprimir(0).Enabled = IIf(pbNuevo, False, IIf(taOpciones(3) Or taOpciones(4), True, False))
cmdImprimir(1).Enabled = IIf(pbNuevo, False, IIf(taOpciones(3) Or taOpciones(4), True, False))
End Property
