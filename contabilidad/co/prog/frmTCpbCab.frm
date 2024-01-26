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
   ScaleHeight     =   6390
   ScaleWidth      =   9330
   StartUpPosition =   1  'CenterOwner
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
      Height          =   315
      Index           =   1
      Left            =   7200
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   4980
      Picture         =   "frmTCpbCab.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   130
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
      Left            =   780
      TabIndex        =   2
      Top             =   1140
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5820
      Width           =   9330
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
         Picture         =   "frmTCpbCab.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTCpbCab.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmTCpbCab.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmTCpbCab.frx":0600
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmTCpbCab.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmTCpbCab.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmTCpbCab.frx":0996
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmTCpbCab.frx":0AE0
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmTCpbCab.frx":0BE2
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   5520
         Picture         =   "frmTCpbCab.frx":0CE4
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmTCpbCab.frx":0DE6
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   7440
         X2              =   7440
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
      Height          =   315
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   520
   End
   Begin MSDataGridLib.DataGrid dgrDetalle 
      Height          =   2595
      Left            =   0
      TabIndex        =   3
      Top             =   1680
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
      TabIndex        =   23
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
         TabIndex        =   19
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label Label5 
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
         Left            =   4440
         TabIndex        =   30
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label7 
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
         Left            =   4440
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpFehCpb 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61669377
      CurrentDate     =   37953
   End
   Begin VB.Label Label2 
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
      Left            =   5760
      TabIndex        =   29
      Top             =   180
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
      Height          =   315
      Index           =   0
      Left            =   1305
      TabIndex        =   27
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label LblMotivo 
      AutoSize        =   -1  'True
      Caption         =   "Glosa:"
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
      Left            =   105
      TabIndex        =   26
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label3 
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
      Left            =   105
      TabIndex        =   25
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label1 
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
      Left            =   105
      TabIndex        =   21
      Top             =   180
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   9240
      Y1              =   600
      Y2              =   600
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

Private Sub Form_Load()
   pbLoad = True
   pbValidada = False

   Me.KeyPreview = True
   
   With frmTCpbGrd                     'Cambiar Formulario de Grid.
    '[Llaves.                          'Cambiar
      txtLlave(0).MaxLength = .uorstMain_0!CodDro.DefinedSize
      txtLlave(1).MaxLength = .uorstMain_0!NroCpb.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
   
      txtDato(0).MaxLength = .uorstMain_0!glocpb.DefinedSize
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
      frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.CodDro='" & frmTCpbGrd.uorstMain_0!CodDro & "' And COCpbDet.NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "' And MesPvs='" & gsMesAct & "'" ' And COCpbDet.CodTDc = TGTDc.CodTDc"
      ppDatosWhere
   End If
   upDatosGrid


'[Propio del Formulario.
'   If Month(dtpFehCpb.Value) <> Val(gsMesAct) Or Year(dtpFehCpb.Value) <> Val(gsAnoAct) Then
   If Month(dtpFehCpb.Value) <> Val(gfMesAct(gsMesAct)) Or Year(dtpFehCpb.Value) <> Val(gsAnoAct) Then
      cmdCorregir.Enabled = False
      cmdNuevo.Enabled = False
      cmdEliminar.Enabled = False
   End If
   If txtLlave(0).Text <> "" Then Call ppAyuDet("L", 0)
'   dtpFehCpb.Value = Date
']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If frmTCpbGrd.uorstMain_0.RecordCount <> 0 Then
'      frmTCpbGrd.uorstMain_0.CancelUpdate   'Cambiar Formulario de Grid.
   End If
'   frmTFacGrd.uorstMain_0.CancelUpdate
End Sub

Private Sub cmdRetroceder_Click()
'   gpTUe_Retroceder frmTFacGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
   If txtDeta(0).Text <> txtDeta(1).Text Then
      MsgBox TEXT_9011, vbExclamation
      If MsgBox("¿Desea Forzar el Cuadre?", vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   If txtDeta(2).Text <> txtDeta(3).Text Then
      MsgBox TEXT_9012, vbExclamation
      If MsgBox("¿Desea Forzar el Cuadre?", vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   gpTUe_Retroceder frmTCpbGrd.uorstMain_0, Me
   frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.CodDro='" & frmTCpbGrd.uorstMain_0!CodDro & "' And COCpbDet.NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "' And MesPvs='" & gsMesAct & "'" ' And COCpbDet.CodTDc = TGTDc.CodTDc "
   ppDatosWhere
   ppbotones_tpognr0 (pbGraba)
End Sub

Private Sub cmdAvanzar_Click()
'   gpTUe_Avanzar frmTFacGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
   If txtDeta(0).Text <> txtDeta(1).Text Then
      MsgBox TEXT_9011, vbExclamation
      If MsgBox("¿Desea Forzar el Cuadre?", vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   If txtDeta(2).Text <> txtDeta(3).Text Then
      MsgBox TEXT_9012, vbExclamation
      If MsgBox("¿Desea Forzar el Cuadre?", vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   gpTUe_Avanzar frmTCpbGrd.uorstMain_0, Me
   frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.CodDro='" & frmTCpbGrd.uorstMain_0!CodDro & "' And COCpbDet.NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "' And MesPvs='" & gsMesAct & "'"  ' And COCpbDet.CodTDc = TGTDc.CodTDc "
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
         If gsMesAct = "00" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb00), "", .uorstCODro!Cpb00), IIf(IsNull(.uorstCODro!Cpb00), 6, Len(.uorstCODro!Cpb00)), 1, "0")
         If gsMesAct = "01" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb01), "", .uorstCODro!Cpb01), IIf(IsNull(.uorstCODro!Cpb01), 6, Len(.uorstCODro!Cpb01)), 1, "0")
         If gsMesAct = "02" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb02), "", .uorstCODro!Cpb02), IIf(IsNull(.uorstCODro!Cpb02), 6, Len(.uorstCODro!Cpb02)), 1, "0")
         If gsMesAct = "03" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb03), "", .uorstCODro!Cpb03), IIf(IsNull(.uorstCODro!Cpb03), 6, Len(.uorstCODro!Cpb03)), 1, "0")
         If gsMesAct = "04" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb04), "", .uorstCODro!Cpb04), IIf(IsNull(.uorstCODro!Cpb04), 6, Len(.uorstCODro!Cpb04)), 1, "0")
         If gsMesAct = "05" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb05), "", .uorstCODro!Cpb05), IIf(IsNull(.uorstCODro!Cpb05), 6, Len(.uorstCODro!Cpb05)), 1, "0")
         If gsMesAct = "06" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb06), "", .uorstCODro!Cpb06), IIf(IsNull(.uorstCODro!Cpb06), 6, Len(.uorstCODro!Cpb06)), 1, "0")
         If gsMesAct = "07" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb07), "", .uorstCODro!Cpb07), IIf(IsNull(.uorstCODro!Cpb07), 6, Len(.uorstCODro!Cpb07)), 1, "0")
         If gsMesAct = "08" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb08), "", .uorstCODro!Cpb08), IIf(IsNull(.uorstCODro!Cpb08), 6, Len(.uorstCODro!Cpb08)), 1, "0")
         If gsMesAct = "09" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb09), "", .uorstCODro!Cpb09), IIf(IsNull(.uorstCODro!Cpb09), 6, Len(.uorstCODro!Cpb09)), 1, "0")
         If gsMesAct = "10" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb10), "", .uorstCODro!Cpb10), IIf(IsNull(.uorstCODro!Cpb10), 6, Len(.uorstCODro!Cpb10)), 1, "0")
         If gsMesAct = "11" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!cpb11), "", .uorstCODro!cpb11), IIf(IsNull(.uorstCODro!cpb11), 6, Len(.uorstCODro!cpb11)), 1, "0")
         If gsMesAct = "12" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb12), "", .uorstCODro!Cpb12), IIf(IsNull(.uorstCODro!Cpb12), 6, Len(.uorstCODro!Cpb12)), 1, "0")
         If gsMesAct = "13" Then txtLlave(1).Text = gfCeros(IIf(IsNull(.uorstCODro!Cpb13), "", .uorstCODro!Cpb13), IIf(IsNull(.uorstCODro!Cpb13), 6, Len(.uorstCODro!Cpb13)), 1, "0")
         If gsMesAct = "00" Then frmTCpbGrd.uorstCODro!Cpb00 = txtLlave(1).Text
         If gsMesAct = "01" Then frmTCpbGrd.uorstCODro!Cpb01 = txtLlave(1).Text
         If gsMesAct = "02" Then frmTCpbGrd.uorstCODro!Cpb02 = txtLlave(1).Text
         If gsMesAct = "03" Then frmTCpbGrd.uorstCODro!Cpb03 = txtLlave(1).Text
         If gsMesAct = "04" Then frmTCpbGrd.uorstCODro!Cpb04 = txtLlave(1).Text
         If gsMesAct = "05" Then frmTCpbGrd.uorstCODro!Cpb05 = txtLlave(1).Text
         If gsMesAct = "06" Then frmTCpbGrd.uorstCODro!Cpb06 = txtLlave(1).Text
         If gsMesAct = "07" Then frmTCpbGrd.uorstCODro!Cpb07 = txtLlave(1).Text
         If gsMesAct = "08" Then frmTCpbGrd.uorstCODro!Cpb08 = txtLlave(1).Text
         If gsMesAct = "09" Then frmTCpbGrd.uorstCODro!Cpb09 = txtLlave(1).Text
         If gsMesAct = "10" Then frmTCpbGrd.uorstCODro!Cpb10 = txtLlave(1).Text
         If gsMesAct = "11" Then frmTCpbGrd.uorstCODro!cpb11 = txtLlave(1).Text
         If gsMesAct = "12" Then frmTCpbGrd.uorstCODro!Cpb12 = txtLlave(1).Text
         If gsMesAct = "13" Then frmTCpbGrd.uorstCODro!Cpb13 = txtLlave(1).Text
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
         frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.CodDro='" & txtLlave(0) & "' And COCpbDet.NroCpb='" & txtLlave(1) & "' And MesPvs='" & gsMesAct & "'"  ' And COCpbDet.CodTDc = TGTDc.CodTDc "
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
   cmdNuevo.SetFocus
   Exit Sub
Err:
   gpErrores
  
'   frmTFacGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
   
'   If uorstMain!NroDoc_Gen <> "" Then
'      MsgBox "No se puede crear ítemes en la factura." & Chr(13) & "Ha sido generada automáticamente.", vbCritical
'      Exit Sub
'   ElseIf uorstDetalle.RecordCount >= MAXITE_FAC Then
'      MsgBox "Máximo puede crear " & MAXITE_FAC & " ítemes.", vbInformation
'      Exit Sub
'   ElseIf uorstMain!IndAnu Then
'      MsgBox TEXT_8009, vbInformation
'      Exit Sub
'   End If
   
   With frmTCpbGrd.uorstMain_0
      If .RecordCount > 0 Then
         .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
         If .EOF Then
            MsgBox "Tiene que Grabar la Cabecera del Diario para poder Registrar el Detalle", vbInformation
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
   frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.CodDro='" & txtLlave(0) & "' And COCpbDet.NroCpb='" & txtLlave(1) & "' And MesPvs='" & gsMesAct & "'"   ' And COCpbDet.CodTDc = TGTDc.CodTDc "
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
   frmTCpbGrd.usConnStrgWher_1 = "WHERE COCpbDet.CodDro='" & txtLlave(0) & "' And COCpbDet.NroCpb='" & txtLlave(1) & "' And MesPvs='" & gsMesAct & "'"  ' And COCpbDet.CodTDc = TGTDc.CodTDc "
   ppDatosWhere
   frmTCpbGrd.uorstMain_1.Bookmark = dvRegistro
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
      If frmTCpbGrd.uorstMain_1!TpoGnr <> TPOGNR_DRO Then
         MsgBox "No se Puede Eliminar este Item", vbInformation
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
               dnBlqIte = !BlqIte
               .MoveFirst
               Do
                  If !BlqIte = dnBlqIte Then
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
'   gpTUg_Refrescar Me
   '[ Datos No Pertenecen al Formulario - Agregado por Angel
   frmTCpbGrd.uorstMain_0.Requery
   frmTCpbGrd.ppDatosGrid
   frmTCpbCab.dgrDetalle.SetFocus
   ']
End Sub

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresión.  'Cambiar.
'   frmEFac.Caption = "Listado de " & Me.Caption
'   frmEFac.Show vbModal
 ']
   
'   Dim dsTpoMon As String
   Dim dsFecha As String, dsGirado As String, dsGirado2 As String, dsImporteNumeros As String, dsImporteLetras As String
   Dim dbHayAux As Boolean, dbHay104 As Boolean

   udFecha = Date                      'Fecha en el encabezado.
   Set porstMRp = New ADODB.Recordset
   With porstMRp
      .ActiveConnection = frmTCpbGrd.uocnnMain
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
    
   Set MRViewer = New MRViewerObject
          
'   With porstMRp
'     'Obtiene el Tipo de Moneda del primer ítem con documento.
'      If .State = adStateOpen Then .Close
'      .Source = "SELECT a.TpoMon " _
'              & "FROM COCpbDet a " _
'              & "WHERE a.CodDro = '" & txtLlave(0).Text & "' AND a.NroCpb = '" & txtLlave(1).Text & "' AND a.MesPvs = '" & gsMesAct & "'" _
'              & "  AND (IFNULL(a.CodAux, '')<>'' AND IFNULL(a.CodTDc, '')<>'' " _
'              & "  AND IFNULL(a.SerDoc, '')<>'' AND IFNULL(a.NroDoc, '')<>'') "
'      .Open
'   End With
'   If porstMRp.RecordCount = 0 Then
'      MsgBox "El comprobante no tiene algún documento registrado.", vbInformation
'   Else
      With porstMRp
'         .MoveFirst
'         dsTpoMon = !TpoMon
         
'         .Close
         .Source = "SELECT c.FehCpb, c.GloCpb, a.ImpTcb, a.TpoPvs, a.TpoMon," _
                 & "  a.MesPvs, a.FehOpe, a.CodCta, a.CodCCo, a.CodAux, d.RazAux, " _
                 & "  CONCAT('(',a.CodDro, '-', a.NroCpb,')') AS cComprobante," _
                 & "  CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc) AS cDocumento," _
                 & "  a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, a.GloIte," _
                 & "  a.ImpME, a.ImpMN," _
                 & "  if(a.TpoCtb='D', a.impME,0) as DebME," _
                 & "  if(a.TpoCtb='H', a.ImpME,0) as HabME," _
                 & "  if(a.TpoCtb='D', a.ImpMN,0) as DebMN," _
                 & "  if(a.TpoCtb='H', a.ImpMN,0) as HabMN " _
                 & "FROM ((COCpbCab c" _
                 & "  LEFT JOIN COCpbDet a ON c.CodDro=a.CodDro and c.NroCpb=a.NroCpb and c.MesPvs = a.MesPvs)" _
                 & "  LEFT JOIN TGTDc b ON a.CodTDc=b.CodTDc) " _
                 & "  LEFT JOIN TGAux d ON a.CodAux=d.CodAux " _
                 & "WHERE a.CodDro = '" & txtLlave(0).Text & "' AND a.NroCpb = '" & txtLlave(1).Text & "' AND a.MesPvs = '" & gsMesAct & "' " _
                 & "ORDER BY a.CodDro, a.NroCpb, a.NroIte"
         .Open

         dsGirado = ""
         dsGirado2 = ""
         dbHayAux = False
         dbHay104 = False
'         dsTpoMon = !TpoMon
         Do While Not .EOF
            'If Not dbHayAux And Not dbHay104 Then dsGirado = !GloIte
            If Len(Trim(!CodAux)) <> 0 And Not dbHayAux Then
               dbHayAux = True
               dsGirado = !RazAux
            End If
           
            If Left(!CodCta, 3) = "104" And Not dbHay104 Then
               dbHay104 = True
               dsFecha = Format(!FehOpe, "d mmmm yyyy")
'               dsImporteNumeros = "********" & Format(CDec(IIf(dsTpoMon = TPOMON_NAC, !ImpMN, !ImpME)), FORMATO_NUM_1)
'               dsImporteLetras = IIf(dsTpoMon = TPOMON_NAC, gfNumLet(!ImpMN, "0"), gfNumLet(!ImpME, "0")) & "********"
               dsImporteNumeros = "********" & Format(CDec(IIf(!TpoMon = TPOMON_NAC, !ImpMN, !ImpME)), FORMATO_NUM_1)
               dsImporteLetras = IIf(!TpoMon = TPOMON_NAC, gfNumLet(!ImpMN, "0"), gfNumLet(!ImpME, "0")) & "********"
               dsGirado2 = !GloIte
            End If
            
            If dbHayAux And dbHay104 Then Exit Do
            .MoveNext
         Loop
         
         .MoveFirst
         If dbHayAux = True Then
            dsGirado = dsGirado & "********"
                       
         Else
            dsGirado = dsGirado2 & "********"
         End If
      End With
      
      
      If Not dbHay104 Then
         MsgBox "El comprobante no tiene alguna cuenta 104.", vbInformation
      Else
         
         With MRViewer
            .DataRecordSet = porstMRp
            .LoadReport gsRutRpt & "rptECheVou.mrp"
            Call gpEncabezadoMRp(MRViewer, "LISTADO DE COMPROBANTES", udFecha, True)
          '[Parámetros adicionales.
            .Parameters("tGirado") = dsGirado
            .Parameters("tFecha") = dsFecha
            .Parameters("tImporteNumeros") = dsImporteNumeros
            .Parameters("tImporteLetras") = dsImporteLetras
          ']
               
            .PreviewReport
            .UnLoadReport
         End With
         Set MRViewer = Nothing
      End If
'   End If
   porstMRp.Close
   Set porstMRp = Nothing
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
      If MsgBox(TEXT_9011 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(CDec(txtDeta(0).Text) - CDec(txtDeta(1).Text)))) & "." & Chr(13) & Chr(13) & "¿Desea Forzar el Cuadre?", vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.Enabled = IIf(cmdCorregir.Enabled, True, True)
         cmdCorregir.SetFocus
         Exit Sub
      End If
   End If
   If txtDeta(2).Text <> txtDeta(3).Text Then
      If MsgBox(TEXT_9012 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(CDec(txtDeta(2).Text) - CDec(txtDeta(3).Text)))) & "." & Chr(13) & Chr(13) & "¿Desea Forzar el Cuadre?", vbYesNo) = 6 Then
         ppCuadreff
      Else
         cmdCorregir.Enabled = IIf(cmdCorregir.Enabled, True, True)
         cmdCorregir.SetFocus
         Exit Sub
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

'Private Sub mskLlave_GotFocus(Index As Integer)
'   mskLlave(Index).SelStart = 0
'   mskLlave(Index).SelLength = mskLlave.Item(Index).MaxLength
'End Sub

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus AYUDAT, Index
'   End If
'End Sub

'Private Sub mskDato_Validate(Index As Integer, Cancel As Boolean)
'   Select Case Index    'Completa con ceros a la izquierda.
'   Case 1                              'Cambiar (añadir índices).
'      If Len(Trim(mskDato(Index).Text)) <> 0 And Len(Trim(mskDato(Index).Text)) <> mskDato(Index).MaxLength Then
'         mskDato(Index) = gfCeros(mskDato(Index).Text, mskDato(Index).MaxLength, 0, "0")
'      End If
'   End Select
'
'   Select Case Index    'Asigna 0 a campos numéricos si están vacíos.
'   Case 2                              'Cambiar (añadir índices).
'      If Not IsNumeric(mskDato(Index).Text) Then
'         mskDato(Index).Text = 0
'      End If
'   End Select
'
'   Select Case Index    'Busca el dato en su tabla principal.
'   Case 1                              'Cambiar (añadir índices).
'      Cancel = ppAyuDet(AYUDAT, Index)
'      If Cancel Then Exit Sub
'   End Select
'End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
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

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
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
''         With porstVTNroDoc
         With frmTCpbGrd.uorstCODro
           
           '[ No pertenece al estandar, se agrego condición para el Diario - Angel
'            If Len(txtLlave(0)) = 2 Then
'               dbSalir = False
'               .Close
'               .Source = "Select CodDro From CoDro Where Left(CodDro,2)='" & txtLlave(0) & "'"
'               .Open
'               If .RecordCount > 1 Then
'                  MsgBox "Registre el Asiento en una Subdivicion Existente del Diario " & txtLlave(0), vbExclamation
'                  dbSalir = True
'               End If
'               .Close
'               .Source = "SELECT CodDro, DetDro, Cpb" & gsMesAct & " " _
                       & "FROM CODro"
'               .Open
'               If dbSalir Then
'                  Exit Sub
'               End If
'            End If
           ']
            
            If .RecordCount <> 0 Then .MoveFirst
''            .Find "cLlave = '" & CODTDC_FAC & txtLlave.Item(0).Text & "'"
            .Find "CodDro='" & txtLlave(0).Text & "'"
            If .EOF Then
               .AddNew
               !CodDro = txtLlave(0).Text
''              !CodTDc = CODTDC_FAC
''               !SerDoc = txtLlave(0).Text
''               !NroDoc = gfCeros("", .Fields("NroDoc").DefinedSize, 0, "0")
            End If
            If gsMesAct = "00" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb00), "", !Cpb00), IIf(IsNull(!Cpb00), 6, Len(!Cpb00)), 1, "0")
            If gsMesAct = "01" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb01), "", !Cpb01), IIf(IsNull(!Cpb01), 6, Len(!Cpb01)), 1, "0")
            If gsMesAct = "02" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb02), "", !Cpb02), IIf(IsNull(!Cpb02), 6, Len(!Cpb02)), 1, "0")
            If gsMesAct = "03" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb03), "", !Cpb03), IIf(IsNull(!Cpb03), 6, Len(!Cpb03)), 1, "0")
            If gsMesAct = "04" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb04), "", !Cpb04), IIf(IsNull(!Cpb04), 6, Len(!Cpb04)), 1, "0")
            If gsMesAct = "05" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb05), "", !Cpb05), IIf(IsNull(!Cpb05), 6, Len(!Cpb05)), 1, "0")
            If gsMesAct = "06" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb06), "", !Cpb06), IIf(IsNull(!Cpb06), 6, Len(!Cpb06)), 1, "0")
            If gsMesAct = "07" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb07), "", !Cpb07), IIf(IsNull(!Cpb07), 6, Len(!Cpb07)), 1, "0")
            If gsMesAct = "08" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb08), "", !Cpb08), IIf(IsNull(!Cpb08), 6, Len(!Cpb08)), 1, "0")
            If gsMesAct = "09" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb09), "", !Cpb09), IIf(IsNull(!Cpb09), 6, Len(!Cpb09)), 1, "0")
            If gsMesAct = "10" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb10), "", !Cpb10), IIf(IsNull(!Cpb10), 6, Len(!Cpb10)), 1, "0")
            If gsMesAct = "11" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!cpb11), "", !cpb11), IIf(IsNull(!cpb11), 6, Len(!cpb11)), 1, "0")
            If gsMesAct = "12" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb12), "", !Cpb12), IIf(IsNull(!Cpb12), 6, Len(!Cpb12)), 1, "0")
            If gsMesAct = "13" Then txtLlave(1).Text = gfCeros(IIf(IsNull(!Cpb13), "", !Cpb13), IIf(IsNull(!Cpb13), 6, Len(!Cpb13)), 1, "0")
''            !NroDoc = txtLlave(1).Text
         End With
''         cmdNuevo.Enabled = True
''         cmdRevisar.Enabled = True
''         cmdEliminar.Enabled = True
''         cmdRefrescar.Enabled = True
''         cmdImprimir.Enabled = True
''         cmdCalcular.Enabled = True
      End If
   End If
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0 Then
'   If Len(Trim(txtLlave(0).Text)) <> 0 Then
      'With frmTFacGrd.uorstMain_0      'Cambiar Formulario de Grid.
      With frmTCpbGrd.uorstMain_0
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistro = .Bookmark
            .MoveFirst
'            .Find "NroDoc='" & txtLlave(0).Text & "'"
'            .Find "CodDro='" & txtLlave(0).Text & "' And NroCpb='" & txtLlave(1).Text & "'"
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
'      cmdLlaveAyud(0).Enabled = False
'      txtLlave(0).Enabled = False
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
         modAyuBus.Dro_Cod "Length(CodDro)=4", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
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
               lblLlaveDeta(tnIndex).Caption = " " & !DetDro
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
            .uorstMain_0!CodDro = txtLlave(0).Text
            .uorstMain_0!NroCpb = txtLlave(1).Text
         End If

        'Datos.
         .uorstMain_0!FehCpb = dtpFehCpb.Value
         .uorstMain_0!glocpb = txtDato(0).Text
         .uorstMain_0!MesPvs = gsMesAct
         .uorstMain_0!TpoGnr = TPOGNR_DRO
         .uorstMain_0!IndAnu = INDANU_FAL
'[ARREGLAR. Eliminar la función gfRetornaValor y crear, en vez de, un recordset fijo.
         svValores = Val(gfRetornaValor(.uocnnMain, "SELECT COUNT(*) AS cRegistro FROM COCpbDet WHERE CodDro = '" & txtLlave(0).Text & "' AND NroCpb='" & txtLlave(1).Text & "' AND MesPvs='" & gsMesAct & "' AND CodCta = '" & CTAFZD_CTA & "'"))
']ARREGLAR.
         .uorstMain_0!IndNCu = IIf(((txtDeta(0).Text <> txtDeta(1).Text) Or (txtDeta(2).Text <> txtDeta(3).Text) Or (svValores > 0)), INDNCU_VER, INDNCU_FAL)
'         .uorstMain_0!FmaPgo = Choose(cboFmaPgo.ListIndex + 1, FMAPGO_CON, FMAPGO_CRE)
'         .uorstMain_0!EstCCo = IIf(chkEstado.Value = vbChecked, ESTCCO_ACT, ESTCCO_INA)
'         .uorstMain_0!CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         .uorstMain_0!FEmDoc = dtpFEmDoc.Value
'         .uorstMain_0!CodMon = optMoneda(1).Value
'         .uorstMain_0!CodCli = txtDato(0).Text
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain_0!CodDro
         txtLlave(1).Text = .uorstMain_0!NroCpb
      
        'Datos.
        'cambio glocpb nulo= "" de los contrario es el igual al campo glosa
         txtDato(0).Text = IIf(IsNull(.uorstMain_0!glocpb), "", .uorstMain_0!glocpb)
         dtpFehCpb.Value = .uorstMain_0!FehCpb
'         cboFmaPgo.ListIndex = InStr(FMAPGO_CON + FMAPGO_CRE, uorstMain_0!FmaPgo) - 1
'         chkEstado.Value = IIf(uorstMain_0!EstCCo = ESTCCO_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(uorstMain_0!CodSoc), "", uorstMain_0!CodSoc)
'          dtpFEmDoc.Value = .uorstMain_0!FEmDoc
'         optMoneda(1).Value = uorstMain_0!CodMon
'         txtDato(0).Text = IIf(IsNull(.uorstMain_0!CodCli), "", uorstMain_0!CodCli)
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
'   cboFmaPgo.ListIndex = FMAPGO_CON_IDX
''   chkEstado.Value = vbChecked
''   dcoSocio.BoundText = ""
   dtpFehCpb.Value = IIf(Month(Date) = Val(gsMesAct) And Year(Date) = Val(gsAnoAct), Date, gfUltDia("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct))
''   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With
'   txtDeta(0).Text = gnPctIGV
'   txtDeta(1).Text = 0

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
'   lblDatoDeta(3).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
'   cboFmaPgo.Enabled = tbHabilitar
'   dtpFEmDoc.Enabled = tbHabilitar
   dtpFehCpb.Enabled = tbHabilitar
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

  'Ayudas.
   cmdLlaveAyud(0).Enabled = (pbNuevo)
'   cmdDatoAyud(3).Enabled = tbHabilitar
   lblLlaveDeta(0).Enabled = tbHabilitar
'   lblDatoDeta(3).Enabled = tbHabilitar
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
            .Item(dnNum).Caption = "Cuenta"
            .Item(dnNum).Width = 850 '70 * (frmTCpbGrd.uorstMain_1.Fields("CodCta").DefinedSize + 1)
         Case 2
            .Item(dnNum).Caption = "C.Cto."
            .Item(dnNum).Width = 90 * (frmTCpbGrd.uorstMain_1.Fields("CodCCo").DefinedSize + 1)
         Case 3
            .Item(dnNum).Caption = "Auxiliar"
            .Item(dnNum).Width = 1000 '100 * (frmTCpbGrd.uorstMain_1.Fields("CodAux").DefinedSize + 1)
         Case 4
            .Item(dnNum).Caption = "Doc"
            .Item(dnNum).Width = 100 * (frmTCpbGrd.uorstMain_1.Fields("AbvTDc").DefinedSize + 2)
         Case 5
            .Item(dnNum).Caption = "Serie"
            .Item(dnNum).Width = 100 * (frmTCpbGrd.uorstMain_1.Fields("SerDoc").DefinedSize + 2)
         Case 6
            .Item(dnNum).Caption = "Numero"
            .Item(dnNum).Width = 90 * (frmTCpbGrd.uorstMain_1.Fields("NroDoc").DefinedSize + 1)
         Case 7
            .Item(dnNum).Caption = "Glosa"
            .Item(dnNum).Width = 1100
         Case 8
            .Item(dnNum).Caption = "Debe"
            .Item(dnNum).Width = 1000 ' * (uorstDetalle.Fields("ImpMN").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 9
            .Item(dnNum).Caption = "Haber"
            .Item(dnNum).Width = 1000 ' * (uorstDetalle.Fields("ImpME").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 10
            .Item(dnNum).Caption = "Tipo"
            .Item(dnNum).Width = 100 * (frmTCpbGrd.uorstMain_0.Fields("TpoGnr").DefinedSize + 5)
            .Item(dnNum).Alignment = dbgCenter
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
      .uorstMain_1!CodDro = txtLlave(0).Text
      .uorstMain_1!NroCpb = txtLlave(1).Text
      With .uorstUltiItem
         .Source = "SELECT IFNULL(MAX(NroIte), 0) AS cUltIte " _
                 & "FROM COCpbDet " _
                 & "WHERE CodDro='" & frmTCpbCab.txtLlave(0).Text & "' And NroCpb='" & frmTCpbCab.txtLlave(1).Text & "' And MesPvs='" & gsMesAct & "'"
'                & "WHERE CodDro='" & frmTCpbGrd.uorstMain_0!CodDro & "' And NroCpb='" & frmTCpbGrd.uorstMain_0!NroCpb & "'"
         .Open
         frmTCpbGrd.uorstMain_1!NroIte = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
         .Close
      End With
      .uorstMain_1!FehOpe = dtpFehCpb.Value
      .uorstMain_1!CodCta = "FF"
      .uorstMain_1!GloIte = "Item de Cuadre Forzado"
      .uorstMain_1!TpoMon = TPOMON_NAC
      .uorstMain_1!FeEDoc = dtpFehCpb.Value                 ' Necesario fecha por defecto
      .uorstMain_1!FeVDoc = dtpFehCpb.Value
      .uorstMain_1!FeRDoc = dtpFehCpb.Value
    '[Modificado por Miguel Angel (17/01/2004).
      .uorstMain_1!ImpTCb = Format(1, FORMATO_NUM_2)
      If frmTCpbGrd.uorstTGTCb.RecordCount <> 0 Then
        frmTCpbGrd.uorstTGTCb.MoveFirst
        frmTCpbGrd.uorstTGTCb.Find "FehTCb = '" & .uorstMain_1!FeEDoc & "'"
        If Not frmTCpbGrd.uorstTGTCb.EOF Then
           .uorstMain_1!ImpTCb = Format(frmTCpbGrd.uorstTGTCb!ImpTCb_Vta, FORMATO_NUM_2)
        End If
      End If
'      If txtDeta(2).Text = 0 Then
'         .uorstMain_1!ImpTCb = gfRedond(CDec(txtDeta(1).Text) / CDec(txtDeta(3).Text), 3)
'      Else
'         .uorstMain_1!ImpTCb = gfRedond(CDec(txtDeta(0).Text) / CDec(txtDeta(2).Text), 3)
'      End If
   ']
      .uorstMain_1!TpoGnr = TPOGNR_DRO
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
      .uorstMain_1!MesPvs = gsMesAct
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
         .uorstMain_1.Find "NroIte='" & dnNroIte & "'"
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
      If frmTCpbGrd.uorstMain_0.RecordCount > 0 And frmTCpbGrd.uorstMain_0!TpoGnr <> TPOGNR_DRO Then
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
      cmdImprimir.Enabled = tbgraba
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
      MsgBox "La fecha debe ser del Mes y Año que se provisiona.", vbExclamation
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
cmdImprimir.Enabled = IIf(pbNuevo, False, IIf(taOpciones(3) Or taOpciones(4), True, False))
End Property
