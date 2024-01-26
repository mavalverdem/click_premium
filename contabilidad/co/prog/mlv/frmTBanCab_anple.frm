VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTBanCab_anple 
   Caption         =   "[Entidad]"
   ClientHeight    =   6660
   ClientLeft      =   2220
   ClientTop       =   2595
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9330
   Begin VB.CheckBox chkGenPrc 
      Caption         =   "Procesado"
      Enabled         =   0   'False
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   7200
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2720
      Width           =   1815
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   9
      Left            =   6360
      TabIndex        =   25
      Top             =   2940
      Width           =   390
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   8
      Left            =   1020
      TabIndex        =   20
      Top             =   2940
      Width           =   5250
   End
   Begin VB.CheckBox chkGenCpb 
      Caption         =   "Comprobante Diario"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   7200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1000
      Width           =   1815
   End
   Begin VB.Frame fraDocumento 
      Caption         =   " Documento "
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   7200
      TabIndex        =   50
      Top             =   1720
      Width           =   1980
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   10
         Left            =   1680
         Picture         =   "frmTBanCab_anple.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   10
         Left            =   60
         TabIndex        =   21
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   7
         Left            =   60
         TabIndex        =   23
         Top             =   570
         Width           =   1785
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
         Height          =   285
         Index           =   10
         Left            =   600
         TabIndex        =   70
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1400
      Width           =   1980
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   6
      Left            =   5445
      TabIndex        =   6
      Top             =   690
      Width           =   735
   End
   Begin VB.ComboBox cboTpoTCb 
      Height          =   315
      Left            =   4485
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   690
      Width           =   915
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   4
      Left            =   1020
      TabIndex        =   16
      Top             =   2295
      Width           =   1275
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   4
      Left            =   6870
      Picture         =   "frmTBanCab_anple.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2295
      Width           =   255
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   3
      Left            =   1020
      TabIndex        =   14
      Top             =   1980
      Width           =   615
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   3
      Left            =   6870
      Picture         =   "frmTBanCab_anple.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1980
      Width           =   255
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   5
      Left            =   6870
      Picture         =   "frmTBanCab_anple.frx":04FE
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2625
      Width           =   255
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   5
      Left            =   1020
      TabIndex        =   18
      Top             =   2625
      Width           =   615
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   2
      Left            =   1020
      TabIndex        =   12
      Top             =   1650
      Width           =   975
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   2
      Left            =   6870
      Picture         =   "frmTBanCab_anple.frx":06A8
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1650
      Width           =   255
   End
   Begin VB.ComboBox cboTpoBan 
      Height          =   315
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   1020
      TabIndex        =   10
      Top             =   1335
      Width           =   6120
   End
   Begin VB.TextBox txtLlave 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   8415
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   8850
      Picture         =   "frmTBanCab_anple.frx":0852
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   1020
      TabIndex        =   8
      Top             =   1020
      Width           =   6120
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
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6090
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
         Picture         =   "frmTBanCab_anple.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   61
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
         Picture         =   "frmTBanCab_anple.frx":0F2E
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Picture         =   "frmTBanCab_anple.frx":10D8
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Picture         =   "frmTBanCab_anple.frx":1282
         Style           =   1  'Graphical
         TabIndex        =   56
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
         Picture         =   "frmTBanCab_anple.frx":1384
         Style           =   1  'Graphical
         TabIndex        =   55
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
         Picture         =   "frmTBanCab_anple.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   54
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
         Picture         =   "frmTBanCab_anple.frx":15D0
         Style           =   1  'Graphical
         TabIndex        =   60
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
         Left            =   8610
         Picture         =   "frmTBanCab_anple.frx":171A
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   0
         Width           =   660
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
         Picture         =   "frmTBanCab_anple.frx":1864
         Style           =   1  'Graphical
         TabIndex        =   57
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
         Picture         =   "frmTBanCab_anple.frx":1966
         Style           =   1  'Graphical
         TabIndex        =   59
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
         Picture         =   "frmTBanCab_anple.frx":1A68
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   0
         Width           =   645
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
         Picture         =   "frmTBanCab_anple.frx":1B6A
         Style           =   1  'Graphical
         TabIndex        =   58
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
         Left            =   7900
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   7840
         X2              =   7840
         Y1              =   0
         Y2              =   550
      End
   End
   Begin VB.TextBox txtLlave 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   4650
      TabIndex        =   2
      Top             =   120
      Width           =   520
   End
   Begin MSDataGridLib.DataGrid dgrDetalle 
      Height          =   1935
      Left            =   0
      TabIndex        =   51
      Top             =   3240
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   3413
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
   Begin VB.Frame fraTotales 
      ForeColor       =   &H80000002&
      Height          =   900
      Left            =   0
      TabIndex        =   65
      Top             =   5160
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   280
         Index           =   4
         Left            =   1440
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   280
         Index           =   5
         Left            =   1440
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   525
         Width           =   1755
      End
      Begin VB.CommandButton cmdAuxiliar 
         Caption         =   "&Auxiliar"
         Height          =   375
         Left            =   30
         TabIndex        =   66
         Top             =   105
         Width           =   1215
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
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   3
         Left            =   7440
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   525
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
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   0
         Left            =   5640
         TabIndex        =   28
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
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   1
         Left            =   7440
         TabIndex        =   29
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
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   2
         Left            =   5640
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   525
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
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   13
         Left            =   4440
         TabIndex        =   68
         Top             =   540
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
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   12
         Left            =   4440
         TabIndex        =   67
         Top             =   210
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpFehBan 
      Height          =   300
      Left            =   1020
      TabIndex        =   4
      Top             =   690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      Format          =   63832065
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
      Height          =   285
      Index           =   9
      Left            =   6735
      TabIndex        =   49
      Top             =   2940
      Width           =   2475
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Portador :"
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
      Index           =   14
      Left            =   105
      TabIndex        =   48
      Top             =   2955
      Width           =   705
   End
   Begin VB.Shape shpCuadro 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   465
      Left            =   6795
      Top             =   510
      Width           =   2460
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda Funcional :"
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
      Height          =   195
      Index           =   11
      Left            =   7200
      TabIndex        =   41
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label lblTexto 
      Caption         =   "Tipo de Cambio:"
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
      Height          =   195
      Index           =   9
      Left            =   3015
      TabIndex        =   36
      Top             =   720
      Width           =   1410
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Auxiliar:"
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
      Height          =   195
      Index           =   7
      Left            =   105
      TabIndex        =   44
      Top             =   2325
      Width           =   900
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
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   45
      Top             =   2295
      Width           =   4605
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "C.Costo:"
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
      Height          =   195
      Index           =   6
      Left            =   105
      TabIndex        =   42
      Top             =   2010
      Width           =   900
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
      Height          =   285
      Index           =   3
      Left            =   1620
      TabIndex        =   43
      Top             =   1980
      Width           =   5265
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
      Height          =   285
      Index           =   5
      Left            =   1620
      TabIndex        =   47
      Top             =   2625
      Width           =   5265
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Caja:"
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
      Height          =   195
      Index           =   8
      Left            =   105
      TabIndex        =   46
      Top             =   2655
      Width           =   900
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
      Height          =   285
      Index           =   2
      Left            =   1980
      TabIndex        =   40
      Top             =   1650
      Width           =   4890
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta:"
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
      Height          =   195
      Index           =   5
      Left            =   105
      TabIndex        =   39
      Top             =   1665
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Transacción :"
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
      Height          =   195
      Index           =   10
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   990
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   38
      Top             =   1350
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   6975
      TabIndex        =   34
      Top             =   645
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
      Height          =   285
      Index           =   0
      Left            =   5175
      TabIndex        =   33
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   37
      Top             =   1035
      Width           =   900
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   35
      Top             =   720
      Width           =   900
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   3975
      TabIndex        =   32
      Top             =   180
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   60
      X2              =   9240
      Y1              =   510
      Y2              =   510
   End
End
Attribute VB_Name = "frmTBanCab_anple"
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
Private pnCta_IndCCo As Integer, _
        psCodCCo_Def As String
'[
Private salirform As Boolean

Private Sub ppGeneraComprobante()
  On Error GoTo Err
  Dim sSentencia As String, sCodFlujo As String
  Dim sIndDoc As String, sIndCco As String, sIndFjo As String
  Dim nNroItem As Long, nRegistro As Long
  
  salirform = True
  
  ' Verifico que exista la cabecera del comprobante de caja
  sSentencia = "SELECT COUNT(*) AS Registro "
  sSentencia = sSentencia & "FROM cobancab "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND coddro='" & Trim(txtLlave(0).Text) & "' "
  sSentencia = sSentencia & "AND nroban='" & Trim(txtLlave(1).Text) & "'"
  If frmTBanGrd_anple.uorstCOTCbMes.State = adStateOpen Then frmTBanGrd_anple.uorstCOTCbMes.Close
  frmTBanGrd_anple.uorstCOTCbMes.Source = sSentencia
  frmTBanGrd_anple.uorstCOTCbMes.Open
  nRegistro = frmTBanGrd_anple.uorstCOTCbMes!registro
  frmTBanGrd_anple.uorstCOTCbMes.Close
  
  If nRegistro = 0 Then Exit Sub
  
  'Verifico si el Comprobante esta cuadrado
  
  Set porstMRp = New ADODB.Recordset
  With porstMRp
    .ActiveConnection = frmTBanGrd_anple.uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With

  With porstMRp
    .Source = "SELECT c.fehban AS fehcpb, " & Choose(gsIdioma, "c.globan", "c.globanx") & " AS glocpb, a.ImpTcb, "
    .Source = .Source & "(CASE a.pvsdoc WHEN " & INDPREGEN_ACT & " THEN '" & TPOPVS_PVS & "' ELSE '" & TPOPVS_CAN & "' END) AS TpoPvs, a.TpoMon, "
    .Source = .Source & "a.MesPvs, c.fehban AS FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, "
    .Source = .Source & "a.codaux, d.razaux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.nroban,')')", "('('+a.CodDro+'-'+a.nroban+')')") & " AS cComprobante, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    .Source = .Source & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
    .Source = .Source & "a.ImpME, a.ImpMN, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_EGR & " THEN a.impME ELSE 0 END) as DebME, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_ING & " THEN a.ImpME ELSE 0 END) as HabME, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_EGR & " THEN a.ImpMN ELSE 0 END) as DebMN, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_ING & " THEN a.ImpMN ELSE 0 END) as HabMN, "
    .Source = .Source & "f.forimp, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf "
    .Source = .Source & "FROM ((cobancab c "
    .Source = .Source & "LEFT JOIN cobandet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.nroban=a.nroban) "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
    .Source = .Source & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
    .Source = .Source & "LEFT JOIN cobco f ON c.codemp=f.codemp AND c.codbco=f.codbco "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs = '" & gsMesAct & "' "
    .Source = .Source & "AND a.CodDro = '" & txtLlave(0).Text & "' AND a.nroban = '" & txtLlave(1).Text & "' "
    .Source = .Source & "UNION "
    .Source = .Source & "SELECT cab.fehban AS fehcpb, " & Choose(gsIdioma, "cab.globan", "cab.globanx") & " AS glocpb, cab.ImpTcb, "
    .Source = .Source & "'" & TPOPVS_CAN & "' AS TpoPvs, cab.TpoMon, "
    .Source = .Source & "cab.MesPvs, cab.fehban AS FehOpe, cab.CodCta, " & Choose(gsIdioma, "cta.DetCta", "cta.DetCtax") & " AS Detcta, cab.CodCCo, "
    .Source = .Source & "(CASE cta.inddoc WHEN " & INDDOC_ACT & " THEN cab.codaux ELSE '' END) AS codaux, "
    .Source = .Source & "(CASE cta.inddoc WHEN " & INDDOC_ACT & " THEN aux.razaux ELSE '' END) AS razaux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',cab.CodDro, '-', cab.nroban,')')", "('('+cab.CodDro+'-'+cab.nroban+')')") & " AS cComprobante, "
    .Source = .Source & "cab.docban AS cDocumento, "
    .Source = .Source & "'' AS CodTDc, '' AS SerDoc, cab.docban AS NroDoc, '' AS RefDoc, " & Choose(gsIdioma, "cab.Globan", "cab.Globanx") & " AS GloIte, "
    .Source = .Source & "cab.ImpME, cab.ImpMN, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_ING & " THEN Abs(cab.impME) ELSE 0 END) as DebME, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_EGR & " THEN Abs(cab.ImpME) ELSE 0 END) as HabME, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_ING & " THEN Abs(cab.ImpMN) ELSE 0 END) as DebMN, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_EGR & " THEN Abs(cab.ImpMN) ELSE 0 END) as HabMN, "
    .Source = .Source & "bco.forimp, cab.UsrCre, cab.FyHCre, cab.UsrMdf, cab.FyHMdf "
    .Source = .Source & "FROM cobancab cab "
    .Source = .Source & "LEFT JOIN TGAux aux ON cab.codemp=aux.codemp AND cab.CodAux=aux.CodAux "
    .Source = .Source & "LEFT JOIN Cocta cta ON cab.codemp=cta.codemp AND cab.pdoano=cta.pdoano AND cab.Codcta=cta.Codcta "
    .Source = .Source & "LEFT JOIN cobco bco ON cab.codemp=bco.codemp AND cab.codbco=bco.codbco "
    .Source = .Source & "WHERE cab.codemp='" & gsCodEmp & "' AND cab.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND cab.MesPvs = '" & gsMesAct & "' "
    .Source = .Source & "AND cab.CodDro = '" & txtLlave(0).Text & "' AND cab.nroban = '" & txtLlave(1).Text & "' "
    .Open
  End With
  
  
  
  If porstMRp.RecordCount > 0 Then
  
  Dim sumadebe As Double
  Dim sumahaber As Double
  
  sumadebe = 0
  sumahaber = 0
  
  porstMRp.MoveFirst
  
  While Not porstMRp.EOF
      
      sumadebe = sumadebe + IIf(IsNull(porstMRp!DebMN), 0, porstMRp!DebMN)
      sumahaber = sumahaber + IIf(IsNull(porstMRp!HabMN), 0, porstMRp!HabMN)
      
      porstMRp.MoveNext
  Wend
       
  If sumadebe <> sumahaber Then
    MsgBox ("Comprobante no esta Cuadrado")
    salirform = False
    Exit Sub
  End If
  
  End If
  
  porstMRp.Close
  Set porstMRp = Nothing

    
  
  '**************************************************************
  
  ' Elimino el comprobante de diario
  frmTBanGrd_anple.uocnnMain.BeginTrans   ' Inicia Transaccion
  
  sSentencia = "DELETE FROM cocpbcab "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND coddro='" & Trim(txtLlave(0).Text) & "' "
  sSentencia = sSentencia & "AND nrocpb='" & Trim(txtLlave(1).Text) & "'"
  frmTBanGrd_anple.uocnnMain.Execute sSentencia
  
  If frmTBanGrd_anple.uorstCOTCbMes.State = adStateOpen Then frmTBanGrd_anple.uorstCOTCbMes.Close
  ' Genero el detalle del comprobante
  sSentencia = "SELECT det.*, cta.inddoc, cta.indcco, cta.indfjo "
  sSentencia = sSentencia & "FROM cobandet det "
  sSentencia = sSentencia & "LEFT JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.Codcta=cta.Codcta "
  sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND det.mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND det.coddro='" & Trim(txtLlave(0).Text) & "' "
  sSentencia = sSentencia & "AND det.nroban='" & Trim(txtLlave(1).Text) & "'"
  frmTBanGrd_anple.uorstCOTCbMes.Source = sSentencia
  frmTBanGrd_anple.uorstCOTCbMes.Open
  chkGenCpb.Value = vbChecked
  'chkGenPrc.Value = vbChecked
  ' Genero la cabecera del comprobante
  sSentencia = "INSERT INTO CoCpbCab(codemp, pdoano, mespvs, CodDro, NroCpb, FehCpb, GloCpb, glocpbx, TpoGnr, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf)"
  sSentencia = sSentencia & " VALUES("
  sSentencia = sSentencia & "'" & gsCodEmp & "', "
  sSentencia = sSentencia & "'" & gsAnoAct & "', "
  sSentencia = sSentencia & "'" & gsMesAct & "', "
  sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
  sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(txtDato(gsIdioma - 1).Text = "", "Null", "'" & txtDato(gsIdioma - 1).Text & "'") & ", "
  sSentencia = sSentencia & IIf(txtDato(2 - gsIdioma).Text = "", "Null", "'" & txtDato(2 - gsIdioma).Text & "'") & ", "
  sSentencia = sSentencia & "'" & TPOGNR_BAN & "', "
  sSentencia = sSentencia & "'" & INDNCU_FAL & "', "
  sSentencia = sSentencia & "'" & INDANU_FAL & "', "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null)"
  frmTBanGrd_anple.uocnnMain.Execute sSentencia
  ' No se actualizo el numero de comprobante de diario se realizo en grabacion
  If frmTBanGrd_anple.uorstCOTCbMes.RecordCount <> 0 Then
  ' Grabación de detalle de comprobante
    While Not frmTBanGrd_anple.uorstCOTCbMes.EOF
      sSentencia = "INSERT INTO CoCpbDet(codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, FehOpe, CodCta, CodCCo, CodAux, CodTDc, SerDoc, NroDoc, FeEDoc, FeVDoc, "
      sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, TpoCtb, TpoPvs, TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, IndFjo_Det, UsrCre, FyHCre, UsrMdf, FyHMdf,tpodoc) "
      sSentencia = sSentencia & "VALUES("
      sSentencia = sSentencia & "'" & gsCodEmp & "', "
      sSentencia = sSentencia & "'" & gsAnoAct & "', "
      sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
      sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
      sSentencia = sSentencia & frmTBanGrd_anple.uorstCOTCbMes!nroitem & ", "
      sSentencia = sSentencia & "'" & gsMesAct & "', "
      sSentencia = sSentencia & frmTBanGrd_anple.uorstCOTCbMes!nroitem & ", "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & "'" & frmTBanGrd_anple.uorstCOTCbMes!CodCta & "', "
      sSentencia = sSentencia & IIf(frmTBanGrd_anple.uorstCOTCbMes!indcco = INDCCO_INA Or IsNull(frmTBanGrd_anple.uorstCOTCbMes!codcco), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!codcco & "'") & ", "
      sSentencia = sSentencia & IIf(frmTBanGrd_anple.uorstCOTCbMes!IndDoc = INDCCO_INA And IsNull(Trim(frmTBanGrd_anple.uorstCOTCbMes!codaux)), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!codaux & "'") & ", "
      sSentencia = sSentencia & IIf(frmTBanGrd_anple.uorstCOTCbMes!IndDoc = INDDOC_INA Or IsNull(frmTBanGrd_anple.uorstCOTCbMes!codtdc), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!codtdc & "'") & ", "
      sSentencia = sSentencia & IIf(frmTBanGrd_anple.uorstCOTCbMes!IndDoc = INDDOC_INA Or IsNull(frmTBanGrd_anple.uorstCOTCbMes!serdoc), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!serdoc & "'") & ", "
      sSentencia = sSentencia & IIf(frmTBanGrd_anple.uorstCOTCbMes!IndDoc = INDDOC_INA Or IsNull(frmTBanGrd_anple.uorstCOTCbMes!nrodoc), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!nrodoc & "'") & ", "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(IsNull(frmTBanGrd_anple.uorstCOTCbMes!RefDoc), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!RefDoc & "'") & ", "
      sSentencia = sSentencia & IIf(IsNull(frmTBanGrd_anple.uorstCOTCbMes!GloIte), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!GloIte & "'") & ", "
      sSentencia = sSentencia & IIf(IsNull(frmTBanGrd_anple.uorstCOTCbMes!GloItex), "Null", "'" & frmTBanGrd_anple.uorstCOTCbMes!GloItex & "'") & ", "
      sSentencia = sSentencia & "'" & IIf(Val(frmTBanGrd_anple.uorstCOTCbMes!tpoban) = TPOBAN_EGR, TPOCTB_DEB, TPOCTB_HAB) & "', "
      sSentencia = sSentencia & "'" & IIf(CInt(frmTBanGrd_anple.uorstCOTCbMes!pvsdoc) = INDPREGEN_INA, TPOPVS_CAN, TPOPVS_PVS) & "', "
      sSentencia = sSentencia & "'" & frmTBanGrd_anple.uorstCOTCbMes!tpomon & "', "
      sSentencia = sSentencia & "'" & frmTBanGrd_anple.uorstCOTCbMes!TpoTcb & "', "
      sSentencia = sSentencia & frmTBanGrd_anple.uorstCOTCbMes!ImpTCb & ", "
      sSentencia = sSentencia & CDec(frmTBanGrd_anple.uorstCOTCbMes!ImpMN) & ", "
      sSentencia = sSentencia & CDec(frmTBanGrd_anple.uorstCOTCbMes!ImpME) & ", "
      sSentencia = sSentencia & "'" & TPOGNR_BAN & "', "
      sCodFlujo = IIf(frmTBanGrd_anple.uorstCOTCbMes!IndFjo = INDFJO_ACT, txtDato(5).Text, "")
      sSentencia = sSentencia & IIf(sCodFlujo = "", INDFJO_INA, INDFJO_ACT) & ", "
      sSentencia = sSentencia & "'" & gsAbvUsr & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
      sSentencia = sSentencia & "Null, Null,'" & txtDato(10).Text & "')"
      frmTBanGrd_anple.uocnnMain.Execute sSentencia
      
      If sCodFlujo <> "" Then
        ' Grabación de detalle de flujo de caja
        sSentencia = "INSERT INTO CoCpbDetFjo(codemp, pdoano, MesPvs, CodDro, NroCpb, NroIte, NroOrd, CodFjo,"
        sSentencia = sSentencia & " CodCta, TpoCtb, ImpMN, ImpME, UsrCre, FyHCre, UsrMdf, FyHMdf)"
        sSentencia = sSentencia & " VALUES("
        sSentencia = sSentencia & "'" & gsCodEmp & "', "
        sSentencia = sSentencia & "'" & gsAnoAct & "', "
        sSentencia = sSentencia & "'" & gsMesAct & "', "
        sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
        sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
        sSentencia = sSentencia & frmTBanGrd_anple.uorstCOTCbMes!nroitem & ", "
        sSentencia = sSentencia & " '1',"
        sSentencia = sSentencia & " '" & sCodFlujo & "',"
        sSentencia = sSentencia & " '" & frmTBanGrd_anple.uorstCOTCbMes!CodCta & "',"
        sSentencia = sSentencia & " '" & IIf(Val(frmTBanGrd_anple.uorstCOTCbMes!tpoban) = TPOBAN_EGR, TPOCTB_DEB, TPOCTB_HAB) & "',"
        sSentencia = sSentencia & CDec(frmTBanGrd_anple.uorstCOTCbMes!ImpMN) & ", "
        sSentencia = sSentencia & frmTBanGrd_anple.uorstCOTCbMes!ImpME & ", "
        sSentencia = sSentencia & " '" & gsAbvUsr & "',"
        sSentencia = sSentencia & " '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
        sSentencia = sSentencia & " Null, Null)"
        frmTBanGrd_anple.uocnnMain.Execute sSentencia
      End If
        'sSentencia = sSentencia & " '" & IIf(frmTBanGrd_anple.uorstCOTCbMes!tpoban = TPOBAN_EGR, TPOCTB_DEB, TPOCTB_HAB) & "',"
      nNroItem = CLng(frmTBanGrd_anple.uorstCOTCbMes!nroitem)
      frmTBanGrd_anple.uorstCOTCbMes.MoveNext
    Wend
    frmTBanGrd_anple.uorstCOTCbMes.Close
    ' Antes cuenta de caja
  End If
  ' Cuenta  de caja bancos
  sSentencia = "SELECT cta.inddoc, cta.indcco, cta.indfjo "
  sSentencia = sSentencia & "FROM cocta cta "
  sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND cta.codcta='" & txtDato(2).Text & "'"
  If frmTBanGrd_anple.uorstCOTCbMes.State = adStateOpen Then frmTBanGrd_anple.uorstCOTCbMes.Close
  frmTBanGrd_anple.uorstCOTCbMes.Source = sSentencia
  frmTBanGrd_anple.uorstCOTCbMes.Open
  If frmTBanGrd_anple.uorstCOTCbMes.RecordCount <> 0 Then
    sIndDoc = frmTBanGrd_anple.uorstCOTCbMes!IndDoc
    sIndCco = frmTBanGrd_anple.uorstCOTCbMes!indcco
    sIndFjo = frmTBanGrd_anple.uorstCOTCbMes!IndFjo
  End If
  frmTBanGrd_anple.uorstCOTCbMes.Close
  sCodFlujo = IIf(sIndFjo = INDFJO_ACT, txtDato(5).Text, "")
  nNroItem = nNroItem + 1
  sSentencia = "INSERT INTO CoCpbDet(codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, FehOpe, CodCta, CodCCo, CodAux, CodTDc, SerDoc, NroDoc, FeEDoc, FeVDoc, "
  sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, TpoCtb, TpoPvs, TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, IndFjo_Det, UsrCre, FyHCre, UsrMdf, FyHMdf,tpodoc) "
  sSentencia = sSentencia & "VALUES("
  sSentencia = sSentencia & "'" & gsCodEmp & "', "
  sSentencia = sSentencia & "'" & gsAnoAct & "', "
  sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
  sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
  sSentencia = sSentencia & nNroItem & ", "
  sSentencia = sSentencia & "'" & gsMesAct & "', "
  sSentencia = sSentencia & nNroItem & ", "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "'" & txtDato(2).Text & "', "
  sSentencia = sSentencia & IIf(sIndCco = INDCCO_INA Or txtDato(3).Text = "", "Null", "'" & txtDato(3).Text & "'") & ", "
  sSentencia = sSentencia & IIf(sIndDoc = INDDOC_INA Or txtDato(4).Text = "", "'" & txtDato(4).Text & "'", "'" & txtDato(4).Text & "'") & ", "
  sSentencia = sSentencia & "Null, "
  sSentencia = sSentencia & "Null, "
  sSentencia = sSentencia & IIf(sIndDoc = INDDOC_INA Or txtDato(7).Text = "", "Null", "'" & txtDato(7).Text & "'") & ", "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehBan.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  'sSentencia = sSentencia & "'" & Choose(cboTpoDoc.ListIndex + 1, "", "DPS", "GRO", "TRA", "ORD", "DEB", "CRE", "CHQ", "OTR", "EFE", "PEX", "LTR", "CGE") & "-" & txtDato(7).Text & "', "
  sSentencia = sSentencia & "'" & txtDato(10).Text & "-" & txtDato(7).Text & "', "
  sSentencia = sSentencia & IIf(txtDato(gsIdioma - 1).Text = "", "Null", "'" & txtDato(gsIdioma - 1).Text & "'") & ", "
  sSentencia = sSentencia & IIf(txtDato(2 - gsIdioma).Text = "", "Null", "'" & txtDato(2 - gsIdioma).Text & "'") & ", "
  sSentencia = sSentencia & "'" & IIf(cboTpoBan.ListIndex = TPOBAN_ING, TPOCTB_DEB, TPOCTB_HAB) & "', "
  sSentencia = sSentencia & "'" & TPOPVS_CAN & "', "
  sSentencia = sSentencia & "'" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT) & "', "
  sSentencia = sSentencia & "'" & IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR) & "', "
  sSentencia = sSentencia & CDec(txtDato(6).Text) & ", "
  sSentencia = sSentencia & Abs(CDec(CDec(txtDeta(4).Text))) & ", "
  sSentencia = sSentencia & Abs(CDec(CDec(txtDeta(5).Text))) & ", "
  sSentencia = sSentencia & "'" & TPOGNR_BAN & "', "
  sSentencia = sSentencia & IIf(sCodFlujo = "", INDFJO_INA, INDFJO_ACT) & ", "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null,'" & txtDato(10).Text & "')"
  frmTBanGrd_anple.uocnnMain.Execute sSentencia
    
  If sCodFlujo <> "" Then
    ' Grabación de detalle de flujo de caja
    sSentencia = "INSERT INTO CoCpbDetFjo(codemp, pdoano, MesPvs, CodDro, NroCpb, NroIte, NroOrd, CodFjo,"
    sSentencia = sSentencia & " CodCta, TpoCtb, ImpMN, ImpME, UsrCre, FyHCre, UsrMdf, FyHMdf)"
    sSentencia = sSentencia & " VALUES("
    sSentencia = sSentencia & "'" & gsCodEmp & "', "
    sSentencia = sSentencia & "'" & gsAnoAct & "', "
    sSentencia = sSentencia & "'" & gsMesAct & "', "
    sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
    sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
    sSentencia = sSentencia & nNroItem & ", "
    sSentencia = sSentencia & " '1',"
    sSentencia = sSentencia & " '" & sCodFlujo & "',"
    sSentencia = sSentencia & " '" & txtDato(2).Text & "',"
    sSentencia = sSentencia & "'" & IIf(cboTpoBan.ListIndex = TPOBAN_ING, TPOCTB_DEB, TPOCTB_HAB) & "', "
    sSentencia = sSentencia & Abs(CDec(CDec(txtDeta(4).Text))) & ", "
    sSentencia = sSentencia & Abs(CDec(CDec(txtDeta(5).Text))) & ", "
    sSentencia = sSentencia & " '" & gsAbvUsr & "',"
    sSentencia = sSentencia & " '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
    sSentencia = sSentencia & " Null, Null)"
    frmTBanGrd_anple.uocnnMain.Execute sSentencia
  End If
  ' Actualizo los datos adicionales
  chkGenCpb.Value = vbChecked
  'chkGenPrc.Value = vbChecked
  ' Actualizo los datos de bancos
  If proceso = False Then
  frmTBanGrd_anple.uorstMain_0!ImpMN = CDec(txtDeta(4).Text)
  frmTBanGrd_anple.uorstMain_0!ImpME = CDec(txtDeta(5).Text)
  frmTBanGrd_anple.uorstMain_0!gencpb = IIf(chkGenCpb.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
  frmTBanGrd_anple.uorstMain_0!genprc = IIf(chkGenPrc.Value = vbChecked, INDPREGEN_ACTx, INDPREGEN_INAx)
  frmTBanGrd_anple.uorstMain_0.Update
  Else
  frmTBanGrd_anple.uorstMain_0Fil!ImpMN = CDec(txtDeta(4).Text)
  frmTBanGrd_anple.uorstMain_0Fil!ImpME = CDec(txtDeta(5).Text)
  frmTBanGrd_anple.uorstMain_0Fil!gencpb = IIf(chkGenCpb.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
  frmTBanGrd_anple.uorstMain_0Fil!genprc = IIf(chkGenPrc.Value = vbChecked, INDPREGEN_ACTx, INDPREGEN_INAx)
  frmTBanGrd_anple.uorstMain_0Fil.Update
  End If
  
  frmTBanGrd_anple.uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  If proceso = False Then
  ' [Actualiza grilla y busco registro actual
  frmTBanGrd_anple.uorstMain_Grd.Requery
  frmTBanGrd_anple.ppDatosGrid
  frmTBanGrd_anple.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
  ' ]
  Else
  ' [Actualiza grilla y busco registro actual
  frmTBanGrd_anple.uorstMain_GrdFil.Requery
  frmTBanGrd_anple.ppDatosGrid
  frmTBanGrd_anple.uorstMain_GrdFil.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
  ' ]
  
  End If
  
  Exit Sub
Err:
  gpErrores
  frmTBanGrd_anple.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.

End Sub
']

Private Sub cboTpoBan_LostFocus()
  txtLlave(0).Text = Choose(cboTpoBan.ListIndex + 1, gsCodDro_Ing, gsCodDro_Egr)
End Sub
Private Sub CboTpoTCb_LostFocus()
  With frmTBanGrd_anple.uorstTGTCb
    txtDato(6).Text = Format(0, FORMATO_NUM_2)
    If .RecordCount <> 0 Then
      .MoveFirst
      .Find "FehTCb = '" & frmTBanCab_anple.dtpFehBan & "'"
      If .EOF Then
        MsgBox TEXT_9015, vbExclamation
        txtDato(6).SetFocus
      Else
        txtDato(6).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
      End If
    End If
  End With
End Sub







'[
Private Sub cmdAuxiliar_Click()
  frmMAuxGrd.Show vbModal
  frmTBanGrd_anple.uorstTGAux.Requery
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  
  Dim dsFecha As String, dsGirado As String, dsGirado2 As String, dsImporteNumeros As String, dsImporteLetras As String
  Dim dbHayAux As Boolean, dbHay104 As Boolean
  Dim sReporte As String, sTipo As String
  Dim sDesBanco As String, sCheque As String, sDia As String, sMes As String, sAno As String
  Dim nFormato As Integer

  'Agregado por Jorge Gomez 04/01/2010
  cmdGrabar_Click
  'Hasta Aqui

   udFecha = Date                      'Fecha en el encabezado.
   Set porstMRp = New ADODB.Recordset
   With porstMRp
    .ActiveConnection = frmTBanGrd_anple.uocnnMain
'      .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
   End With

' Obtengo la información del comprobante
  With porstMRp
    .Source = "SELECT c.fehban AS fehcpb, " & Choose(gsIdioma, "c.globan", "c.globanx") & " AS glocpb, a.ImpTcb, "
    .Source = .Source & "(CASE a.pvsdoc WHEN " & INDPREGEN_ACT & " THEN '" & TPOPVS_PVS & "' ELSE '" & TPOPVS_CAN & "' END) AS TpoPvs, a.TpoMon, "
    .Source = .Source & "a.MesPvs, c.fehban AS FehOpe, a.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS Detcta, a.CodCCo, "
    .Source = .Source & "a.codaux, d.razaux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',a.CodDro, '-', a.nroban,')')", "('('+a.CodDro+'-'+a.nroban+')')") & " AS cComprobante, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(b.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    .Source = .Source & "a.CodTDc, a.SerDoc, a.NroDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
    .Source = .Source & "a.ImpME, a.ImpMN, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_EGR & " THEN a.impME ELSE 0 END) as DebME, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_ING & " THEN a.ImpME ELSE 0 END) as HabME, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_EGR & " THEN a.ImpMN ELSE 0 END) as DebMN, "
    .Source = .Source & "(CASE a.Tpoban WHEN " & TPOBAN_ING & " THEN a.ImpMN ELSE 0 END) as HabMN, "
    .Source = .Source & "f.forimp, c.UsrCre, c.FyHCre, c.UsrMdf, c.FyHMdf "
    .Source = .Source & "FROM ((cobancab c "
    .Source = .Source & "LEFT JOIN cobandet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.MesPvs=a.MesPvs AND c.CodDro=a.CodDro and c.nroban=a.nroban) "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux "
    .Source = .Source & "LEFT JOIN Cocta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.Codcta=e.Codcta "
    .Source = .Source & "LEFT JOIN cobco f ON c.codemp=f.codemp AND c.codbco=f.codbco "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs = '" & gsMesAct & "' "
    .Source = .Source & "AND a.CodDro = '" & txtLlave(0).Text & "' AND a.nroban = '" & txtLlave(1).Text & "' "
    .Source = .Source & "UNION "
    .Source = .Source & "SELECT cab.fehban AS fehcpb, " & Choose(gsIdioma, "cab.globan", "cab.globanx") & " AS glocpb, cab.ImpTcb, "
    .Source = .Source & "'" & TPOPVS_CAN & "' AS TpoPvs, cab.TpoMon, "
    .Source = .Source & "cab.MesPvs, cab.fehban AS FehOpe, cab.CodCta, " & Choose(gsIdioma, "cta.DetCta", "cta.DetCtax") & " AS Detcta, cab.CodCCo, "
    .Source = .Source & "(CASE cta.inddoc WHEN " & INDDOC_ACT & " THEN cab.codaux ELSE '' END) AS codaux, "
    .Source = .Source & "(CASE cta.inddoc WHEN " & INDDOC_ACT & " THEN aux.razaux ELSE '' END) AS razaux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT('(',cab.CodDro, '-', cab.nroban,')')", "('('+cab.CodDro+'-'+cab.nroban+')')") & " AS cComprobante, "
    .Source = .Source & "cab.docban AS cDocumento, "
    .Source = .Source & "'' AS CodTDc, '' AS SerDoc, cab.docban AS NroDoc, '' AS RefDoc, " & Choose(gsIdioma, "cab.Globan", "cab.Globanx") & " AS GloIte, "
    .Source = .Source & "cab.ImpME, cab.ImpMN, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_ING & " THEN Abs(cab.impME) ELSE 0 END) as DebME, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_EGR & " THEN Abs(cab.ImpME) ELSE 0 END) as HabME, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_ING & " THEN Abs(cab.ImpMN) ELSE 0 END) as DebMN, "
    .Source = .Source & "(CASE cab.Tpoban WHEN " & TPOBAN_EGR & " THEN Abs(cab.ImpMN) ELSE 0 END) as HabMN, "
    .Source = .Source & "bco.forimp, cab.UsrCre, cab.FyHCre, cab.UsrMdf, cab.FyHMdf "
    .Source = .Source & "FROM cobancab cab "
    .Source = .Source & "LEFT JOIN TGAux aux ON cab.codemp=aux.codemp AND cab.CodAux=aux.CodAux "
    .Source = .Source & "LEFT JOIN Cocta cta ON cab.codemp=cta.codemp AND cab.pdoano=cta.pdoano AND cab.Codcta=cta.Codcta "
    .Source = .Source & "LEFT JOIN cobco bco ON cab.codemp=bco.codemp AND cab.codbco=bco.codbco "
    .Source = .Source & "WHERE cab.codemp='" & gsCodEmp & "' AND cab.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND cab.MesPvs = '" & gsMesAct & "' "
    .Source = .Source & "AND cab.CodDro = '" & txtLlave(0).Text & "' AND cab.nroban = '" & txtLlave(1).Text & "' "
    .Open
  End With
  ' Inicializo los datos del cheque
  
  sDesBanco = "": sCheque = "": sDia = "": sMes = "": sAno = ""
  dsGirado = "": dsGirado2 = ""
  
  dbHayAux = False
  dbHay104 = False
  
  dsGirado = Trim(txtDato(8).Text) & "********"
  
'  If cboTpoDoc.ListIndex = TPODOC_CHQ_IND Then
'    dbHay104 = True
'    dsFecha = Format(dtpFehBan.Value, "d mmmm yyyy")
'    dsImporteNumeros = "********" & Format(Abs(CDec(IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CDec(txtDeta(4).Text), CDec(txtDeta(5).Text)))), FORMATO_NUM_1)
'    dsImporteLetras = gfNumLet(Abs(IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CDec(txtDeta(4).Text), CDec(txtDeta(5).Text))), "0") & "********"
'    dsGirado2 = Trim(txtDato(0).Text)
'    sDesBanco = Trim(lblDatoDeta(2).Caption)
'    sCheque = Trim(txtDato(7).Text)
'    sDia = Format(dtpFehBan.Value, "dd")
'    sMes = Format(dtpFehBan.Value, "mm")
'    sAno = Format(dtpFehBan.Value, "yyyy")
'  End If
  
  If txtDato(10).Text = "102" Or txtDato(10).Text = "007" Then
    dbHay104 = True
    dsFecha = Format(dtpFehBan.Value, "d mmmm yyyy")
    dsImporteNumeros = "********" & Format(Abs(CDec(IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CDec(txtDeta(4).Text), CDec(txtDeta(5).Text)))), FORMATO_NUM_1)
    dsImporteLetras = gfNumLet(Abs(IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CDec(txtDeta(4).Text), CDec(txtDeta(5).Text))), "0") & "********"
    dsGirado2 = Trim(txtDato(0).Text)
    sDesBanco = Trim(lblDatoDeta(2).Caption)
    sCheque = Trim(txtDato(7).Text)
    sDia = Format(dtpFehBan.Value, "dd")
    sMes = Format(dtpFehBan.Value, "mm")
    sAno = Format(dtpFehBan.Value, "yyyy")
  End If
  
  
  ' Verifico el tipo de impresion
  sReporte = "rptEComPro": sTipo = "V"
  
  If MsgBox(Choose(gsIdioma, " Imprimir Cheque " & IIf(Index = 0, "Formato?", "Voucher?"), "Print Cheque " & IIf(Index = 0, "Format?", "Voucher?")), vbQuestion + vbYesNo + vbDefaultButton1, "Consulta") = vbYes Then
    sReporte = "rptECheVou"
    sTipo = "C"
    If Not dbHay104 Then
      MsgBox Choose(gsIdioma, "El comprobante no tiene alguna cuenta 104.", "The voucher doesn't have any account 104."), vbInformation
      porstMRp.Close
      Set porstMRp = Nothing
      Exit Sub
    End If
  End If
  If porstMRp.RecordCount > 0 Then nFormato = IIf(IsNull(porstMRp!forimp), 0, porstMRp!forimp)
  If Index = 0 Then
    sReporte = "rptECheVou" & Trim(IIf(sTipo = "C", nFormato, ""))
    dsFecha = IIf(sTipo = "C", sDia & sMes & sAno, dsFecha)
    ' Genero el reporte
    gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "LISTADO DE COMPROBANTES", "LISTING OF VOUCHERS"), udFecha, True, True, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & sReporte & ".rpt"
      '.WindowShowGroupTree = True
      '[Formulas adicionales
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
    sReporte = "rptECheVou" & Trim(IIf(sTipo = "C", nFormato, ""))
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
  
  If proceso = False Then
  
  With frmTBanGrd_anple                     'Cambiar Formulario de Grid.
    
    '[Llaves.                          'Cambiar
    txtLlave(0).MaxLength = .uorstMain_0!coddro.DefinedSize
    txtLlave(1).MaxLength = .uorstMain_0!nroban.DefinedSize
    ']
    
    '[Datos.                           'Cambiar.
    txtDato(gsIdioma - 1).MaxLength = .uorstMain_0!Globan.DefinedSize
    txtDato(2 - gsIdioma).MaxLength = .uorstMain_0!globanx.DefinedSize
    
    txtDato(2).MaxLength = .uorstMain_0!CodCta.DefinedSize
    txtDato(3).MaxLength = .uorstMain_0!codcco.DefinedSize
    txtDato(4).MaxLength = .uorstMain_0!codaux.DefinedSize
    txtDato(5).MaxLength = .uorstMain_0!CodFjo.DefinedSize
    txtDato(6).MaxLength = 8
    txtDato(7).MaxLength = .uorstMain_0!docban.DefinedSize
    txtDato(8).MaxLength = .uorstMain_0!portador.DefinedSize
    txtDato(9).MaxLength = .uorstMain_0!codbco.DefinedSize
    txtDato(10).MaxLength = .uorstMain_0!tpodoc.DefinedSize
    
    With dtpFehBan
      .MinDate = CDate("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct)
      .MaxDate = gfUltDia(.MinDate)
      .Value = .MaxDate
    End With
    With cboTpoTCb
      .AddItem TPOTCB_VTA_TXT, TPOTCB_VTA_IND
      .AddItem TPOTCB_CPR_TXT, TPOTCB_CPR_IND
    End With
    With cboTpoBan
      .AddItem TPOBAN_ING_TXT, TPOBAN_ING
      .AddItem TPOBAN_EGR_TXT, TPOBAN_EGR
    End With
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_2, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_2, TPOMON_EXT_IND
    End With
'    With cboTpoDoc
'      .AddItem "", 0
'      .AddItem TPODOC_DPS_TXT, TPODOC_DPS_IND
'      .AddItem TPODOC_GRO_TXT, TPODOC_GRO_IND
'      .AddItem TPODOC_TRA_TXT, TPODOC_TRA_IND
'      .AddItem TPODOC_ORD_TXT, TPODOC_ORD_IND
'      .AddItem TPODOC_DEB_TXT, TPODOC_DEB_IND
'      .AddItem TPODOC_CRE_TXT, TPODOC_CRE_IND
'      .AddItem TPODOC_CHQ_TXT, TPODOC_CHQ_IND
'      .AddItem TPODOC_OTR_TXT, TPODOC_OTR_IND
'      .AddItem TPODOC_EFE_TXT, TPODOC_EFE_IND
'      .AddItem TPODOC_PEX_TXT, TPODOC_PEX_IND
'      .AddItem TPODOC_LTR_TXT, TPODOC_LTR_IND
'      .AddItem TPODOC_CGE_TXT, TPODOC_CGE_IND
'    End With
    txtDeta(0).Text = Format(0, FORMATO_NUM_1)
    txtDeta(2).Text = Format(0, FORMATO_NUM_1)
    txtDeta(1).Text = Format(0, FORMATO_NUM_1)
    txtDeta(3).Text = Format(0, FORMATO_NUM_1)
    txtDeta(4).Text = Format(0, FORMATO_NUM_1)
    txtDeta(5).Text = Format(0, FORMATO_NUM_1)
    txtLlave(1).Enabled = False
    ']
    dgrDetalle.MarqueeStyle = dbgHighlightRow
    Set dgrDetalle.DataSource = .uorstMain_1
  End With
  Else
   With frmTBanGrd_anple                     'Cambiar Formulario de Grid.
    '[Llaves.                          'Cambiar
    txtLlave(0).MaxLength = .uorstMain_0Fil!coddro.DefinedSize
    txtLlave(1).MaxLength = .uorstMain_0Fil!nroban.DefinedSize
    ']
    
    '[Datos.                           'Cambiar.
    txtDato(gsIdioma - 1).MaxLength = .uorstMain_0Fil!Globan.DefinedSize
    txtDato(2 - gsIdioma).MaxLength = .uorstMain_0Fil!globanx.DefinedSize
    
    txtDato(2).MaxLength = .uorstMain_0Fil!CodCta.DefinedSize
    txtDato(3).MaxLength = .uorstMain_0Fil!codcco.DefinedSize
    txtDato(4).MaxLength = .uorstMain_0Fil!codaux.DefinedSize
    txtDato(5).MaxLength = .uorstMain_0Fil!CodFjo.DefinedSize
    txtDato(6).MaxLength = 8
    txtDato(7).MaxLength = .uorstMain_0Fil!docban.DefinedSize
    txtDato(8).MaxLength = .uorstMain_0Fil!portador.DefinedSize
    txtDato(9).MaxLength = .uorstMain_0Fil!codbco.DefinedSize
    txtDato(10).MaxLength = .uorstMain_0Fil!tpodoc.DefinedSize
    
    With dtpFehBan
      .MinDate = CDate("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct)
      .MaxDate = gfUltDia(.MinDate)
      .Value = .MaxDate
    End With
    With cboTpoTCb
      .AddItem TPOTCB_VTA_TXT, TPOTCB_VTA_IND
      .AddItem TPOTCB_CPR_TXT, TPOTCB_CPR_IND
    End With
    With cboTpoBan
      .AddItem TPOBAN_ING_TXT, TPOBAN_ING
      .AddItem TPOBAN_EGR_TXT, TPOBAN_EGR
    End With
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_2, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_2, TPOMON_EXT_IND
    End With
'    With cboTpoDoc
'      .AddItem "", 0
'      .AddItem TPODOC_DPS_TXT, TPODOC_DPS_IND
'      .AddItem TPODOC_GRO_TXT, TPODOC_GRO_IND
'      .AddItem TPODOC_TRA_TXT, TPODOC_TRA_IND
'      .AddItem TPODOC_ORD_TXT, TPODOC_ORD_IND
'      .AddItem TPODOC_DEB_TXT, TPODOC_DEB_IND
'      .AddItem TPODOC_CRE_TXT, TPODOC_CRE_IND
'      .AddItem TPODOC_CHQ_TXT, TPODOC_CHQ_IND
'      .AddItem TPODOC_OTR_TXT, TPODOC_OTR_IND
'      .AddItem TPODOC_EFE_TXT, TPODOC_EFE_IND
'      .AddItem TPODOC_PEX_TXT, TPODOC_PEX_IND
'      .AddItem TPODOC_LTR_TXT, TPODOC_LTR_IND
'      .AddItem TPODOC_CGE_TXT, TPODOC_CGE_IND
'    End With
    txtDeta(0).Text = Format(0, FORMATO_NUM_1)
    txtDeta(2).Text = Format(0, FORMATO_NUM_1)
    txtDeta(1).Text = Format(0, FORMATO_NUM_1)
    txtDeta(3).Text = Format(0, FORMATO_NUM_1)
    txtDeta(4).Text = Format(0, FORMATO_NUM_1)
    txtDeta(5).Text = Format(0, FORMATO_NUM_1)
    txtLlave(1).Enabled = False
    ']
    dgrDetalle.MarqueeStyle = dbgHighlightRow
    Set dgrDetalle.DataSource = .uorstMain_1Fil
  End With
  End If
   
  If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
  End If
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(14, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diario : ", "Nº Comprobante : ", "Fecha : ", "Glosa : ", "Traducción : ", "Cuenta:", "C.Costo : ", "Auxiliar : ", "F.Caja : ", "Tipo de Cambio : ", "Transacción : ", "Moneda Funcional : ", "Totales MN : ", "Totales ME : ")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journal : ", "Nº Voucher : ", "Date : ", "Gloss : ", "Translation : ", "Account : ", "C.Center : ", "Auxiliary : ", "Cash F. : ", "Rate of Exchange:", "Transaction : ", "Functional Currency : ", "Totals NC : ", "Totals FC : ")
  Next nElemento
  chkGenCpb.Caption = Choose(gsIdioma, "Comprobante Diario", "Voucher Journal")
  chkGenPrc.Caption = Choose(gsIdioma, "Procesado", "Processing")
  fraDocumento.Caption = Choose(gsIdioma, " Documento ", " Document ")
  cmdCalcular.Caption = Choose(gsIdioma, "Calc&ular", "Calc&ulate")
  cmdAuxiliar.Caption = Choose(gsIdioma, "&Auxiliar", "&Auxiliary")
  CaptionBotones Me, False, False, True, True, True, True, True, True, False, True, True, True, True, aLabel
  ']
   
  
   
End Sub

Private Sub Form_Activate()
  '[Busca detalle de códigos.           'Cambiar (habilitar/deshabilitar).
  If txtDato(3).Text <> "" Then ppAyuDet AYUDAT, 3
  If txtDato(4).Text <> "" Then ppAyuDet AYUDAT, 4
  If txtDato(5).Text <> "" Then ppAyuDet AYUDAT, 5
  If txtDato(9).Text <> "" Then ppAyuDet AYUDAT, 9
  If txtDato(10).Text <> "" Then ppAyuDet AYUDAT, 10
  If txtDato(2).Text <> "" Then
    ppAyuDet AYUDAT, 2
    pnCta_IndCCo = frmTBanGrd_anple.uorstCoCta!indcco
    psCodCCo_Def = IIf(IsNull(frmTBanGrd_anple.uorstCoCta!codcco_def), "", frmTBanGrd_anple.uorstCoCta!codcco_def)
    ' Actualiza los datos de centro de costo
    txtDato(3).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(2).Enabled)
    cmdDatoAyud(3).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(2).Enabled)
  End If

  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
  ' Modifico habiltar/deshabilitar  comandos
  pbGraba = Not pbNuevo
  ppbotones_tpognr0 pbGraba
  pbGraba = False
   
  If pbLoad And Not pbNuevo Then
    pbLoad = False
    frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
    If proceso = False Then
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.coddro='" & frmTBanGrd_anple.uorstMain_0!coddro & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & frmTBanGrd_anple.uorstMain_0!nroban & "' "
    Else
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.coddro='" & frmTBanGrd_anple.uorstMain_0Fil!coddro & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & frmTBanGrd_anple.uorstMain_0Fil!nroban & "' "
    
    End If
    ppDatosWhere
  End If
  upDatosGrid

  '[Propio del Formulario.
  If Month(dtpFehBan.Value) <> Val(gfMesAct(gsMesAct)) Or Year(dtpFehBan.Value) <> Val(gsAnoAct) Then
    cmdCorregir.Enabled = False
    cmdNuevo.Enabled = False
    cmdEliminar.Enabled = False
  End If
  If txtLlave(0).Text <> "" Then ppAyuDet AYULLA, 0
  ']
  
  'If pbNuevo Then
  '  chkGenCpb.Value = Unchecked
  'End If
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If proceso = False Then
   If frmTBanGrd_anple.uorstMain_0.RecordCount <> 0 Then
'      frmTBanGrd_anple.uorstMain_0.CancelUpdate   'Cambiar Formulario de Grid.
   End If
Else
   If frmTBanGrd_anple.uorstMain_0Fil.RecordCount <> 0 Then
'      frmTBanGrd_anple.uorstMain_0.CancelUpdate   'Cambiar Formulario de Grid.
   End If
End If
End Sub

Private Sub cmdRetroceder_Click()
  If proceso = False Then
    gpTUe_Retroceder frmTBanGrd_anple.uorstMain_0, Me
    frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.CodDro='" & frmTBanGrd_anple.uorstMain_0!coddro & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & frmTBanGrd_anple.uorstMain_0!nroban & "' "
    ppDatosWhere
    ppbotones_tpognr0 (pbGraba)
    
    ' Busca ítem MA
    frmTBanGrd_anple.uorstMain_Grd.MoveFirst
    frmTBanGrd_anple.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
 Else
    gpTUe_Retroceder frmTBanGrd_anple.uorstMain_0Fil, Me
    frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.CodDro='" & frmTBanGrd_anple.uorstMain_0Fil!coddro & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & frmTBanGrd_anple.uorstMain_0Fil!nroban & "'"
    ppDatosWhere
    ppbotones_tpognr0 (pbGraba)
    
    ' Busca ítem MA
    frmTBanGrd_anple.uorstMain_GrdFil.MoveFirst
    frmTBanGrd_anple.uorstMain_GrdFil.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
 End If
End Sub

Private Sub cmdAvanzar_Click()
  If proceso = False Then
  
    gpTUe_Avanzar frmTBanGrd_anple.uorstMain_0, Me
    frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.CodDro='" & frmTBanGrd_anple.uorstMain_0!coddro & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & frmTBanGrd_anple.uorstMain_0!nroban & "' "
    ppDatosWhere
    ppbotones_tpognr0 (pbGraba)
    
    ' Busca ítem MA
    frmTBanGrd_anple.uorstMain_Grd.MoveFirst
    frmTBanGrd_anple.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
  Else
 
    gpTUe_Avanzar frmTBanGrd_anple.uorstMain_0Fil, Me
    frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.CodDro='" & frmTBanGrd_anple.uorstMain_0Fil!coddro & "' "
    frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & frmTBanGrd_anple.uorstMain_0Fil!nroban & "'"
    ppDatosWhere
    ppbotones_tpognr0 (pbGraba)
    
    ' Busca ítem MA
    frmTBanGrd_anple.uorstMain_GrdFil.MoveFirst
    frmTBanGrd_anple.uorstMain_GrdFil.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
  
  End If
End Sub

Public Sub cmdCorregir_Click()
   'Verificación de Mes Cerrado.
   If gbCieCpb Then MsgBox TEXT_9016, vbCritical: Exit Sub
   
   pbNuevo = False
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   ' Inhabilita botones
   cmdNuevo.Enabled = False
   cmdRevisar.Enabled = False
   cmdEliminar.Enabled = False
   cmdRefrescar.Enabled = False
   cmdImprimir(0).Enabled = False
   cmdImprimir(1).Enabled = False
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
  Dim sExpresion As String
  
  On Error GoTo Err
  '[Validacion de Datos
  If Trim(txtDato(0).Text) = "" Then MsgBox TEXT_8005, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If Len(Trim(txtDato(2).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(2).SetFocus: Exit Sub
  If pnCta_IndCCo = INDCCO_ACT And Len(Trim(txtDato(3).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(1).SetFocus: Exit Sub
  If Len(Trim(txtDato(4).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(4).SetFocus: Exit Sub
  If Len(Trim(txtDato(5).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(5).SetFocus: Exit Sub
  If CDec(txtDato(6).Text) <= 0 Then MsgBox TEXT_9015, vbExclamation: txtDato(6).SetFocus: Exit Sub
  'If cboTpoDoc.Text = "" Then MsgBox TEXT_8005, vbExclamation: cboTpoDoc.SetFocus: Exit Sub
  If txtDato(10).Text = "" Then MsgBox TEXT_8005, vbExclamation: txtDato(10).SetFocus: Exit Sub
  If Trim(txtDato(7).Text) = "" Then MsgBox TEXT_8005, vbExclamation: txtDato(7).SetFocus: Exit Sub
  ' verifico si comprobante existe
  If pbNuevo Then
    sExpresion = "SELECT COUNT(NroCpb) AS cNumeroCpb "
    sExpresion = sExpresion & "FROM CoCpbCab "
    sExpresion = sExpresion & "WHERE codemp='" & gsCodEmp & "' "
    sExpresion = sExpresion & "AND pdoano='" & gsAnoAct & "' "
    sExpresion = sExpresion & "AND MesPvs='" & gsMesAct & "' "
    sExpresion = sExpresion & "AND CodDro='" & txtLlave(0).Text & "' "
    sExpresion = sExpresion & "AND nrocpb='" & txtLlave(1).Text & "'"
    If gfRetornaValor(CONNSTRG & gsNomBDS, sExpresion) = "1" Then
      If MsgBox(Choose(gsIdioma, "Comprobante ya existe. Desea actualizar", "The voucher already exists. To update"), vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
        txtLlave(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtLlave(0).Text)
      Else
         txtLlave(1).SetFocus: Exit Sub
      End If
    End If
  End If
  
  With frmTBanGrd_anple                     'Cambiar Formulario de Grid.
    .uocnnMain.BeginTrans            'INICIA TRANSACCION.
    '[ No Pertenece al Formulario
    If pbNuevo Then
      pbGraba = True
      ' MA para conservar el numero digitado por el usuario
      .uorstCODro.Fields("Cpb" & gsMesAct).Value = txtLlave(1).Text
      .uorstCODro.Update
    End If
    ']
    If pbNuevo Then
      If proceso = False Then
      .uorstMain_0.AddNew
      Else
      .uorstMain_0Fil.AddNew
      End If
    End If
    upDatosDesconectados 0
    If proceso = False Then
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
    Else
    With .uorstMain_0Fil
      If pbNuevo Then
        !UsrCre = gsAbvUsr
        !FyHCre = Now
      Else
        !UsrMdf = gsAbvUsr
        !FyHMdf = Now
      End If
      .Update
    End With
    
    End If
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    If proceso = False Then
    ' [Actualiza grilla y busco registro actual
    .uorstMain_Grd.Requery
    .ppDatosGrid
    .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
    ' ]
    Else
    ' [Actualiza grilla y busco registro actual
    .uorstMain_GrdFil.Requery
    .ppDatosGrid
    .uorstMain_GrdFil.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
    ' ]
    
    End If
    
    If pbNuevo Then
      cmdGrabar.Enabled = False
      upHabilitacion False
      txtLlave(1).Enabled = False
      cmdLlaveAyud(0).Enabled = False
      dtpFehBan.Enabled = False
      cboTpoTCb.Enabled = False
      cboTpoTCb.Enabled = False
      txtDato(4).Enabled = False
      txtDato(6).Enabled = False
      cmdDatoAyud(4).Enabled = False
      pbNuevo = False
   
      '[ No Pertenece al Formulario
      
      If proceso = False Then
        .uorstMain_0.Requery
      Else
        .uorstMain_0Fil.Requery
      End If
      
      cmdNuevo_Click
      cmdRetroceder.Enabled = True
      cmdAvanzar.Enabled = True
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      ' Habilita botones
      cmdNuevo.Enabled = True
      cmdRevisar.Enabled = True
      cmdEliminar.Enabled = True
      cmdRefrescar.Enabled = True
      cmdImprimir(0).Enabled = True
      cmdImprimir(1).Enabled = True
      frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
      frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
      frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
      frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.CodDro='" & txtLlave(0).Text & "' "
      frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & txtLlave(1).Text & "' "
      ppDatosWhere
      ']
      
    Else
      cmdRetroceder.Enabled = True
      cmdAvanzar.Enabled = True
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      ' Habilita botones
      cmdNuevo.Enabled = True
      cmdRevisar.Enabled = True
      cmdEliminar.Enabled = True
      cmdRefrescar.Enabled = True
      cmdImprimir(0).Enabled = True
      cmdImprimir(1).Enabled = True
      upHabilitacion False
    End If
  End With
  cmdNuevo.SetFocus
  
'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
fEstMayUpd
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion

  Exit Sub
Err:
  gpErrores
  frmTBanGrd_anple.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.

End Sub

Public Sub cmdDeshacer_Click()
  gpTUe_Deshacer Me
  ' Habilita botones
  cmdNuevo.Enabled = True
  cmdRevisar.Enabled = True
  cmdEliminar.Enabled = True
  cmdRefrescar.Enabled = True
  cmdImprimir(0).Enabled = True
  cmdImprimir(1).Enabled = True
End Sub

Public Sub cmdNuevo_Click()
  'Verificación de Mes Cerrado.
  If gbCieCpb Then
    MsgBox TEXT_9016, vbCritical
    Exit Sub
  End If
   
  If proceso = False Then
  With frmTBanGrd_anple.uorstMain_0
    If .RecordCount > 0 Then
      .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
      If .EOF Then
        MsgBox Choose(gsIdioma, "Tiene que Grabar la Cabecera del Comprobante para poder Registrar el Detalle", "You have to save Voucher Header to register Detail"), vbInformation
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
        Exit Sub
      End If
    End If
  End With
  Else
  With frmTBanGrd_anple.uorstMain_0Fil
    If .RecordCount > 0 Then
      .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
      If .EOF Then
        MsgBox Choose(gsIdioma, "Tiene que Grabar la Cabecera del Comprobante para poder Registrar el Detalle", "You have to save Voucher Header to register Detail"), vbInformation
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
        Exit Sub
      End If
    End If
  End With
  
  End If
   
  gpTVd_Nuevo Me, frmTBanDet_anple
  '[Agregado por Angel
  frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.coddro='" & txtLlave(0).Text & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & txtLlave(1).Text & "' "
  ppDatosWhere
  If proceso = False Then
  If frmTBanGrd_anple.uorstMain_1.RecordCount > 0 Then
    frmTBanGrd_anple.uorstMain_1.MoveLast
  End If
  Else
  If frmTBanGrd_anple.uorstMain_1Fil.RecordCount > 0 Then
    frmTBanGrd_anple.uorstMain_1Fil.MoveLast
  End If
  
  End If
End Sub

Public Sub cmdRevisar_click()
  On Error GoTo Err
  Dim dvRegistro As Variant
  'Verificación de ítemes creados.     'Cambiar Formulario de Grid.
  If proceso = False Then
    If frmTBanGrd_anple.uorstMain_1.RecordCount = 0 Then
        MsgBox TEXT_8001, vbInformation
        Exit Sub
    End If
  Else
    If frmTBanGrd_anple.uorstMain_1Fil.RecordCount = 0 Then
        MsgBox TEXT_8001, vbInformation
        Exit Sub
    End If
  End If
  
  'With frmTFacDet                     'Cambiar Formulario de Datos.
  With frmTBanDet_anple
    .zbNuevo = False
    .upDatosDesconectados 1
    .Caption = TEXT_MODIF & " " & Me.Caption
    .Show vbModal
  End With
  '[Agregado por Angel
  If proceso = False Then
    dvRegistro = frmTBanGrd_anple.uorstMain_1.Bookmark
  Else
    dvRegistro = frmTBanGrd_anple.uorstMain_1Fil.Bookmark
  End If
  
  frmTBanGrd_anple.usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.MesPvs='" & gsMesAct & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.CodDro='" & txtLlave(0).Text & "' "
  frmTBanGrd_anple.usConnStrgWher_1 = frmTBanGrd_anple.usConnStrgWher_1 & "AND cobandet.nroban='" & txtLlave(1).Text & "' "
  ppDatosWhere
  If proceso = False Then
        frmTBanGrd_anple.uorstMain_1.Bookmark = dvRegistro
  Else
        frmTBanGrd_anple.uorstMain_1Fil.Bookmark = dvRegistro
  End If
  ']
  dgrDetalle.SetFocus
  Exit Sub
Err:
  gpErrores
End Sub

Public Sub cmdEliminar_Click()
  On Error GoTo Err

  Dim dnBlqIte As Integer
  Dim dvRegistro As Variant

  With frmTBanGrd_anple                     'Cambiar Formulario de Grid.
    'Verificaciones.
    If gbCieCpb Then                 'Mes Cerrado.
      MsgBox TEXT_9016, vbCritical
      Exit Sub
    End If
    If proceso = False Then
    If .uorstMain_1.RecordCount = 0 Then
      MsgBox TEXT_8001, vbInformation
      Exit Sub
    ElseIf .uorstMain_1.BOF Then
      .uorstMain_1.MoveNext
    ElseIf .uorstMain_1.EOF Then
      .uorstMain_1.MovePrevious
    End If
    Else
    If .uorstMain_1Fil.RecordCount = 0 Then
      MsgBox TEXT_8001, vbInformation
      Exit Sub
    ElseIf .uorstMain_1Fil.BOF Then
      .uorstMain_1Fil.MoveNext
    ElseIf .uorstMain_1Fil.EOF Then
      .uorstMain_1Fil.MovePrevious
    End If
    End If
    
    
    'Confirmación                     'Cambiar.
    If MsgBox(TEXT_1021 & " " & Trim(dgrDetalle.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      
      If proceso = False Then
      
      If Trim(.uorstMain_1!tpoban) = TPOBAN_EGR Then
        frmTBanCab_anple.txtDeta(0) = Format(CDec(frmTBanCab_anple.txtDeta(0).Text) - .uorstMain_1!ImpMN, FORMATO_NUM_1)
        frmTBanCab_anple.txtDeta(2) = Format(CDec(frmTBanCab_anple.txtDeta(2).Text) - .uorstMain_1!ImpME, FORMATO_NUM_1)
      Else
        frmTBanCab_anple.txtDeta(1) = Format(CDec(frmTBanCab_anple.txtDeta(1).Text) - .uorstMain_1!ImpMN, FORMATO_NUM_1)
        frmTBanCab_anple.txtDeta(3) = Format(CDec(frmTBanCab_anple.txtDeta(3).Text) - .uorstMain_1!ImpME, FORMATO_NUM_1)
      End If
      frmTBanCab_anple.txtDeta(4) = Format(CDec(frmTBanCab_anple.txtDeta(1).Text) - CDec(frmTBanCab_anple.txtDeta(0).Text), FORMATO_NUM_1)
      frmTBanCab_anple.txtDeta(5) = Format(CDec(frmTBanCab_anple.txtDeta(3).Text) - CDec(frmTBanCab_anple.txtDeta(2).Text), FORMATO_NUM_1)
      .uocnnMain.BeginTrans     ' Inicio transaccion
      dvRegistro = .uorstMain_1.Bookmark
      If Not .uorstMain_1.RecordCount = 0 Then
        dnBlqIte = .uorstMain_1!nroitem
        .uorstMain_1.MoveFirst
        Do
          If Trim(.uorstMain_1!nroitem) = dnBlqIte Then
            ' Actualizo los importes de la cabecera
            .uorstMain_0!ImpMN = Round(.uorstMain_0!ImpMN - (.uorstMain_1!ImpMN * IIf(Trim(.uorstMain_1!tpoban) = TPOBAN_ING, 1, -1)), 2)
            .uorstMain_0!ImpME = Round(.uorstMain_0!ImpME - (.uorstMain_1!ImpME * IIf(Trim(.uorstMain_1!tpoban) = TPOBAN_ING, 1, -1)), 2)
            .uorstMain_0.Update
            ' Elimino el registro
            .uorstMain_1.Delete
          End If
          .uorstMain_1.MoveNext
        Loop Until .uorstMain_1.EOF
      End If
      If .uorstMain_1.RecordCount > 0 Then
        If dvRegistro > .uorstMain_1.RecordCount Then
          .uorstMain_1.MoveLast
        Else
          .uorstMain_1.Bookmark = dvRegistro
        End If
      End If
      .uocnnMain.CommitTrans    ' Fin alizo la transaccion
      
      Else
      
      If Trim(.uorstMain_1Fil!tpoban) = TPOBAN_EGR Then
        frmTBanCab_anple.txtDeta(0) = Format(CDec(frmTBanCab_anple.txtDeta(0).Text) - .uorstMain_1Fil!ImpMN, FORMATO_NUM_1)
        frmTBanCab_anple.txtDeta(2) = Format(CDec(frmTBanCab_anple.txtDeta(2).Text) - .uorstMain_1Fil!ImpME, FORMATO_NUM_1)
      Else
        frmTBanCab_anple.txtDeta(1) = Format(CDec(frmTBanCab_anple.txtDeta(1).Text) - .uorstMain_1Fil!ImpMN, FORMATO_NUM_1)
        frmTBanCab_anple.txtDeta(3) = Format(CDec(frmTBanCab_anple.txtDeta(3).Text) - .uorstMain_1Fil!ImpME, FORMATO_NUM_1)
      End If
      frmTBanCab_anple.txtDeta(4) = Format(CDec(frmTBanCab_anple.txtDeta(1).Text) - CDec(frmTBanCab_anple.txtDeta(0).Text), FORMATO_NUM_1)
      frmTBanCab_anple.txtDeta(5) = Format(CDec(frmTBanCab_anple.txtDeta(3).Text) - CDec(frmTBanCab_anple.txtDeta(2).Text), FORMATO_NUM_1)
      .uocnnMain.BeginTrans     ' Inicio transaccion
      dvRegistro = .uorstMain_1Fil.Bookmark
      If Not .uorstMain_1Fil.RecordCount = 0 Then
        dnBlqIte = .uorstMain_1Fil!nroitem
        .uorstMain_1Fil.MoveFirst
        Do
          If Trim(.uorstMain_1Fil!nroitem) = dnBlqIte Then
            ' Actualizo los importes de la cabecera
            .uorstMain_0Fil!ImpMN = Round(.uorstMain_0!ImpMN - (.uorstMain_1Fil!ImpMN * IIf(Trim(.uorstMain_1Fil!tpoban) = TPOBAN_ING, 1, -1)), 2)
            .uorstMain_0Fil!ImpME = Round(.uorstMain_0!ImpME - (.uorstMain_1Fil!ImpME * IIf(Trim(.uorstMain_1Fil!tpoban) = TPOBAN_ING, 1, -1)), 2)
            .uorstMain_0Fil.Update
            ' Elimino el registro
            .uorstMain_1Fil.Delete
          End If
          .uorstMain_1Fil.MoveNext
        Loop Until .uorstMain_1Fil.EOF
      End If
      If .uorstMain_1Fil.RecordCount > 0 Then
        If dvRegistro > .uorstMain_1Fil.RecordCount Then
          .uorstMain_1Fil.MoveLast
        Else
          .uorstMain_1Fil.Bookmark = dvRegistro
        End If
      End If
      .uocnnMain.CommitTrans    ' Fin alizo la transaccion
      
      
      End If
      
    End If
    dgrDetalle.SetFocus
   End With

   Exit Sub
Err:
   gpErrores
   
   frmTBanGrd_anple.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
  '[ Datos No Pertenecen al Formulario - Agregado por Angel
  If proceso = False Then
    frmTBanGrd_anple.uorstMain_1.Requery
  Else
    frmTBanGrd_anple.uorstMain_1Fil.Requery
  End If
  frmTBanCab_anple.upDatosGrid
  frmTBanCab_anple.dgrDetalle.SetFocus
  ']
End Sub

Public Sub cmdCalcular_Click()
  If proceso = False Then
  With frmTBanGrd_anple.uorstMain_1
    txtDeta(0).Text = Format(0, FORMATO_NUM_1)
    txtDeta(1).Text = Format(0, FORMATO_NUM_1)
    txtDeta(2).Text = Format(0, FORMATO_NUM_1)
    txtDeta(3).Text = Format(0, FORMATO_NUM_1)
    txtDeta(4).Text = Format(0, FORMATO_NUM_1)
    txtDeta(5).Text = Format(0, FORMATO_NUM_1)
    If frmTBanGrd_anple.uorstMain_1.RecordCount > 0 Then
      frmTBanGrd_anple.uorstMain_1.MoveFirst
      Do
        txtDeta(0).Text = Format(CDec(txtDeta(0).Text) + !cImpMN_Deb, FORMATO_NUM_1)
        txtDeta(1).Text = Format(CDec(txtDeta(1).Text) + !cImpMN_Hab, FORMATO_NUM_1)
        txtDeta(2).Text = Format(CDec(txtDeta(2).Text) + !cImpME_Deb, FORMATO_NUM_1)
        txtDeta(3).Text = Format(CDec(txtDeta(3).Text) + !cImpME_Hab, FORMATO_NUM_1)
        frmTBanGrd_anple.uorstMain_1.MoveNext
      Loop Until .EOF
    End If
  End With
  txtDeta(4).Text = Format(CDec(txtDeta(1).Text) - CDec(txtDeta(0).Text), FORMATO_NUM_1)
  txtDeta(5).Text = Format(CDec(txtDeta(3).Text) - CDec(txtDeta(2).Text), FORMATO_NUM_1)
  Set dgrDetalle.DataSource = frmTBanGrd_anple.uorstMain_1
  Else
  With frmTBanGrd_anple.uorstMain_1Fil
    txtDeta(0).Text = Format(0, FORMATO_NUM_1)
    txtDeta(1).Text = Format(0, FORMATO_NUM_1)
    txtDeta(2).Text = Format(0, FORMATO_NUM_1)
    txtDeta(3).Text = Format(0, FORMATO_NUM_1)
    txtDeta(4).Text = Format(0, FORMATO_NUM_1)
    txtDeta(5).Text = Format(0, FORMATO_NUM_1)
    If frmTBanGrd_anple.uorstMain_1Fil.RecordCount > 0 Then
      frmTBanGrd_anple.uorstMain_1Fil.MoveFirst
      Do
        txtDeta(0).Text = Format(CDec(txtDeta(0).Text) + !cImpMN_Deb, FORMATO_NUM_1)
        txtDeta(1).Text = Format(CDec(txtDeta(1).Text) + !cImpMN_Hab, FORMATO_NUM_1)
        txtDeta(2).Text = Format(CDec(txtDeta(2).Text) + !cImpME_Deb, FORMATO_NUM_1)
        txtDeta(3).Text = Format(CDec(txtDeta(3).Text) + !cImpME_Hab, FORMATO_NUM_1)
        frmTBanGrd_anple.uorstMain_1Fil.MoveNext
      Loop Until .EOF
    End If
  End With
  txtDeta(4).Text = Format(CDec(txtDeta(1).Text) - CDec(txtDeta(0).Text), FORMATO_NUM_1)
  txtDeta(5).Text = Format(CDec(txtDeta(3).Text) - CDec(txtDeta(2).Text), FORMATO_NUM_1)
  Set dgrDetalle.DataSource = frmTBanGrd_anple.uorstMain_1Fil
  
  End If
  upDatosGrid
End Sub

Private Sub cmdSalir_Click()
  ' Genera comprobante contabilidad
  ppGeneraComprobante
  
  If salirform = True Then
    Activar = False
    Unload Me
  End If
  
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0
    txtLlave(Index).SetFocus
  End Select
  ppAyuBus AYULLA, Index
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 2, 3, 4, 5, 10
    txtDato(Index).SetFocus
  End Select
  ppAyuBus AYUDAT, Index
End Sub

Private Sub txtllave_GotFocus(Index As Integer)
  txtLlave(Index).SelStart = 0
  txtLlave(Index).SelLength = txtLlave(Index).MaxLength
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
    cboTpoBan.Enabled = False
    txtLlave(0).Enabled = False
    cmdLlaveAyud(0).Enabled = False
    lblLlaveDeta(0).Enabled = False
    dtpFehBan.SetFocus
  End If
  
  If Index = 1 Then
  
  txtDato(0).SetFocus
  
  End If
  
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  Dim dbSalir As Boolean
  Dim dvRegistro As Variant
  
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
  Select Case Index                    'Cambiar (añadir índices).
   Case 1
    If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
      txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
    End If
  End Select
  
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
  Select Case Index                    'Cambiar (añadir índices).
   Case 0
    Cancel = ppAyuDet(AYULLA, Index)
    If Cancel Then Exit Sub
    If Len(txtLlave(0)) <> 4 Then txtLlave(0).SetFocus: Exit Sub
  End Select
 
  'Captura del Siguiente Número.       'Cambiar (Activar/Inactivar).
  If Index = 0 And Len(Trim(txtLlave.Item(0).Text)) <> 0 Then
    If pbNuevo Then
      With frmTBanGrd_anple.uorstCODro
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
    If proceso = False Then
    With frmTBanGrd_anple.uorstMain_0
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
    Else
     With frmTBanGrd_anple.uorstMain_0Fil
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
    End If
    
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
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYUDAT, Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  ' Busca el dato en su tabla principal.
  Select Case Index
   Case 2                           'Cambiar (añadir índices).
    
    Cancel = ppAyuDet(AYUDAT, Index)
    
    If Cancel Then Exit Sub
    
    If lblDatoDeta(Index).Caption <> "" Then
    
        pnCta_IndCCo = frmTBanGrd_anple.uorstCoCta!indcco
        psCodCCo_Def = IIf(IsNull(frmTBanGrd_anple.uorstCoCta!codcco_def), "", frmTBanGrd_anple.uorstCoCta!codcco_def)
      
      'Actualizo los datos adicionales
      
      'If psCodCCo_Def <> "" Then
      
        txtDato(3).Text = ""
        txtDato(3).Text = IIf(txtDato(3).Text = "", psCodCCo_Def, txtDato(3).Text)
        txtDato(3).Text = IIf(pnCta_IndCCo = INDCCO_ACT, txtDato(3).Text, "")
        lblDatoDeta(3).Caption = IIf(pnCta_IndCCo = INDCCO_ACT, lblDatoDeta(3).Caption, "")
        txtDato(3).Enabled = (pnCta_IndCCo = INDCCO_ACT)
        cmdDatoAyud(3).Enabled = (pnCta_IndCCo = INDCCO_ACT)
        cboTpoMon.ListIndex = IIf(frmTBanGrd_anple.uorstCoCta!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
        
      'End If
    
        txtDato(9).Text = IIf(IsNull(frmTBanGrd_anple.uorstCoCta!codbco), "", frmTBanGrd_anple.uorstCoCta!codbco)
    
    End If
    
    ppAyuDet AYUDAT, 9
   
   Case 3, 4, 5, 9, 10
   
    Cancel = ppAyuDet(AYUDAT, Index)
    
    If Cancel Then Exit Sub
    If Index = 4 Then
      txtDato(8).Text = Trim(IIf(pbNuevo Or (txtDato(8).Text = "" And Not pbNuevo), lblDatoDeta(Index).Caption, txtDato(8).Text))
    End If
   Case 6
    txtDato(Index).Text = IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)
    txtDato(Index).Text = Format(CDec(txtDato(Index).Text), FORMATO_NUM_2)
  End Select
  
  Exit Sub
Err:
  gpErrores
End Sub

Private Sub dgrDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
   If proceso = False Then
   If frmTBanGrd_anple.uorstMain_1.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      frmTBanGrd_anple.uorstMain_1.MoveFirst
   Case vbKeyEnd
      frmTBanGrd_anple.uorstMain_1.MoveLast
   End Select
   Else
   If frmTBanGrd_anple.uorstMain_1Fil.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      frmTBanGrd_anple.uorstMain_1Fil.MoveFirst
   Case vbKeyEnd
      frmTBanGrd_anple.uorstMain_1Fil.MoveLast
   End Select
   
   End If
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
  Dim sCondicion As String
  
   
  If tsTipo = AYULLA Then
    Select Case tnIndex
     Case 0                           'Cambiar (añadir índices).
      sCondicion = IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(coddro)=4 AND LEFT(coddro, " & Len(Choose(cboTpoBan.ListIndex + 1, gsCodDro_Ing, gsCodDro_Egr)) & ")='" & Choose(cboTpoBan.ListIndex + 1, gsCodDro_Ing, gsCodDro_Egr) & "' "
      modAyuBus.Dro_Cod sCondicion, txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
      txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
      lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  Else
    Select Case tnIndex
     Case 2             ' Cuenta caja bancos
      modAyuBus.Cta_Cod "tpocta=" & TPOCTA_TRA & " AND estcta='" & ESTCTA_ACT & "' AND LEFT(codcta, 2)='10' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 3             ' Centro de costos
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(codcco)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 4             ' Auxiliares
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 5             ' Flujo de caja
      modAyuBus.Fjo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodFjo)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(4).Caption = " " & frmOAyuBus.uvDato2
     Case 10             ' Medio de Pago
      modAyuBus.Med_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(10).Caption = " " & frmOAyuBus.uvDato2
      xIndicador = frmOAyuBus.uvDato4
    End Select
  End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
  Dim sDiario As String
  
   
  If tsTipo = AYULLA Then
    Select Case tnIndex                 'Cambiar.
     Case 0
      If txtLlave(tnIndex).Text = "" Then
        lblLlaveDeta(tnIndex).Caption = ""
        Exit Function
      End If
      sDiario = Choose(cboTpoBan.ListIndex + 1, gsCodDro_Ing, gsCodDro_Egr)
      If Left(txtLlave(tnIndex).Text, Len(sDiario)) <> sDiario Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
        Exit Function
      End If
      With frmTBanGrd_anple.uorstCODro
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
    Select Case tnIndex                 'Cambiar.
     Case 2
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTBanGrd_anple.uorstCoCta
        If .RecordCount > 0 Then .MoveFirst
        .Find "codcta='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
        End If
      End With
     Case 3
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTBanGrd_anple.uorstCoCCo
        If .RecordCount > 0 Then .MoveFirst
        .Find "codcco='" & txtDato(tnIndex).Text & "'"
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
      With frmTBanGrd_anple.uorstTGAux
        If .RecordCount > 0 Then .MoveFirst
        .Find "codaux='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!razAux), "", !razAux)
        End If
      End With
     Case 5
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTBanGrd_anple.uorstCOFjo
        If .RecordCount > 0 Then .MoveFirst
        .Find "codfjo='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetFjo), "", !DetFjo)
        End If
      End With
     Case 9
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTBanGrd_anple.uorstCoBco
        If .RecordCount > 0 Then .MoveFirst
        .Find "codbco='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detbco), "", !detbco)
        End If
      End With
    Case 10
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      
      With frmTBanGrd_anple.uorstmedio
      
        If .RecordCount > 0 Then .MoveFirst
        .Find "codmed='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!abvmed), "", !abvmed)
          
          xIndicador = IIf(IsNull(!indmod), "", !indmod)
          
        End If
      End With
    End Select
   End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
  '[Propio del formulario.
  Dim svValores As Variant
  ']
  On Error GoTo Err
  
  With frmTBanGrd_anple                     'Cambiar Formulario de Grid.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
      If proceso = False Then
        .uorstMain_0!codemp = gsCodEmp
        .uorstMain_0!pdoano = gsAnoAct
        .uorstMain_0!mespvs = gsMesAct
        .uorstMain_0!coddro = txtLlave(0).Text
        .uorstMain_0!nroban = txtLlave(1).Text
      Else
        .uorstMain_0Fil!codemp = gsCodEmp
        .uorstMain_0Fil!pdoano = gsAnoAct
        .uorstMain_0Fil!mespvs = gsMesAct
        .uorstMain_0Fil!coddro = txtLlave(0).Text
        .uorstMain_0Fil!nroban = txtLlave(1).Text
      End If
      End If
      ' reemplazo los caracteres
      txtDato(0).Text = gfSacaEntRetApos(txtDato(0).Text)
      txtDato(1).Text = gfSacaEntRetApos(txtDato(1).Text)
      txtDato(8).Text = gfSacaEntRetApos(txtDato(8).Text)
      
      If proceso = False Then
      'Datos.
      .uorstMain_0!fehban = dtpFehBan.Value
      .uorstMain_0!Globan = txtDato(gsIdioma - 1).Text
      .uorstMain_0!globanx = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
      .uorstMain_0!tpoban = cboTpoBan.ListIndex
      .uorstMain_0!CodCta = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      .uorstMain_0!codcco = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
      .uorstMain_0!codaux = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
      .uorstMain_0!CodFjo = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
      '.uorstMain_0!tpodoc = cboTpoDoc.ListIndex
      .uorstMain_0!tpodoc = txtDato(10).Text
      .uorstMain_0!docban = IIf(txtDato(7).Text = "", Null, txtDato(7).Text)
      .uorstMain_0!codbco = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
      .uorstMain_0!portador = IIf(txtDato(8).Text = "", Null, txtDato(8).Text)
      .uorstMain_0!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      .uorstMain_0!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
      .uorstMain_0!ImpTCb = CDec(txtDato(6).Text)
      .uorstMain_0!ImpMN = CDec(txtDeta(4).Text)
      .uorstMain_0!ImpME = CDec(txtDeta(5).Text)
      .uorstMain_0!gencpb = IIf(chkGenCpb.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      .uorstMain_0!genprc = IIf(chkGenPrc.Value = vbChecked, INDPREGEN_ACTx, INDPREGEN_INAx)
      Else
      'Datos.
      .uorstMain_0Fil!fehban = dtpFehBan.Value
      .uorstMain_0Fil!Globan = txtDato(gsIdioma - 1).Text
      .uorstMain_0Fil!globanx = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
      .uorstMain_0Fil!tpoban = cboTpoBan.ListIndex
      .uorstMain_0Fil!CodCta = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      .uorstMain_0Fil!codcco = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
      .uorstMain_0Fil!codaux = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
      .uorstMain_0Fil!CodFjo = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
      '.uorstMain_0Fil!tpodoc = cboTpoDoc.ListIndex
      .uorstMain_0Fil!tpodoc = txtDato(10).Text
      .uorstMain_0Fil!docban = IIf(txtDato(7).Text = "", Null, txtDato(7).Text)
      .uorstMain_0Fil!codbco = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
      .uorstMain_0Fil!portador = IIf(txtDato(8).Text = "", Null, txtDato(8).Text)
      .uorstMain_0Fil!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      .uorstMain_0Fil!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
      .uorstMain_0Fil!ImpTCb = CDec(txtDato(6).Text)
      .uorstMain_0Fil!ImpMN = CDec(txtDeta(4).Text)
      .uorstMain_0Fil!ImpME = CDec(txtDeta(5).Text)
      .uorstMain_0Fil!gencpb = IIf(chkGenCpb.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      .uorstMain_0Fil!genprc = IIf(chkGenPrc.Value = vbChecked, INDPREGEN_ACTx, INDPREGEN_INAx)
      
      End If
    Else
      If proceso = False Then
      'Llaves.
      txtLlave(0).Text = .uorstMain_0!coddro
      txtLlave(1).Text = .uorstMain_0!nroban
      cboTpoBan.ListIndex = .uorstMain_0!tpoban
      
      'Datos.
      dtpFehBan.Value = .uorstMain_0!fehban
      txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain_0!Globan), "", .uorstMain_0!Globan)
      txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain_0!globanx), "", .uorstMain_0!globanx)
      txtDato(2).Text = IIf(IsNull(.uorstMain_0!CodCta), "", .uorstMain_0!CodCta)
      txtDato(3).Text = IIf(IsNull(.uorstMain_0!codcco), "", .uorstMain_0!codcco)
      txtDato(4).Text = IIf(IsNull(.uorstMain_0!codaux), "", .uorstMain_0!codaux)
      txtDato(5).Text = IIf(IsNull(.uorstMain_0!CodFjo), "", .uorstMain_0!CodFjo)
      'cboTpoDoc.ListIndex = .uorstMain_0!tpodoc
      txtDato(10).Text = .uorstMain_0!tpodoc
      txtDato(7).Text = IIf(IsNull(.uorstMain_0!docban), "", .uorstMain_0!docban)
      txtDato(9).Text = IIf(IsNull(.uorstMain_0!codbco), "", .uorstMain_0!codbco)
      txtDato(8).Text = IIf(IsNull(.uorstMain_0!portador), "", .uorstMain_0!portador)
      cboTpoMon.ListIndex = IIf(.uorstMain_0!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      cboTpoTCb.ListIndex = IIf(.uorstMain_0!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
      txtDato(6).Text = Format(IIf(IsNull(.uorstMain_0!ImpTCb), 0, .uorstMain_0!ImpTCb), FORMATO_NUM_2)
      txtDeta(4).Text = Format(.uorstMain_0!ImpMN, FORMATO_NUM_1)
      txtDeta(5).Text = Format(.uorstMain_0!ImpME, FORMATO_NUM_1)
      chkGenCpb.Value = .uorstMain_0!gencpb
      chkGenPrc.Value = .uorstMain_0!genprc
      Else
      'Llaves.
      txtLlave(0).Text = .uorstMain_0Fil!coddro
      txtLlave(1).Text = .uorstMain_0Fil!nroban
      cboTpoBan.ListIndex = .uorstMain_0Fil!tpoban
      
      'Datos.
      dtpFehBan.Value = .uorstMain_0Fil!fehban
      txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain_0Fil!Globan), "", .uorstMain_0Fil!Globan)
      txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain_0Fil!globanx), "", .uorstMain_0Fil!globanx)
      txtDato(2).Text = IIf(IsNull(.uorstMain_0Fil!CodCta), "", .uorstMain_0Fil!CodCta)
      txtDato(3).Text = IIf(IsNull(.uorstMain_0Fil!codcco), "", .uorstMain_0Fil!codcco)
      txtDato(4).Text = IIf(IsNull(.uorstMain_0Fil!codaux), "", .uorstMain_0Fil!codaux)
      txtDato(5).Text = IIf(IsNull(.uorstMain_0Fil!CodFjo), "", .uorstMain_0Fil!CodFjo)
      'cboTpoDoc.ListIndex = .uorstMain_0Fil!tpodoc
      txtDato(10).Text = .uorstMain_0Fil!tpodoc
      txtDato(7).Text = IIf(IsNull(.uorstMain_0Fil!docban), "", .uorstMain_0Fil!docban)
      txtDato(9).Text = IIf(IsNull(.uorstMain_0Fil!codbco), "", .uorstMain_0Fil!codbco)
      txtDato(8).Text = IIf(IsNull(.uorstMain_0Fil!portador), "", .uorstMain_0Fil!portador)
      cboTpoMon.ListIndex = IIf(.uorstMain_0Fil!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      cboTpoTCb.ListIndex = IIf(.uorstMain_0Fil!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
      txtDato(6).Text = Format(IIf(IsNull(.uorstMain_0Fil!ImpTCb), 0, .uorstMain_0Fil!ImpTCb), FORMATO_NUM_2)
      txtDeta(4).Text = Format(.uorstMain_0Fil!ImpMN, FORMATO_NUM_1)
      txtDeta(5).Text = Format(.uorstMain_0Fil!ImpME, FORMATO_NUM_1)
      chkGenCpb.Value = .uorstMain_0Fil!gencpb
      chkGenPrc.Value = .uorstMain_0Fil!genprc
      
      End If
      
      ppAyuDet AYULLA, 0
      ppAyuDet AYUDAT, 2
      ppAyuDet AYUDAT, 3
      ppAyuDet AYUDAT, 4
      ppAyuDet AYUDAT, 5
      ppAyuDet AYUDAT, 9
      ppAyuDet AYUDAT, 10
    
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
  cboTpoBan.ListIndex = TPOBAN_EGR
  txtLlave(0).Text = gsCodDro_Egr
  txtLlave(1).Text = ""
  chkGenCpb.Value = INDPREGEN_ACT
  chkGenPrc.Value = INDPREGEN_INAx
  'Datos.
  dtpFehBan.Value = IIf(Month(Date) = Val(gsMesAct) And Year(Date) = Val(gsAnoAct), Date, gfUltDia("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct))
  With txtDato
    For dnContador = 0 To .Count - 1: .Item(dnContador).Text = "": Next dnContador
  End With
  txtDato(6).Text = Format(0, FORMATO_NUM_2)
  cboTpoTCb.ListIndex = TPOTCB_VTA_IND
  cboTpoMon.ListIndex = TPOMON_NAC_IND

  'Ayudas.
  lblLlaveDeta(0).Caption = ""
  lblDatoDeta(2).Caption = "": lblDatoDeta(3).Caption = ""
  lblDatoDeta(4).Caption = "": lblDatoDeta(5).Caption = ""
  lblDatoDeta(9).Caption = "": lblDatoDeta(10).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Integer
  
  'Datos.
  dtpFehBan.Enabled = pbNuevo
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  cboTpoMon.Enabled = False
  cboTpoTCb.Enabled = pbNuevo
  'cboTpoDoc.Enabled = tbHabilitar
  txtDato(10).Enabled = tbHabilitar
  txtDato(9).Enabled = False
  txtDato(4).Enabled = tbHabilitar
  txtDato(6).Enabled = pbNuevo
  chkGenCpb.Enabled = tbHabilitar
  chkGenPrc.Enabled = tbHabilitar
  'Ayudas.
  cmdLlaveAyud(0).Enabled = pbNuevo: lblLlaveDeta(0).Enabled = tbHabilitar
  cmdDatoAyud(2).Enabled = tbHabilitar: lblDatoDeta(2).Enabled = tbHabilitar
  cmdDatoAyud(3).Enabled = tbHabilitar: lblDatoDeta(3).Enabled = tbHabilitar
  cmdDatoAyud(4).Enabled = tbHabilitar: lblDatoDeta(4).Enabled = tbHabilitar
  cmdDatoAyud(5).Enabled = tbHabilitar: lblDatoDeta(5).Enabled = tbHabilitar
  cmdDatoAyud(10).Enabled = tbHabilitar: lblDatoDeta(10).Enabled = tbHabilitar
  lblDatoDeta(9).Enabled = tbHabilitar
End Sub

'[Código propio del formulario.

Private Sub ppDatosWhere()             'Cambiar.
  If proceso = False Then
    frmTBanGrd_anple.uorstMain_1.Close
    frmTBanGrd_anple.uorstMain_1.Source = frmTBanGrd_anple.usConnStrgSele_1 & frmTBanGrd_anple.usConnStrgWher_1 & frmTBanGrd_anple.usConnStrgOrde_1
    frmTBanGrd_anple.uorstMain_1.Open
    frmTBanGrd_anple.uorstMain_1.Properties("Unique Table").Value = "cobandet"
    cmdCalcular_Click
  Else
    frmTBanGrd_anple.uorstMain_1Fil.Close
    frmTBanGrd_anple.uorstMain_1Fil.Source = frmTBanGrd_anple.usConnStrgSele_1 & frmTBanGrd_anple.usConnStrgWher_1 & frmTBanGrd_anple.usConnStrgOrde_1
    frmTBanGrd_anple.uorstMain_1Fil.Open
    frmTBanGrd_anple.uorstMain_1Fil.Properties("Unique Table").Value = "cobandet"
    cmdCalcular_Click
  
  End If
End Sub

Public Sub upDatosGrid()               'Cambiar Datos Grid.
  Dim dnNum As Integer
  
  With dgrDetalle.Columns
    For dnNum = 0 To .Count - 1
      Select Case dnNum
       Case 0
        .Item(dnNum).Caption = "Item"
        .Item(dnNum).Width = 380
        .Item(dnNum).Alignment = dbgLeft
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
        .Item(dnNum).Width = 750
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "C.Cto.", "Cost Center")
        If proceso = False Then
        .Item(dnNum).Width = 90 * (frmTBanGrd_anple.uorstMain_1.Fields("CodCCo").DefinedSize + 1)
        Else
        .Item(dnNum).Width = 90 * (frmTBanGrd_anple.uorstMain_1Fil.Fields("CodCCo").DefinedSize + 1)
        End If
       Case 3
        .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
        If proceso = False Then
        .Item(dnNum).Width = 90 * (frmTBanGrd_anple.uorstMain_1.Fields("codaux").DefinedSize + 1)
        Else
        .Item(dnNum).Width = 90 * (frmTBanGrd_anple.uorstMain_1Fil.Fields("codaux").DefinedSize + 1)
        End If
       Case 4
        .Item(dnNum).Caption = "Doc"
        If proceso = False Then
        .Item(dnNum).Width = 100 * (frmTBanGrd_anple.uorstMain_1.Fields("AbvTDc").DefinedSize + 1)
        Else
        .Item(dnNum).Width = 100 * (frmTBanGrd_anple.uorstMain_1Fil.Fields("AbvTDc").DefinedSize + 1)
        End If
       Case 5
        .Item(dnNum).Caption = Choose(gsIdioma, "Serie", "Series")
        If proceso = False Then
        .Item(dnNum).Width = 100 * (frmTBanGrd_anple.uorstMain_1.Fields("SerDoc").DefinedSize + 1.5)
        Else
        .Item(dnNum).Width = 100 * (frmTBanGrd_anple.uorstMain_1Fil.Fields("SerDoc").DefinedSize + 1.5)
        End If
       Case 6
        .Item(dnNum).Caption = Choose(gsIdioma, "Numero", "Number")
        If proceso = False Then
        .Item(dnNum).Width = 90 * (frmTBanGrd_anple.uorstMain_1.Fields("NroDoc").DefinedSize + 1)
        Else
        .Item(dnNum).Width = 90 * (frmTBanGrd_anple.uorstMain_1Fil.Fields("NroDoc").DefinedSize + 1)
        End If
       Case 7
        .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
        .Item(dnNum).Width = 1100
       Case 8
        .Item(dnNum).Caption = Choose(gsIdioma, "Debe", "Debit")
        .Item(dnNum).Width = 980 ' * (uorstDetalle.Fields("ImpMN").DefinedSize + 4)
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 9
        .Item(dnNum).Caption = Choose(gsIdioma, "Haber", "Credit")
        .Item(dnNum).Width = 980 ' * (uorstDetalle.Fields("ImpME").DefinedSize + 4)
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 10
        .Item(dnNum).Caption = Choose(gsIdioma, "Debe M.E.", "Debit F.C.")
        .Item(dnNum).Width = 980 ' * (uorstDetalle.Fields("ImpMN").DefinedSize + 4)
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 11
        .Item(dnNum).Caption = Choose(gsIdioma, "Haber M.E.", "Credit F.C.")
        .Item(dnNum).Width = 980 ' * (uorstDetalle.Fields("ImpME").DefinedSize + 4)
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case Else
        .Item(dnNum).Visible = False
      End Select
    Next
  End With
End Sub

Private Sub ppbotones_tpognr0(tbgraba As Boolean)
  '[ No pertenece al Formulario - Agregado por Angel
  cmdCorregir.Enabled = Not pbNuevo
  cmdNuevo.Enabled = True
  cmdEliminar.Enabled = True
  cmdRefrescar.Enabled = True
  cmdCalcular.Enabled = True
  If pbNuevo Then
    cmdNuevo.Enabled = tbgraba
    cmdRevisar.Enabled = tbgraba
    cmdEliminar.Enabled = tbgraba
    cmdRefrescar.Enabled = tbgraba
    cmdImprimir(0).Enabled = tbgraba
    cmdImprimir(1).Enabled = tbgraba
    cmdCalcular.Enabled = tbgraba
    cmdCorregir.Enabled = False
    cmdLlaveAyud(0).Enabled = IIf(pbNuevo, tbgraba, False)
  End If
   ']
End Sub

Private Sub dtpFehBan_LostFocus()
  If Month(dtpFehBan.Value) <> Val(gsMesAct) Or Year(dtpFehBan.Value) <> Val(gsAnoAct) Then
    If Month(dtpFehBan.Value) <> Val(gsMesAct) Then
      If Month(dtpFehBan.Value) = 1 And gsMesAct = "00" Or Month(dtpFehBan.Value) = 12 And gsMesAct = "13" Then
        Exit Sub
      End If
    End If
    MsgBox Choose(gsIdioma, "La fecha debe ser del Mes y Año que se provisiona.", "The date must correspond with Month and Year that provision."), vbExclamation
    dtpFehBan.SetFocus
  End If
  
  With frmTBanGrd_anple.uorstTGTCb
    txtDato(6).Text = Format(0, FORMATO_NUM_2)
    If .RecordCount <> 0 Then
      .MoveFirst
      .Find "FehTCb = '" & frmTBanCab_anple.dtpFehBan & "'"
      If .EOF Then
        MsgBox TEXT_9015, vbExclamation
        txtDato(6).SetFocus
      Else
        txtDato(6).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
      End If
    End If
  End With
End Sub

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
  pbNuevo = tbNuevo
  
  'Orden: Corregir.
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
