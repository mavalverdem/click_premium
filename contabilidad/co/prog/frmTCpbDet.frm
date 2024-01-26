VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTCpbDet 
   Caption         =   "[Entidad]"
   ClientHeight    =   6195
   ClientLeft      =   2025
   ClientTop       =   1500
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6195
   ScaleMode       =   0  'User
   ScaleWidth      =   7980
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
      Height          =   315
      Index           =   6
      Left            =   1080
      TabIndex        =   16
      Top             =   3900
      Width           =   1695
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "&Auxiliar"
      Height          =   375
      Left            =   60
      TabIndex        =   63
      Top             =   4620
      Width           =   1215
   End
   Begin VB.ComboBox CboTpoTCb 
      Height          =   315
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4800
      Width           =   915
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H80000002&
      Height          =   975
      Left            =   3405
      TabIndex        =   52
      Top             =   5160
      Width           =   4545
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
         Left            =   900
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   195
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
         Index           =   1
         Left            =   2700
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   190
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
         Index           =   2
         Left            =   900
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   550
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
         Index           =   3
         Left            =   2700
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   550
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total M.N."
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
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total M.E."
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
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   690
      End
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
      Index           =   8
      Left            =   3180
      TabIndex        =   19
      Top             =   4800
      Width           =   735
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
      Index           =   7
      Left            =   1080
      TabIndex        =   15
      Top             =   3540
      Width           =   6435
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4800
      Width           =   675
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   4500
      TabIndex        =   20
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,###,###.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   6300
      TabIndex        =   22
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,###,###.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   4500
      TabIndex        =   21
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,###,###.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   6300
      TabIndex        =   23
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      ForeColor       =   &H80000002&
      Height          =   1875
      Left            =   60
      TabIndex        =   42
      Top             =   1560
      Width           =   7875
      Begin VB.CommandButton cmdRtcPcp 
         Caption         =   "&Ret./Perc."
         Height          =   375
         Left            =   6600
         TabIndex        =   11
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton optTpoPvs 
         Caption         =   "&Otro"
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
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "P&endientes"
         Height          =   375
         Index           =   4
         Left            =   3360
         Picture         =   "frmTCpbDet.frx":0000
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optTpoPvs 
         Caption         =   "Pro&visión"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   350
         Width           =   975
      End
      Begin VB.OptionButton optTpoPvs 
         Caption         =   "&Cancelación"
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
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   350
         Width           =   1215
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   6240
         Picture         =   "frmTCpbDet.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   715
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
         Left            =   1320
         TabIndex        =   8
         Top             =   700
         Width           =   315
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
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   435
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
         Index           =   5
         Left            =   1740
         TabIndex        =   10
         Top             =   1080
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker dtpFeEDoc 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         OLEDropMode     =   1
         Format          =   62062593
         CurrentDate     =   37959.8076041667
      End
      Begin MSComCtl2.DTPicker dtpFeVDoc 
         Height          =   315
         Left            =   3900
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62062593
         CurrentDate     =   37962.5159027778
      End
      Begin MSComCtl2.DTPicker dtpFeRDoc 
         Height          =   315
         Left            =   6360
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62062593
         CurrentDate     =   37962.5159722222
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Recepción:"
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
         Left            =   5400
         TabIndex        =   62
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento:"
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
         Left            =   2820
         TabIndex        =   61
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión:"
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
         Left            =   120
         TabIndex        =   60
         Top             =   1500
         Width           =   1080
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
         Left            =   1620
         TabIndex        =   45
         Top             =   705
         Width           =   4635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento:"
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
         Left            =   120
         TabIndex        =   44
         Top             =   760
         Width           =   1200
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "NºDocumento:"
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
         Left            =   120
         TabIndex        =   43
         Top             =   1140
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7320
      Picture         =   "frmTCpbDet.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   30
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
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   5160
      Picture         =   "frmTCpbDet.frx":04FE
      Style           =   1  'Graphical
      TabIndex        =   31
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
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   7620
      Picture         =   "frmTCpbDet.frx":06A8
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1215
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
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3480
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
         Left            =   60
         Picture         =   "frmTCpbDet.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   60
         Width           =   360
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
         Left            =   60
         Picture         =   "frmTCpbDet.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   338
         Width           =   360
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
         Picture         =   "frmTCpbDet.frx":0BA6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   60
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
         Picture         =   "frmTCpbDet.frx":0CF0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   60
         Width           =   720
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
         Picture         =   "frmTCpbDet.frx":0DF2
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   60
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
         Left            =   2580
         Picture         =   "frmTCpbDet.frx":0EF4
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComCtl2.DTPicker dtpFehOpe 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62062593
      CurrentDate     =   37924.6695138889
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Referencia:"
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
      Left            =   120
      TabIndex        =   64
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label Label20 
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
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2520
      TabIndex        =   59
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   51
      Top             =   3600
      Width           =   465
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Mon. Func.:"
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
      Left            =   1440
      TabIndex        =   50
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Debe"
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
      Left            =   5100
      TabIndex        =   49
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Haber"
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
      Left            =   6840
      TabIndex        =   48
      Top             =   4200
      Width           =   435
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   " M.N."
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
      Left            =   4080
      TabIndex        =   47
      Top             =   4500
      Width           =   360
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   " M.E."
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
      Left            =   4080
      TabIndex        =   46
      Top             =   4860
      Width           =   345
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   120
      TabIndex        =   41
      Top             =   540
      Width           =   555
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
      Left            =   1800
      TabIndex        =   40
      Top             =   480
      Width           =   5535
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
      Left            =   1440
      TabIndex        =   39
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   120
      TabIndex        =   38
      Top             =   900
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
      Index           =   2
      Left            =   2100
      TabIndex        =   37
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   120
      TabIndex        =   36
      Top             =   1260
      Width           =   585
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Operación:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "frmTCpbDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean
Private pnCta_IndCCo As Integer, _
        pnCta_IndAjD As Integer, _
        pnCta_IndDoc As Integer, _
        pncta_Indanl As Integer
Private pcCodCta_AjD_Deb As String, _
        pcCodCta_AjD_Hab As String
Public pnUltIte, pnTpoMon As Integer
Public pnNroIte As Integer
Public psGlosa As String
Public pbHayRtcPcp As Boolean

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmTCpbGrd                     'Cambiar Formulario de Grid.
    '[Datos.                           'Cambiar.
'      txtDato(0).MaxLength = .uorstMain_1![a.CodArt].DefinedSize
'      txtDato(1).MaxLength = 9
'      txtDato(2).MaxLength = 6
      With cboTpoMon
         .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
         .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
      End With
      With cboTpoTCb
         .AddItem TPOTCB_VTA_TXT, TPOTCB_VTA_IND
         .AddItem TPOTCB_CPR_TXT, TPOTCB_CPR_IND
      End With
      txtDato(0).MaxLength = .uorstMain_1![CodCta].DefinedSize
      txtDato(1).MaxLength = .uorstMain_1![CodCCo].DefinedSize
      txtDato(2).MaxLength = .uorstMain_1![CodAux].DefinedSize
      txtDato(3).MaxLength = .uorstMain_1![CodTDc].DefinedSize
      txtDato(4).MaxLength = .uorstMain_1![SerDoc].DefinedSize
      txtDato(5).MaxLength = .uorstMain_1![NroDoc].DefinedSize
      txtDato(6).MaxLength = .uorstMain_1![RefDoc].DefinedSize
      txtDato(7).MaxLength = .uorstMain_1![GloIte].DefinedSize
      txtDato(8).MaxLength = 7
      txtImporte(0).MaxLength = 14
      txtImporte(1).MaxLength = 14
      txtImporte(2).MaxLength = 14
      txtImporte(3).MaxLength = 14
      txtDeta(0).Text = Format(frmTCpbCab.txtDeta(0).Text, FORMATO_NUM_1)
      txtDeta(1).Text = Format(frmTCpbCab.txtDeta(1).Text, FORMATO_NUM_1)
      txtDeta(2).Text = Format(frmTCpbCab.txtDeta(2).Text, FORMATO_NUM_1)
      txtDeta(3).Text = Format(frmTCpbCab.txtDeta(3).Text, FORMATO_NUM_1)
      psGlosa = frmTCpbCab.txtDato(0).Text
      pnTpoMon = TPOMON_NAC_IND
      
      With dtpFehOpe
         .MinDate = CDate("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct)
         .MaxDate = gfUltDia(.MinDate)
      End With
      dtpFehOpe.Value = frmTCpbCab.dtpFehCpb.Value
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = pbNuevo
   cmdDeshacer.Enabled = False
   upHabilitacion pbNuevo
End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos.           'Cambiar (habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then
      ppAyuDet 0
      pnCta_IndDoc = frmTCpbGrd.uorstCOCta!IndDoc
      pnCta_IndAjD = frmTCpbGrd.uorstCOCta!IndAjD
      pnCta_IndCCo = frmTCpbGrd.uorstCOCta!IndCCo
      pncta_Indanl = frmTCpbGrd.uorstCOCta!TpoAnl
      pcCodCta_AjD_Deb = frmTCpbGrd.uorstCOCta!CodCta_AjD_Deb
      pcCodCta_AjD_Hab = frmTCpbGrd.uorstCOCta!CodCta_AjD_Hab
      If pnCta_IndCCo = INDCCO_INA Then
         txtDato(1).Enabled = False
         cmdDatoAyud(1).Enabled = False
      Else
         txtDato(1).Enabled = True
         cmdDatoAyud(1).Enabled = True
      End If
   End If
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
   If txtDato(3).Text <> "" Then ppAyuDet 3
 ']

   If Not pbNuevo Then
      If frmTCpbGrd.uorstMain_1.RecordCount > 0 And frmTCpbGrd.uorstMain_1!TpoGnr <> TPOGNR_DRO Then cmdCorregir.Enabled = False
   End If
   
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   frmTFacGrd.uorstMain_1.CancelUpdate 'Cambiar Formulario de Grid.
   If frmTCpbGrd.uorstMain_1.RecordCount <> 0 Then
      frmTCpbGrd.uorstMain_1.CancelUpdate
   End If
End Sub

Private Sub cmdRetroceder_Click()
'   gpTUe_Retroceder frmTFacGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
'''   frmTCpbCab.cmdCalcular_Click
'''   frmTCpbGrd.uorstMain_1.Find "cLlave='" & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1) & Trim(Str(pnUltIte)) & "'"
   gpTUe_Retroceder frmTCpbGrd.uorstMain_1, Me
End Sub

Private Sub cmdAvanzar_Click()
'   gpTUe_Avanzar frmTFacGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
'''   frmTCpbCab.cmdCalcular_Click
'''   frmTCpbGrd.uorstMain_1.Find "cLlave='" & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1) & Trim(Str(pnUltIte)) & "'"
   gpTUe_Avanzar frmTCpbGrd.uorstMain_1, Me
End Sub

Public Sub cmdCorregir_Click()
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   ppHabilitaDatosDocumento
   If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
      cmdRtcPcp.Enabled = True
   Else
      cmdRtcPcp.Enabled = False
   End If
 '[Dato con el foco al corregir.       'Cambiar.
   dtpFehOpe.SetFocus
   cmdDatoAyud(4).Enabled = (optTpoPvs(1).Value)    ' Efectos de cancelacion de documentos
   dtpFeEDoc.Enabled = (optTpoPvs(0).Value)
   dtpFeVDoc.Enabled = (optTpoPvs(0).Value)
   dtpFeRDoc.Enabled = (optTpoPvs(0).Value)
'   txtDato(0).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err

 '[No pertenece al Formulario - Agregado por Angel
   Dim dnNroIte As Integer
   Dim dnImpMN, dnImpME As Double
   Dim dcTpoMon, dcTpoCtb As String
   Dim dvRegistro As Variant
   Dim dbSinImportes As Boolean

 '[Validacion de Datos segun Indicadores de Cuenta.
   If Len(Trim(txtDato(0).Text)) = 0 Then
      MsgBox TEXT_6002, vbExclamation
      txtDato(0).SetFocus
      Exit Sub
   End If
   If pnCta_IndCCo = INDCCO_ACT And Len(Trim(txtDato(1).Text)) = 0 Then
'      MsgBox "Debe asignar el Centro de Costo.", vbExclamation
      MsgBox TEXT_6002, vbExclamation
      txtDato(1).SetFocus
      Exit Sub
   End If
   If pnCta_IndDoc = INDDOC_ACT And (Len(Trim(txtDato(3).Text)) = 0 Or Len(Trim(txtDato(4).Text)) = 0 Or Len(Trim(txtDato(5).Text)) = 0) Then
'      MsgBox "Debe registrar datos completos del documento.", vbExclamation
      MsgBox TEXT_6002, vbExclamation
      txtDato(3).SetFocus
      Exit Sub
   End If
   'valida cta+auxiliar
   If pnCta_IndAjD = INDAJD_ACT And pncta_Indanl = TPOANL_AUX And pnCta_IndDoc = INDAUX_ACT And Len(Trim(txtDato(2).Text)) = 0 Then
'     MsgBox "Debe registrar datos completos del documento.", vbExclamation
      MsgBox TEXT_6002, vbExclamation
      txtDato(2).SetFocus
      Exit Sub
   End If
   
   If Len(Trim(txtDato(3).Text)) <> 0 And Len(Trim(txtDato(2).Text)) = 0 Then
'      MsgBox "Debe asignar el auxiliar para el documento registrado.", vbExclamation
      MsgBox TEXT_6002, vbExclamation
      txtDato(2).SetFocus
      Exit Sub
   End If
   If CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0 And CDec(txtImporte(2).Text) = 0 And CDec(txtImporte(3).Text) = 0 Then
      If MsgBox("¿Grabará el detalle sin asignar importes?", vbYesNo) = vbNo Then
          txtImporte(0).SetFocus
          Exit Sub
      Else
         dbSinImportes = True
      End If
   Else
      If CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0 Then
         If MsgBox("¿Grabará el detalle sin asignar importes en " & TPOMON_NAC_TXT_2 & "?", vbYesNo) = 7 Then
             txtImporte(0).SetFocus
             Exit Sub
         End If
      End If
      If CDec(txtImporte(2).Text) = 0 And CDec(txtImporte(3).Text) = 0 Then
         If MsgBox("¿Grabará el detalle sin asignar importes en " & TPOMON_EXT_TXT_2 & "?", vbYesNo) = 7 Then
             txtImporte(2).SetFocus
             Exit Sub
         End If
      End If
   End If
   If Not dbSinImportes Then
      If cboTpoMon.ListIndex = TPOMON_NAC_IND And (CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0) Then
         MsgBox "Debe ingresar el importe en Moneda Nacional.", vbInformation
         txtImporte(0).SetFocus
         Exit Sub
      ElseIf cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtImporte(2).Text) = 0 And CDec(txtImporte(3).Text) = 0) Then
         MsgBox "Debe ingresar el importe en Moneda Extranjera.", vbInformation
         txtImporte(2).SetFocus
         Exit Sub
      End If
   End If
   
   If Len(Trim(txtDato(2).Text)) <> 0 And Len(Trim(txtDato(3).Text)) <> 0 And Len(Trim(txtDato(4).Text)) <> 0 And Len(Trim(txtDato(5).Text)) <> 0 Then
      With frmTCpbGrd.uorstCOCpbDet
         .Source = "SELECT CodAux, CodTDc, SerDoc, NroDoc, ImpMN, ImpME, ImpTCb, TpoTCb, TpoMon, TpoPvs, TpoCtb, CodCta, CodDro, NroCpb " _
                 & "FROM CoCpbDet " _
                 & "WHERE CodAux ='" & txtDato(2).Text & "'" _
                 & "  AND CodCta ='" & txtDato(0).Text & "'" _
                 & "  AND CodTDc='" & txtDato(3).Text & "'" _
                 & "  AND SerDoc='" & txtDato(4).Text & "'" _
                 & "  AND NroDoc='" & txtDato(5).Text & "'" _
                 & "  AND TpoPvs<>'" & TPOPVS_OTR & "'"
         .Open
         dnImpMN = 0
         dnImpME = 0
         If optTpoPvs(0).Value Then
            frmTCpbGrd.uorstCOCpbDet.Find "TpoPvs='" & TPOPVS_PVS & "'"
            If Not frmTCpbGrd.uorstCOCpbDet.EOF Then
               If frmTCpbGrd.uorstCOCpbDet!CodDro <> frmTCpbCab.txtLlave(0).Text Or frmTCpbGrd.uorstCOCpbDet!NroCpb <> frmTCpbCab.txtLlave(1).Text Then
                  MsgBox "Ya está registrada la provision del documento.", vbExclamation
                  frmTCpbGrd.uorstCOCpbDet.Close
                  optTpoPvs(0).SetFocus
                  Exit Sub
               End If
            End If
         Else
            If pbNuevo And optTpoPvs(1).Value Then
               frmTCpbGrd.uorstCOCpbDet.Find "TpoPvs='" & TPOPVS_PVS & "'"
               If .EOF Then
                  MsgBox "No está registrada la provisión del documento.", vbExclamation
                  frmTCpbGrd.uorstCOCpbDet.Close
                  optTpoPvs(1).SetFocus
                  Exit Sub
               Else
                  If frmTCpbGrd.uorstCOCpbDet!CodCta <> txtDato(0).Text Then
                     MsgBox "La cuenta de la cancelación no es igual a la de la provisión.", vbExclamation
                     frmTCpbGrd.uorstCOCpbDet.Close
                     txtDato(0).SetFocus
                     Exit Sub
                  End If
                  If frmTCpbGrd.uorstCOCpbDet!TpoCtb = TPOCTB_DEB And (CDec(txtImporte(0).Text) > 0 Or CDec(txtImporte(2).Text) > 0) Then
                     MsgBox "Revise la información. La provisión está registrada en el DEBE."
                     frmTCpbGrd.uorstCOCpbDet.Close
                     txtImporte(1).SetFocus
                     Exit Sub
                  End If
                  If frmTCpbGrd.uorstCOCpbDet!TpoCtb = TPOCTB_HAB And (CDec(txtImporte(1).Text) > 0 Or CDec(txtImporte(3).Text) > 0) Then
                     MsgBox "Revise la información. La provisión está registrada en el HABER."
                     frmTCpbGrd.uorstCOCpbDet.Close
                     txtImporte(0).SetFocus
                     Exit Sub
                  End If
               End If
            End If
            If (Not .EOF) And optTpoPvs(1).Value Then
               dcTpoMon = frmTCpbGrd.uorstCOCpbDet!TpoMon
               frmTCpbGrd.uorstCOCpbDet.MoveFirst
               Do
'                  If frmTCpbGrd.uorstCOCpbDet!CodDro <> frmTCpbCab.txtLlave(0).Text And frmTCpbGrd.uorstCOCpbDet!NroCpb <> frmTCpbCab.txtLlave(1).Text Then
                  If frmTCpbGrd.uorstCOCpbDet!CodDro & frmTCpbGrd.uorstCOCpbDet!NroCpb <> frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1).Text Then
                     dnImpMN = dnImpMN + IIf(frmTCpbGrd.uorstCOCpbDet!TpoPvs = TPOPVS_PVS, frmTCpbGrd.uorstCOCpbDet!ImpMN, frmTCpbGrd.uorstCOCpbDet!ImpMN * (-1))
                     dnImpME = dnImpME + IIf(frmTCpbGrd.uorstCOCpbDet!TpoPvs = TPOPVS_PVS, frmTCpbGrd.uorstCOCpbDet!ImpME, frmTCpbGrd.uorstCOCpbDet!ImpME * (-1))
                  End If
                  frmTCpbGrd.uorstCOCpbDet.MoveNext
               Loop Until .EOF
               If dcTpoMon = TPOMON_NAC Then
                  If CDec(txtImporte(0).Text) > 0 Then
                     If dnImpMN < CDec(txtImporte(0).Text) Then
                        MsgBox "El monto de la cancelación es mayor al de la provisión.", vbExclamation
                        frmTCpbGrd.uorstCOCpbDet.Close
                        txtImporte(0).SetFocus
                        Exit Sub
                     End If
                  End If
                  If CDec(txtImporte(1).Text) > 0 Then
                     If dnImpMN < CDec(txtImporte(1).Text) Then
                        MsgBox "El monto de la cancelación es mayor al de la provisión.", vbExclamation
                        frmTCpbGrd.uorstCOCpbDet.Close
                        txtImporte(1).SetFocus
                        Exit Sub
                     End If
                  End If
               Else
                  If CDec(txtImporte(2).Text) > 0 Then
                     If dnImpME < CDec(txtImporte(2).Text) Then
                        MsgBox "El monto de la cancelación es mayor al de la provisión.", vbExclamation
                        frmTCpbGrd.uorstCOCpbDet.Close
                        txtImporte(2).SetFocus
                        Exit Sub
                     End If
                  End If
                  If CDec(txtImporte(3).Text) > 0 Then
                     If dnImpME < CDec(txtImporte(3).Text) Then
                        MsgBox "El monto de la cancelación es mayor al de la provisión.", vbExclamation
                        frmTCpbGrd.uorstCOCpbDet.Close
                        txtImporte(3).SetFocus
                        Exit Sub
                     End If
                  End If
               End If
            End If
         End If
         frmTCpbGrd.uorstCOCpbDet.Close
      End With
   End If
      
   With frmTCpbGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain_1.AddNew
      Else
         .uorstMain_1.Find "cLlave='" & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1) & Trim(Str(pnUltIte)) & "'"
      End If
      upDatosDesconectados 0
      With .uorstMain_1
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
            !FyHMdf = Now
         End If
         .Update
      End With
    '[Propio del formulario.
     'Generación por Percepciones/Retenciones.
      If pbHayRtcPcp Then
         With .uorstCOCpbDetRP
            If pbNuevo Then
               .AddNew
            End If
            !MesPvs = gsMesAct
            !CodDro = frmTCpbCab.txtLlave(0).Text
            !NroCpb = frmTCpbCab.txtLlave(1).Text
            !NroIte = pnNroIte
            !CodAux = frmTCpbDet.txtDato(2).Text
            !CodCta = frmTCpbDet.txtDato(0).Text
            !CodTDc = frmTCpbDet.txtDato(3).Text
            !SerDoc = frmTCpbDet.txtDato(4).Text
            !NroDoc = frmTCpbDet.txtDato(5).Text
            !CodTDc_RtcPcp = "90"
            !SerDoc_RtcPcp = frmTCpbDetRet.txtDato(0).Text
            !NroDoc_RtcPcp = frmTCpbDetRet.txtDato(1).Text
            If pbNuevo Then
               !UsrCre = gsAbvUsr
               !FyHCre = Now
            Else
               !UsrMdf = gsAbvUsr
               !FyHMdf = Now
            End If
            .Update
         
            ppRtcPcp pbNuevo
         End With
      End If
     
     'Ajustes por Diferencia de Cambio.
      If pnCta_IndAjD = INDAJD_ACT And optTpoPvs(1).Value Then
        'Eliminación de Ajuste.
         If Not pbNuevo Then
            With .uorstMain_1
               dvRegistro = .Bookmark
               If Not .RecordCount = 0 Then
                  .MoveFirst
                  Do
                     If !CodDro = frmTCpbCab.txtLlave(0).Text And !NroCpb = frmTCpbCab.txtLlave(1).Text And !BlqIte = pnNroIte And !TpoGnr = TPOGNR_DCA Then
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
               .Update
            End With
         End If
        'Generación de Ajuste.
         ppGenera_AjD
      End If
    ']
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      
      If pbNuevo Then
         dnNroIte = .uorstMain_1!NroIte
         .uorstMain_1.Requery
         frmTCpbCab.upDatosGrid
         
         If .uorstMain_1.RecordCount <> 0 Then
          '[Búsqueda de llave actual.     'Cambiar.
            .uorstMain_1.MoveFirst
            .uorstMain_1.Find "cLlave='" & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
          ']
         End If
   
         upDatosPredeterminados
       '[Dato con el foco al añadir.   'Cambiar.
         dtpFehOpe.SetFocus
       ']
      Else
         If .uorstMain_1.RecordCount <> 0 Then
            dnNroIte = .uorstMain_1!NroIte
          '[Búsqueda de llave actual.     'Cambiar.
            .uorstMain_1.MoveFirst
            .uorstMain_1.Find "cLlave='" & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
            If .uorstMain_1.EOF Then .uorstMain_1.MoveFirst
          ']
         End If
         cmdRetroceder.Enabled = True
         cmdAvanzar.Enabled = True
         cmdCorregir.Enabled = True
         cmdGrabar.Enabled = False
         cmdDeshacer.Enabled = False
         upHabilitacion False
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
  
   frmTCpbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
    '[Propio del formulario.
   With frmTCpbCab
      .txtDeta(0).Text = Format(.txtDeta(0).Tag, FORMATO_NUM_1)
      .txtDeta(1).Text = Format(.txtDeta(1).Tag, FORMATO_NUM_1)
      .txtDeta(2).Text = Format(.txtDeta(2).Tag, FORMATO_NUM_1)
      .txtDeta(3).Text = Format(.txtDeta(3).Tag, FORMATO_NUM_1)
   End With
    ']
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2, 3
      If (pnCta_IndCCo = INDCCO_ACT And Index = 1) Or Index <> 1 Then
         txtDato(Index).SetFocus
      End If
'   Case 2, 3
'      mskDato(Index).SetFocus
    Case 4  ' Inserto los documentos agrupados a la tabla tempolral
        txtDato(Index).SetFocus
        cmdDatoAyud(4).Tag = "INSERT INTO codoctmp1"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " SELECT CodCta, CodAux, CodTDc, SerDoc, NroDoc,"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " SUM(IF(TpoCtb = 'D', ImpMN, 0)) as DebeMN, SUM(IF(TpoCtb = 'H', ImpMN, 0)) as HaberMN,"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " SUM(IF(TpoCtb = 'D', ImpME, 0)) as DebeME, SUM(IF(TpoCtb = 'H', ImpME, 0)) as HaberME,"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " SUM((IF(TpoCtb = 'D', ImpMN, 0) - IF(TpoCtb = 'H', ImpMN, 0))) as SaldoMN,"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " SUM((IF(TpoCtb = 'D', ImpME, 0) - IF(TpoCtb = 'H', ImpME, 0))) as SaldoME, '" & gsAbvUsr & "'"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " FROM cocpbdet"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " WHERE CodCta = '" & txtDato(0).Text & "'"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " AND CodAux = '" & txtDato(2).Text & "'"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " AND MesPvs <= '" & gsMesAct & "'"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " AND (FeEDoc) <= '" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "'"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " GROUP BY CodTDc, SerDoc, NroDoc"
        cmdDatoAyud(4).Tag = cmdDatoAyud(4).Tag & " ORDER BY CodTDc, SerDoc, NroDoc"
        frmTCpbGrd.uocnnMain.Execute cmdDatoAyud(4).Tag
   End Select
   If (pnCta_IndCCo = INDCCO_ACT And Index = 1) Or Index <> 1 Then
      ppAyuBus Index
   End If
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
   If Index = 8 Then
      txtDato(Index).SelStart = 0
      txtDato(Index).SelLength = txtDato(Index).MaxLength + 1
   ''txtDato(8).Text = ""
   End If
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
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_LostFocus(Index As Integer)
   If Index = 0 Then
      ppHabilitaDatosDocumento
   End If
   If Index = 8 Then
      If Val(txtDato(Index).Text) > 0 Then
         txtDato(Index).Text = Format(Val(txtDato(Index).Text), FORMATO_NUM_2)
      End If
   End If
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
  
  'Completa con ceros a la izquierda.
   Select Case Index
   Case 3, 4, 5                              'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength And IsNumeric(txtDato(Index).Text) Then
         ' cuando sea numericos
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

  'Asigna 0 a campos numéricos si están vacíos.
'   Select Case Index
'   Case 1, 2                           'Cambiar (añadir índices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      If lblDatoDeta(Index).Caption <> "" Then
        pnCta_IndDoc = frmTCpbGrd.uorstCOCta!IndDoc
        pnCta_IndAjD = frmTCpbGrd.uorstCOCta!IndAjD
        pnCta_IndCCo = frmTCpbGrd.uorstCOCta!IndCCo
        pncta_Indanl = frmTCpbGrd.uorstCOCta!TpoAnl
        pcCodCta_AjD_Deb = frmTCpbGrd.uorstCOCta!CodCta_AjD_Deb
        pcCodCta_AjD_Hab = frmTCpbGrd.uorstCOCta!CodCta_AjD_Hab
        If pnCta_IndCCo = INDCCO_INA Then
           txtDato(1).Enabled = False
           cmdDatoAyud(1).Enabled = False
        Else
           txtDato(1).Enabled = True
           cmdDatoAyud(1).Enabled = True
        End If
      End If
   Case 1, 2, 3
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   Case 8
      txtDato(Index).Text = Format(CDec(IIf(txtDato(Index).Text = "", 0, txtDato(Index).Text)), FORMATO_NUM_2)
   End Select
 
 '[Propio del formulario. - Agregado por Angel, jalado de formulario anterior
   If Index = 0 Or Index = 8 Then
'      If txtDato(8).Tag <> frmTCpbGrd.uorstCOCta!TpoTCb Or Val(txtDato(8)) = 0 Then
'         txtDato(8).Tag = frmTCpbGrd.uorstCOCta!TpoTCb
''       If txtDato(8).Tag <> TPOTCB_VTA Or Val(txtDato(8)) = 0 Then
       If Val(txtDato(8)) = 0 Then
         txtDato(8).Tag = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
         With frmTCpbGrd.uorstTGTCb
            If .RecordCount <> 0 Then
               .MoveFirst
               .Find "FehTCb = '" & IIf(pnCta_IndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmTCpbDet.dtpFehOpe, frmTCpbDet.dtpFeEDoc), frmTCpbDet.dtpFehOpe) & "'"
         ' [Adicional Agregado por Angel
               If .EOF Then
                  MsgBox TEXT_9015, vbExclamation
                  txtDato(8).Text = Format(0, FORMATO_NUM_2)
                  Index = Index - 1
                  txtDato(0).SetFocus
               Else
'            frmTCpbGrd.uorstMain_1!ImpTCb = IIf(frmTCpbGrd.uorstCOCta!TpoTCb = TPOTCB_COM, !ImpTCb_Cpr, !ImpTCb_Vta)
'                  txtDato(8).Text = Format(IIf(frmTCpbGrd.uorstCOCta!TpoTCb = TPOTCB_COM, !ImpTCb_Cpr, !ImpTCb_Vta), FORMATO_NUM_2)
                  txtDato(8).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
               End If
         ']
            Else
               txtDato(8).Text = Format(0, FORMATO_NUM_2)
            End If
         End With
      End If
   End If
   If Index = 3 Or Index = 4 Or Index = 5 Then
      If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
         cmdRtcPcp.Enabled = True
      Else
         cmdRtcPcp.Enabled = False
      End If
   End If
 ']
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1                              'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "Length(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2                              'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 3                              'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 4                              'Cambiar (añadir índices).
      cmdDatoAyud(tnIndex).Tag = "b.TpoPvs = '" & TPOPVS_PVS & "' AND a.UsrCre = '" & gsAbvUsr & "'"
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & " AND IF(b.TpoMon = '" & TPOMON_NAC & "', a.ImpSMN, a.ImpSME) <> 0"
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & " AND b.MesPvs <= '" & gsMesAct & "'"
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & " AND b.FeEDoc <= ('" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "')"
      modAyuBus.Sal_Doc cmdDatoAyud(tnIndex).Tag, txtDato(3).Text & txtDato(4).Text & txtDato(5).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      ' Elimino los datos de la tabla temporal
      frmTCpbGrd.uocnnMain.Execute "DELETE FROM CODocTmp1 WHERE UsrCre = '" & gsAbvUsr & "'"
      txtDato(3).Text = Left(frmOAyuBus.uvDato1, 2)
      txtDato(4).Text = Mid(frmOAyuBus.uvDato1, 3, 3)
      txtDato(5).Text = Mid(frmOAyuBus.uvDato1, 6)
      ' Obtengo los datos por default del documento
      With frmTCpbGrd.uorstCOTCbMes
         .Source = "SELECT FeEDoc, FeVDoc, FeRDoc, TpoMon FROM COCpbDet " _
                 & "WHERE CodCta='" & txtDato(0).Text & "'" _
                 & "  AND CodAux='" & txtDato(2).Text & "'" _
                 & "  AND CodTDc='" & txtDato(3).Text & "'" _
                 & "  AND SerDoc='" & txtDato(4).Text & "'" _
                 & "  AND NroDoc='" & txtDato(5).Text & "'" _
                 & "  AND TpoPvs='" & TPOPVS_PVS & "'"
         .Open
         If .RecordCount <> 0 Then
            dtpFeEDoc = Format(!FeEDoc, "dd/mm/yyyy")
            dtpFeVDoc = Format(!FeVDoc, "dd/mm/yyyy")
            dtpFeRDoc = Format(!FeRDoc, "dd/mm/yyyy")
            cboTpoMon.ListIndex = IIf(!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         End If
         .Close
      End With
      If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
         txtImporte(IIf(frmOAyuBus.uvDato2 < 0, 0, 1)).Text = Abs(Val(frmOAyuBus.uvDato2))
      Else
         txtImporte(IIf(Val(frmOAyuBus.uvDato2) < 0, 2, 3)).Text = Abs(Val(frmOAyuBus.uvDato2))
      End If
      If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
         cmdRtcPcp.Enabled = True
      Else
         cmdRtcPcp.Enabled = False
      End If
   End Select
End Sub
Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTCpbGrd.uorstCOCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTCpbGrd.uorstCOCta!DetCta
         End If
      End With
   Case 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTCpbGrd.uorstCOCCo
         .MoveFirst
         .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTCpbGrd.uorstCOCCo!DetCCo
         End If
      End With
   Case 2
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTCpbGrd.uorstTGAux
         .MoveFirst
         .Find "CodAux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTCpbGrd.uorstTGAux!RazAux
         End If
      End With
   Case 3
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTCpbGrd.uorstTGTDc
         .MoveFirst
         .Find "CodTDc='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTCpbGrd.uorstTGTDc!DetTDc
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   Dim dnContador As Integer
   On Error GoTo Err

   With frmTCpbGrd                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain_1!CodDro = frmTCpbCab.txtLlave(0).Text
            .uorstMain_1!NroCpb = frmTCpbCab.txtLlave(1).Text
            .uorstMain_1!MesPvs = gsMesAct
            With .uorstUltiItem
               .Source = "SELECT IFNULL(MAX(NroIte), 0) AS cUltIte " _
                       & "FROM COCpbDet " _
                       & "WHERE CodDro='" & frmTCpbCab.txtLlave(0).Text & "' And NroCpb='" & frmTCpbCab.txtLlave(1).Text & "' And MesPvs='" & gsMesAct & "'"
               .Open
               pnNroIte = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
               frmTCpbGrd.uorstMain_1!NroIte = pnNroIte
               .Close
            End With
            .uorstMain_1!BlqIte = pnNroIte
         End If
        
        'Datos.
'         uorstMain_1!EstCCo = IIf(chkEstado.Value = vbChecked, ESTCCO_ACT, ESTCCO_INA)
'         uorstMain_1!CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         uorstMain_1!FehOpe = dtpFecha.Value
'         uorstMain_1!CodMon = optMoneda(1).Value
         .uorstMain_1!FehOpe = dtpFehOpe.Value
         .uorstMain_1!CodCta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
         .uorstMain_1!CodCCo = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
         .uorstMain_1!CodAux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
         .uorstMain_1!CodTDc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
         .uorstMain_1!SerDoc = txtDato(4).Text
         .uorstMain_1!NroDoc = txtDato(5).Text
         .uorstMain_1!RefDoc = txtDato(6).Text
         .uorstMain_1!FeEDoc = dtpFeEDoc.Value
         .uorstMain_1!FeVDoc = dtpFeVDoc.Value
         .uorstMain_1!FeRDoc = dtpFeRDoc.Value
         .uorstMain_1!GloIte = txtDato(7).Text
         psGlosa = txtDato(7).Text
         .uorstMain_1!TpoPvs = IIf(optTpoPvs(0).Value, TPOPVS_PVS, IIf(optTpoPvs(1).Value, TPOPVS_CAN, TPOPVS_OTR))
         .uorstMain_1!TpoMon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
         pnTpoMon = cboTpoMon.ListIndex
         .uorstMain_1!ImpTCb = txtDato(8).Text
         .uorstMain_1!ImpMN = CDec(IIf(txtImporte(0).Text <> 0, txtImporte(0).Text, txtImporte(1).Text))
         .uorstMain_1!ImpME = CDec(IIf(txtImporte(2).Text <> 0, txtImporte(2).Text, txtImporte(3).Text))
         'cambio  .uorstMain_1!TpoCtb = IIf(txtImporte(0).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
         .uorstMain_1!TpoCtb = IIf(txtImporte(0).Text = 0 And txtImporte(2).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
         .uorstMain_1!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
         .uorstMain_1!TpoGnr = TPOGNR_DRO
      Else
        'Llaves.
'         txtLlave(0).Text = .uorstMain_1!CodSvc
        
        'Datos.
'         chkEstado.Value = IIf(uorstMain_1!EstCCo = ESTCCO_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(uorstMain_1!CodSoc), "", uorstMain_1!CodSoc)
'         dtpFecha.Value = uorstMain_1!FehOpe
'         optMoneda(1).Value = uorstMain_1!CodMon
         If .uorstMain_1.EOF Then .uorstMain_1.MoveLast
         pnUltIte = .uorstMain_1!NroIte
         pnNroIte = .uorstMain_1!BlqIte
         dtpFehOpe.Value = .uorstMain_1!FehOpe
         txtDato(0).Text = IIf(IsNull(.uorstMain_1!CodCta), "", .uorstMain_1!CodCta)
         txtDato(1).Text = IIf(IsNull(.uorstMain_1!CodCCo), "", .uorstMain_1!CodCCo)
         txtDato(2).Text = IIf(IsNull(.uorstMain_1!CodAux), "", .uorstMain_1!CodAux)
         txtDato(3).Text = IIf(IsNull(.uorstMain_1!CodTDc), "", .uorstMain_1!CodTDc)
         txtDato(4).Text = IIf(IsNull(.uorstMain_1!SerDoc), "", .uorstMain_1!SerDoc)
         txtDato(5).Text = IIf(IsNull(.uorstMain_1!NroDoc), "", .uorstMain_1!NroDoc)
         txtDato(6).Text = IIf(IsNull(.uorstMain_1!RefDoc), "", .uorstMain_1!RefDoc)
         dtpFeEDoc.Value = IIf(IsNull(.uorstMain_1!FeEDoc), .uorstMain_1!FehOpe, .uorstMain_1!FeEDoc)
         dtpFeVDoc.Value = IIf(IsNull(.uorstMain_1!FeVDoc), .uorstMain_1!FehOpe, .uorstMain_1!FeVDoc)
         dtpFeRDoc.Value = IIf(IsNull(.uorstMain_1!FeRDoc), .uorstMain_1!FehOpe, .uorstMain_1!FeRDoc)
         txtDato(7).Text = IIf(IsNull(.uorstMain_1!GloIte), "", .uorstMain_1!GloIte)
         optTpoPvs(0).Value = IIf(.uorstMain_1!TpoPvs = TPOPVS_PVS, TPOPVS_PVS_VER, TPOPVS_PVS_FAL)
         optTpoPvs(1).Value = IIf(.uorstMain_1!TpoPvs = TPOPVS_CAN, TPOPVS_CAN_VER, TPOPVS_CAN_FAL)
         optTpoPvs(2).Value = IIf(.uorstMain_1!TpoPvs = TPOPVS_OTR, TPOPVS_OTR_VER, TPOPVS_OTR_FAL)
         cboTpoMon.ListIndex = IIf(.uorstMain_1!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         cboTpoTCb.ListIndex = IIf(.uorstMain_1!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
'         txtDato(8).Text = IIf(IsNull(.uorstMain_1!ImpTCb), 0, .uorstMain_1!ImpTCb)
         txtDato(8).Text = Format(IIf(IsNull(.uorstMain_1!ImpTCb), 0, .uorstMain_1!ImpTCb), FORMATO_NUM_2)
         If .uorstMain_1!TpoCtb = TPOCTB_DEB Then
            txtImporte(0).Text = Format(IIf(IsNull(.uorstMain_1!ImpMN), 0, .uorstMain_1!ImpMN), FORMATO_NUM_1)
            txtImporte(2).Text = Format(IIf(IsNull(.uorstMain_1!ImpME), 0, .uorstMain_1!ImpME), FORMATO_NUM_1)
            txtImporte(1).Text = Format(0, FORMATO_NUM_1)
            txtImporte(3).Text = Format(0, FORMATO_NUM_1)
         Else
            txtImporte(0).Text = Format(0, FORMATO_NUM_1)
            txtImporte(2).Text = Format(0, FORMATO_NUM_1)
            txtImporte(1).Text = Format(IIf(IsNull(.uorstMain_1!ImpMN), 0, .uorstMain_1!ImpMN), FORMATO_NUM_1)
            txtImporte(3).Text = Format(IIf(IsNull(.uorstMain_1!ImpME), 0, .uorstMain_1!ImpME), FORMATO_NUM_1)
         End If
         txtImporte(0).Tag = Format(txtImporte(0).Text, FORMATO_NUM_1)
         txtImporte(1).Tag = Format(txtImporte(1).Text, FORMATO_NUM_1)
         txtImporte(2).Tag = Format(txtImporte(2).Text, FORMATO_NUM_1)
         txtImporte(3).Tag = Format(txtImporte(3).Text, FORMATO_NUM_1)
         '[ Para mostrar los totales
         With txtDeta
            For dnContador = 0 To .Count - 1
               .Item(dnContador).Text = Format(frmTCpbCab.txtDeta.Item(dnContador).Text, FORMATO_NUM_1)
            Next
         End With
         cmdRtcPcp.Enabled = False
         ppAyuDet (0)
         ppAyuDet (1)
         ppAyuDet (2)
         ppAyuDet (3)
         ']
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
'   txtLlave(0).Text = ""

  'Datos.
'   chkEstado.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
         If dnContador = 7 Then .Item(dnContador).Text = psGlosa
         If dnContador = 8 Then .Item(dnContador).Tag = ""
      Next
   End With
   txtDato(8).Text = Format(0, FORMATO_NUM_2)
   cboTpoMon.ListIndex = pnTpoMon
   cboTpoTCb.ListIndex = TPOTCB_VTA_IND
   optTpoPvs.Item(0) = TPOPVS_PVS_FAL
   optTpoPvs.Item(1) = TPOPVS_CAN_FAL
   optTpoPvs.Item(2) = TPOPVS_OTR_VER
   dtpFeEDoc.Value = frmTCpbCab.dtpFehCpb.Value
   dtpFeVDoc.Value = frmTCpbCab.dtpFehCpb.Value
   dtpFeRDoc.Value = frmTCpbCab.dtpFehCpb.Value
   With txtImporte
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = Format(0, FORMATO_NUM_1)
         .Item(dnContador).Tag = Format(0, FORMATO_NUM_1)
      Next
   End With
  
  'Ayudas.
   With lblDatoDeta
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Caption = ""
      Next
   End With
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   With txtImporte
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   cboTpoMon.Enabled = tbHabilitar
   cboTpoTCb.Enabled = tbHabilitar
   dtpFehOpe.Enabled = tbHabilitar
   dtpFeEDoc.Enabled = tbHabilitar
   dtpFeVDoc.Enabled = tbHabilitar
   dtpFeRDoc.Enabled = tbHabilitar
   With optTpoPvs
      For dnContador = 0 To .Count - 1
        .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
  'Ayudas.
  With cmdDatoAyud
     For dnContador = 0 To .Count - 1
        .Item(dnContador).Enabled = tbHabilitar
     Next
  End With
  With lblDatoDeta
     For dnContador = 0 To .Count - 1
        .Item(dnContador).Enabled = tbHabilitar
     Next
  End With
  cmdRtcPcp.Enabled = False
'   cmdDatoAyud(0).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
End Sub

'[Código propio del formulario.

Private Sub cmdAuxiliar_Click()
   frmMAuxGrd.Show vbModal
   frmTCpbGrd.uorstTGAux.Requery
End Sub

Private Sub cmdRtcPcp_Click()
   frmTCpbDetRet.zbNuevo = pbNuevo
   frmTCpbDetRet.Show vbModal
End Sub

Private Sub dtpFeEDoc_LostFocus()
   Dim dcValIndDoc As Integer
   
   If IsNull(frmTCpbDet.dtpFeEDoc) Then
      MsgBox "Verifique activación y registro de la Fecha de Emisión del documento.", vbExclamation
      dtpFeEDoc.Enabled = Not dtpFeEDoc.Enabled
      dtpFeEDoc.SetFocus
      Exit Sub
   End If
   dcValIndDoc = frmTCpbGrd.uorstCOCta!IndDoc
   With frmTCpbGrd.uorstTGTCb
      If .RecordCount <> 0 Then
         .MoveFirst
         .Find "FehTCb = '" & IIf(dcValIndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmTCpbDet.dtpFehOpe, frmTCpbDet.dtpFeEDoc), frmTCpbDet.dtpFehOpe) & "'"

   ' [Adicional Agregado por Angel
         If .EOF Then
            MsgBox TEXT_9015, vbExclamation
            txtDato(8).Text = Format(0, FORMATO_NUM_2)
            txtDato(8).SetFocus
         Else
            txtDato(8).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
         End If
   ']
      Else
         txtDato(8).Text = Format(0, FORMATO_NUM_2)
      End If
   End With
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
   '[ Agregado por Angel
   If Val(txtDato(8).Text) = 0 Then
      txtDato(8).Text = Format(0, FORMATO_NUM_2)
      txtDato(8).SetFocus
      MsgBox TEXT_9015, vbExclamation
      Exit Sub
   End If
   txtImporte.Item(Index).SelStart = 0
   txtImporte.Item(Index).SelLength = txtImporte.Item(Index).MaxLength
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
   If Val(txtImporte(Index).Text) = 0 Then
      txtImporte(Index).Text = Format(0, FORMATO_NUM_1)
   End If
   
   Select Case Index
   Case 0
      If CDec(txtImporte(Index).Text) <> 0 Then
         txtImporte(1).Text = Format(0, FORMATO_NUM_1)
         txtImporte(3).Text = Format(0, FORMATO_NUM_1)
         If cboTpoMon.ListIndex = TPOMON_NAC_IND And (txtImporte(2).Text = 0 Or CDec(txtImporte(0).Text) <> CDec(txtImporte(0).Tag)) Then
            txtImporte(2).Text = Format(gfRedond(CDec(txtImporte(0).Text) / CDec(txtDato(8).Text), 2), FORMATO_NUM_1)
         End If
      End If
   Case 2
      If CDec(txtImporte(Index).Text) <> 0 Then
         txtImporte(1).Text = Format(0, FORMATO_NUM_1)
         txtImporte(3).Text = Format(0, FORMATO_NUM_1)
         If cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtImporte(0).Text) = 0 Or CDec(txtImporte(2).Text) <> CDec(txtImporte(2).Tag)) Then
            txtImporte(0).Text = Format(gfRedond(CDec(txtImporte(2).Text) * CDec(txtDato(8).Text), 2), FORMATO_NUM_1)
         End If
      End If
   Case 1
      If CDec(txtImporte(Index).Text) <> 0 Then
         txtImporte(0).Text = Format(0, FORMATO_NUM_1)
         txtImporte(2).Text = Format(0, FORMATO_NUM_1)
         If cboTpoMon.ListIndex = TPOMON_NAC_IND And (CDec(txtImporte(3).Text) = 0 Or CDec(txtImporte(1).Text) <> CDec(txtImporte(1).Tag)) Then
            txtImporte(3).Text = Format(gfRedond(CDec(txtImporte(1).Text) / CDec(txtDato(8).Text), 2), FORMATO_NUM_1)
         End If
      End If
   Case 3
      If txtImporte(Index).Text <> 0 Then
         txtImporte(0).Text = Format(0, FORMATO_NUM_1)
         txtImporte(2).Text = Format(0, FORMATO_NUM_1)
         If cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtImporte(1).Text) = 0 Or CDec(txtImporte(3).Text) <> CDec(txtImporte(3).Tag)) Then
            txtImporte(1).Text = Format(gfRedond(CDec(txtImporte(3).Text) * CDec(txtDato(8).Text), 2), FORMATO_NUM_1)
         End If
      End If
   End Select
  
   With frmTCpbCab
      .cmdCalcular_Click
      If pbNuevo Then
         txtDeta(0).Text = Format(CDec(.txtDeta(0).Text) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
         txtDeta(1).Text = Format(CDec(.txtDeta(1).Text) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
         txtDeta(2).Text = Format(CDec(.txtDeta(2).Text) + CDec(txtImporte(2).Text), FORMATO_NUM_1)
         txtDeta(3).Text = Format(CDec(.txtDeta(3).Text) + CDec(txtImporte(3).Text), FORMATO_NUM_1)
         .txtDeta(0).Text = Format(CDec(.txtDeta(0).Text) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
         .txtDeta(1).Text = Format(CDec(.txtDeta(1).Text) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
         .txtDeta(2).Text = Format(CDec(.txtDeta(2).Text) + CDec(txtImporte(2).Text), FORMATO_NUM_1)
         .txtDeta(3).Text = Format(CDec(.txtDeta(3).Text) + CDec(txtImporte(3).Text), FORMATO_NUM_1)
      Else
         txtDeta(0).Text = Format(CDec(.txtDeta(0).Text) - CDec(txtImporte(0).Tag) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
         txtDeta(1).Text = Format(CDec(.txtDeta(1).Text) - CDec(txtImporte(1).Tag) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
         txtDeta(2).Text = Format(CDec(.txtDeta(2).Text) - CDec(txtImporte(2).Tag) + CDec(txtImporte(2).Text), FORMATO_NUM_1)
         txtDeta(3).Text = Format(CDec(.txtDeta(3).Text) - CDec(txtImporte(3).Tag) + CDec(txtImporte(3).Text), FORMATO_NUM_1)
      End If
      
      txtImporte(Index).Tag = Format(CDec(txtImporte(Index).Text), FORMATO_NUM_1)
      txtImporte(Index).Text = Format(CDec(txtImporte(Index).Text), FORMATO_NUM_1)
   End With
End Sub

Private Sub ppHabilitaDatosDocumento()
   If Not frmTCpbGrd.uorstCOCta.EOF Then
      If Not optTpoPvs(1).Value Then
         txtDato(3).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         lblDatoDeta(3).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         cmdDatoAyud(3).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         txtDato(4).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         txtDato(5).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         dtpFeEDoc.Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         dtpFeVDoc.Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         dtpFeRDoc.Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         optTpoPvs(0).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         optTpoPvs(1).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
         optTpoPvs(2).Enabled = IIf(frmTCpbGrd.uorstCOCta!IndDoc = INDDOC_ACT, True, False)
      End If
   End If
End Sub

Private Sub CboTpoTCb_LostFocus()
   Dim dcValIndDoc As Integer
   
   dcValIndDoc = frmTCpbGrd.uorstCOCta!IndDoc
   
   With frmTCpbGrd.uorstTGTCb
      If .RecordCount <> 0 Then
         .MoveFirst
         .Find "FehTCb = '" & IIf(dcValIndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmTCpbDet.dtpFehOpe, frmTCpbDet.dtpFeEDoc), frmTCpbDet.dtpFehOpe) & "'"
         If .EOF Then
            MsgBox TEXT_9015, vbExclamation
            txtDato(8).Text = Format(0, FORMATO_NUM_2)
            txtDato(8).SetFocus
         Else
            txtDato(8).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
         End If
      Else
         txtDato(8).Text = Format(0, FORMATO_NUM_2)
      End If
   End With
End Sub

Private Sub ppGenera_AjD()
   Dim dnImpTCb_Pvs, dnImpTot_Can, dnImpTot_AjD, _
       dnImpTCb_Can, dnImpMN_Pvs, dnImpME_Pvs As Variant
   Dim dcTpoTCb_Pvs As String, dcTpoMon_Pvs As String, dcCodCta_Pvs As String, _
       dcTpoCtb_Can As String, dcTpoMon_Can As String, dcTpoCtb_AjD As String, _
       dcTpoCtb_Pvs As String, dcMes As String
   
   dnImpTCb_Pvs = CDec(dnImpTCb_Pvs)
   dnImpTot_Can = CDec(dnImpTot_Can)
   dnImpTot_AjD = CDec(dnImpTot_AjD)
   dnImpTCb_Can = CDec(dnImpTCb_Can)
   dnImpMN_Pvs = CDec(dnImpMN_Pvs)
   dnImpME_Pvs = CDec(dnImpME_Pvs)

   With frmTCpbGrd.uorstCOCpbDet
      .Source = "SELECT CodAux, CodTDc, SerDoc, NroDoc, ImpMN, ImpME, ImpTCb, TpoTCb, TpoMon, TpoPvs," _
              & "  TpoCtb, CodCta, FehOpe, FeEDoc, FeVDoc, FeRDoc," _
              & "  Concat(CodAux,CodTDc,SerDoc,NroDoc) as cLlave " _
              & "FROM CoCpbDet " _
              & "WHERE CodAux ='" & txtDato(2).Text & "'" _
              & "  AND CodCta ='" & txtDato(0).Text & "'" _
              & "  AND CodTDc='" & txtDato(3).Text & "'" _
              & "  AND SerDoc='" & txtDato(4).Text & "'" _
              & "  AND NroDoc='" & txtDato(5).Text & "'" _
              & "  AND TpoPvs='" & TPOPVS_PVS & "'"
      .Open
      If Not .EOF Then
         If .RecordCount > 1 Then
            MsgBox "Existe más de una provisión para el documento generado." & Chr(13) & "No se generará el Ajuste por Tipo de Cambio. Revise y hágalo manualmente.", vbExclamation
            .Close
            Exit Sub
         Else
            dcCodCta_Pvs = !CodCta
            dcTpoCtb_Pvs = !TpoCtb
            dcTpoMon_Pvs = !TpoMon
            dcTpoTCb_Pvs = !TpoTcb
            dnImpTCb_Pvs = CDec(!ImpTCb)
            dnImpMN_Pvs = CDec(!ImpMN)
            dnImpME_Pvs = CDec(!ImpME)
            .MoveFirst
            .Find "cLlave='" & txtDato(2).Text & txtDato(3).Text & txtDato(4).Text & txtDato(5).Text & "'"
            If Month(!FehOpe) <> Month(dtpFehOpe.Value) Then
               dcMes = gfCeros(Str(Month(dtpFehOpe.Value)), 2, -1, "0")
               With frmTCpbGrd.uorstCOTCbMes
                  .Source = "SELECT ImpTCb_Cpr, ImpTCb_Vta " _
                          & "FROM COTCbMes " _
                          & "WHERE MesPvs='" & dcMes & "'"
                  .Open
                  dnImpTCb_Pvs = CDec(IIf(dcTpoTCb_Pvs = TPOTCB_VTA, !ImpTCb_Vta, !ImpTCb_Cpr))
                  .Close
               End With
            End If
            If gnIndMNE = INDMNE_ACT Then
''              If (dcTpoMon_Pvs = TPOMON_EXT And gsTpoMon_Fnc = TPOMON_NAC) Or (dcTpoMon_Pvs = TPOMON_NAC And gsTpoMon_Fnc = TPOMON_EXT) Then
               If (dcTpoMon_Pvs = TPOMON_EXT And cboTpoMon.ListIndex = TPOMON_NAC_IND) _
                  Or (dcTpoMon_Pvs = TPOMON_NAC And cboTpoMon.ListIndex = TPOMON_EXT_IND) _
                  Or (dcTpoMon_Pvs = TPOMON_EXT And cboTpoMon.ListIndex = TPOMON_EXT_IND And CDec(txtDato(8).Text) <> dnImpTCb_Pvs) _
                  Or (dcTpoMon_Pvs = TPOMON_NAC And cboTpoMon.ListIndex = TPOMON_NAC_IND And CDec(txtDato(8).Text) <> dnImpTCb_Pvs) Then
                  dnImpTCb_Can = CDec(txtDato(8).Text)
'[REVISAR. Cambiado (21/3/04).
'                  If CDec(txtImporte(0).Text) > 0 Or CDec(txtImporte(1).Text) > 0 Then
                  If dcTpoMon_Pvs = TPOMON_EXT Then
']
                     dnImpTot_Can = IIf(CDec(txtImporte(0).Text) = 0, CDec(txtImporte(1).Text), CDec(txtImporte(0).Text))
                     dcTpoCtb_Can = IIf(CDec(txtImporte(0).Text) = 0, TPOCTB_HAB, TPOCTB_DEB)
                     dnImpTot_AjD = CDec(CDec(IIf(CDec(txtImporte(2).Text) = 0, CDec(txtImporte(3).Text), CDec(txtImporte(2).Text)) * dnImpMN_Pvs) / dnImpME_Pvs)
                     dcTpoCtb_AjD = IIf(dnImpTot_Can > dnImpTot_AjD, IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_DEB, TPOCTB_HAB), IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
                     dnImpTot_AjD = Abs((dnImpTot_Can - dnImpTot_AjD))
                     dcTpoMon_Can = TPOMON_NAC
                  Else
                     dnImpTot_Can = IIf(CDec(txtImporte(2).Text) = 0, CDec(txtImporte(3).Text), CDec(txtImporte(2).Text))
                     dcTpoCtb_Can = IIf(CDec(txtImporte(2).Text) = 0, TPOCTB_HAB, TPOCTB_DEB)
                     dnImpTot_AjD = CDec(CDec(IIf(CDec(txtImporte(0).Text) = 0, CDec(txtImporte(1).Text), CDec(txtImporte(0).Text)) * dnImpME_Pvs) / dnImpMN_Pvs)
                     dcTpoCtb_AjD = IIf(dnImpTot_Can > dnImpTot_AjD, IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_DEB, TPOCTB_HAB), IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
                     dnImpTot_AjD = Abs((dnImpTot_Can - dnImpTot_AjD))
                     dcTpoMon_Can = TPOMON_EXT
                  End If
'                  If pbNuevo And dnImpTot_AjD > 0 Then
                  If dnImpTot_AjD > 0 Then
                    'Generación de Item 1/2.
                     With frmTCpbGrd.uorstMain_1
                        .AddNew
                        !MesPvs = gsMesAct
                        !CodDro = frmTCpbCab.txtLlave(0).Text
                        !NroCpb = frmTCpbCab.txtLlave(1).Text
                        With frmTCpbGrd.uorstUltiItem
                           .Source = "SELECT IFNULL(MAX(NroIte), 0) AS cUltIte " _
                                   & "FROM COCpbDet " _
                                   & "WHERE CodDro='" & frmTCpbCab.txtLlave(0).Text & "' And NroCpb='" & frmTCpbCab.txtLlave(1).Text & "' And MesPvs='" & gsMesAct & "'"
                           .Open
                           frmTCpbGrd.uorstMain_1!NroIte = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
                           .Close
                        End With
                        !BlqIte = pnNroIte
                        !FehOpe = dtpFehOpe.Value
                        !CodCta = IIf(dcTpoCtb_AjD = TPOCTB_DEB, pcCodCta_AjD_Deb, pcCodCta_AjD_Hab)
                        !CodCCo = CODCCO_AJD
                        !GloIte = "Ajuste por Diferencia de Cambio"
                        !TpoPvs = TPOPVS_OTR
                        !FeEDoc = dtpFeEDoc.Value
                        !FeVDoc = dtpFeVDoc.Value
                        !FeRDoc = dtpFeRDoc.Value
                        !TpoMon = dcTpoMon_Can
                        !ImpTCb = dnImpTCb_Can
                        !ImpMN = IIf(dcTpoMon_Pvs = TPOMON_EXT, dnImpTot_AjD, 0)
                        !ImpME = IIf(dcTpoMon_Pvs = TPOMON_NAC, dnImpTot_AjD, 0)
                        !TpoCtb = IIf(dcTpoCtb_AjD = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
                        !TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
                        !TpoGnr = TPOGNR_DCA
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                        .Update
                        
                       'Generación de Item 2/2.
                        .AddNew
                        !CodDro = frmTCpbCab.txtLlave(0).Text
                        !NroCpb = frmTCpbCab.txtLlave(1).Text
                        !MesPvs = gsMesAct
                        With frmTCpbGrd.uorstUltiItem
                           .Source = "SELECT IFNULL(MAX(NroIte), 0) AS cUltIte " _
                                   & "FROM COCpbDet " _
                                   & "WHERE CodDro='" & frmTCpbCab.txtLlave(0).Text & "' And NroCpb='" & frmTCpbCab.txtLlave(1).Text & "' And MesPvs='" & gsMesAct & "'"
                           .Open
                           frmTCpbGrd.uorstMain_1!NroIte = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
                           .Close
                        End With
                        !BlqIte = pnNroIte
                        !FehOpe = dtpFehOpe.Value
                        !CodCta = dcCodCta_Pvs
                        !CodAux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
                        !CodTDc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
                        !SerDoc = txtDato(4).Text
                        !NroDoc = txtDato(5).Text
                        !RefDoc = txtDato(6).Text
                        !GloIte = "Ajuste por Diferencia de Cambio"
                        !FeEDoc = dtpFeEDoc.Value
                        !FeVDoc = dtpFeVDoc.Value
                        !FeRDoc = dtpFeRDoc.Value
                        !TpoPvs = TPOPVS_OTR
                        !TpoMon = dcTpoMon_Can
                        !ImpTCb = dnImpTCb_Can
                        !ImpMN = IIf(dcTpoMon_Pvs = TPOMON_EXT, dnImpTot_AjD, 0)
                        !ImpME = IIf(dcTpoMon_Pvs = TPOMON_NAC, dnImpTot_AjD, 0)
                        !TpoCtb = dcTpoCtb_AjD
                        !TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
                        !TpoGnr = TPOGNR_DCA
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                        .Update
                     End With
                  End If
               End If
            End If
         End If
      End If
      frmTCpbGrd.uorstCOCpbDet.Close
   End With
End Sub


Private Sub dtpFehOpe_LostFocus()
   If optTpoPvs(0).Value Then
      dtpFeEDoc.Value = dtpFehOpe.Value
      dtpFeRDoc.Value = dtpFehOpe.Value
      dtpFeVDoc.Value = dtpFehOpe.Value
   End If

   If Month(dtpFehOpe.Value) <> Val(gsMesAct) Or Year(dtpFehOpe.Value) <> Val(gsAnoAct) Then
      If Month(dtpFehOpe.Value) <> Val(gsMesAct) Then
         If Month(dtpFehOpe.Value) = 1 And gsMesAct = "00" Or Month(dtpFehOpe.Value) = 12 And gsMesAct = "13" Then
            Exit Sub
         End If
      End If
      MsgBox "La fecha debe ser del Mes y Año que provisiona.", vbExclamation
      dtpFehOpe.SetFocus
   End If
End Sub

Private Sub optTpoPvs_Click(Index As Integer)
   cmdDatoAyud(4).Enabled = (cmdGrabar.Enabled And Index = 1)
   dtpFeEDoc.Enabled = (Index = 0)
   dtpFeVDoc.Enabled = (Index = 0)
   dtpFeRDoc.Enabled = (Index = 0)
   If optTpoPvs(1).Value Then
       With frmTCpbGrd.uorstTGTCb
          If .RecordCount <> 0 Then
             .MoveFirst
             .Find "FehTCb = '" & frmTCpbDet.dtpFehOpe & "'"
             If .EOF Then
                MsgBox TEXT_9015, vbExclamation
                frmTCpbDet.txtDato(8).Text = Format(0, FORMATO_NUM_2)
                Index = Index - 1
                frmTCpbDet.txtDato(0).SetFocus
             Else
                frmTCpbDet.txtDato(8).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
             End If
            ']
          Else
             frmTCpbDet.txtDato(8).Text = Format(0, FORMATO_NUM_2)
          End If
       End With
   End If
   If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
      cmdRtcPcp.Enabled = True
   Else
      cmdRtcPcp.Enabled = False
   End If
End Sub

Public Sub ppRtcPcp(tnFase As Boolean)
   On Error GoTo Err
   
   Dim dnContador As Integer
   Dim dbIndRP As Boolean
   Dim dsCodTdc As String, dsSerDoc As String, dsNroDoc As String
   
   dbIndRP = False
   If (txtDato(3).Text = CODTDC_FAC Or txtDato(3).Text = CODTDC_NCR) And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" Then
      With frmTCpbGrd.uorstCOCpbDetRP    'Cambiar RecordSet.
         .Requery
         If .RecordCount > 0 Then
            .MoveFirst
            .Find "cLlave='" & gsMesAct & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1).Text & pnNroIte & "'"
         End If
         If Not .EOF Then
            dsCodTdc = !CodTDc_RtcPcp
            dsSerDoc = !SerDoc_RtcPcp
            dsNroDoc = !NroDoc_RtcPcp
            dbIndRP = True
         End If
      End With
   End If
   If dbIndRP Then
      With frmTCpbGrd                     'Cambiar Formulario de Grid.
           'Llaves.
            If tnFase Then
               .uorstMain_1.AddNew
               .uorstMain_1!CodDro = frmTCpbCab.txtLlave(0).Text   '.uorstMain_0!CodDro
               .uorstMain_1!NroCpb = frmTCpbCab.txtLlave(1).Text   '.uorstMain_0!NroCpb
               .uorstMain_1!MesPvs = gsMesAct
               .uorstMain_1!BlqIte = pnNroIte
               With .uorstUltiItem
                  .Source = "SELECT IFNULL(MAX(NroIte), 0) AS cUltIte " _
                          & "FROM COCpbDet " _
                          & "WHERE CodDro='" & frmTCpbCab.txtLlave(0).Text & "' And NroCpb='" & frmTCpbCab.txtLlave(1).Text & "' And MesPvs='" & gsMesAct & "'"
                  .Open
                  frmTCpbGrd.uorstMain_1!NroIte = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
                  .Close
               End With
            Else
               .uorstMain_1.Find "cLlave='" & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1) & "'"
            End If
           'Datos.
            .uorstMain_1!FehOpe = dtpFehOpe.Value
            .uorstMain_1!CodCta = IIf(dsCodTdc = gsCodTDc_Rtc, gsCodCta_Rtc, gsCodCta_Pcp)
            .uorstMain_1!CodCCo = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
            .uorstMain_1!CodAux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
            .uorstMain_1!CodTDc = dsCodTdc
            .uorstMain_1!SerDoc = dsSerDoc
            .uorstMain_1!NroDoc = dsNroDoc
            .uorstMain_1!RefDoc = txtDato(6).Text
            .uorstMain_1!FeEDoc = dtpFeEDoc.Value
            .uorstMain_1!FeVDoc = dtpFeVDoc.Value
            .uorstMain_1!FeRDoc = dtpFeRDoc.Value
            .uorstMain_1!GloIte = txtDato(7).Text
            .uorstMain_1!TpoPvs = TPOPVS_PVS
            .uorstMain_1!TpoMon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
            .uorstMain_1!ImpTCb = txtDato(8).Text
            .uorstMain_1!ImpMN = gfRedond(CDec(IIf(txtImporte(0).Text <> 0, txtImporte(0).Text, txtImporte(1).Text)) * (IIf(dsCodTdc = gsCodTDc_Rtc, gnPctRtc, gnPctPcp) / 100), 2)
            .uorstMain_1!ImpME = gfRedond((IIf(txtImporte(2).Text <> 0, txtImporte(2).Text, txtImporte(3).Text)) * (IIf(dsCodTdc = gsCodTDc_Rtc, gnPctRtc, gnPctPcp) / 100), 2)
            .uorstMain_1!TpoCtb = IIf(dsCodTdc = gsCodTDc_Rtc, IIf(txtImporte(0).Text = 0, TPOCTB_DEB, TPOCTB_HAB), IIf(txtImporte(0).Text = 0, TPOCTB_HAB, TPOCTB_DEB))
            .uorstMain_1!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
            .uorstMain_1!TpoGnr = TPOGNR_DRO
            ']
            If tnFase Then
               .uorstMain_1!UsrCre = gsAbvUsr
               .uorstMain_1!FyHCre = Now
            Else
               .uorstMain_1!UsrMdf = gsAbvUsr
               .uorstMain_1!FyHMdf = Now
            End If
            .uorstMain_1.Update
      End With
   End If
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property

Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
   
   'Orden: Corregir.
   zaOpciones = Array(gbPms02)
End Property

Public Property Get zaOpciones() As Variant
End Property

Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdCorregir.Enabled = IIf(pbNuevo, False, taOpciones(0))
End Property

