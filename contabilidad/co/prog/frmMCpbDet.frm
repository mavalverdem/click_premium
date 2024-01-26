VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMCpbDet 
   Caption         =   "[Entidad Tipo Asiento]"
   ClientHeight    =   6525
   ClientLeft      =   2025
   ClientTop       =   1500
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleMode       =   0  'User
   ScaleWidth      =   13606.92
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5680
      Width           =   3480
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
         Picture         =   "frmMCpbDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
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
         Picture         =   "frmMCpbDet.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   54
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
         Picture         =   "frmMCpbDet.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   60
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
         Picture         =   "frmMCpbDet.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   60
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
         Left            =   60
         Picture         =   "frmMCpbDet.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   338
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
         Left            =   60
         Picture         =   "frmMCpbDet.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   60
         Width           =   360
      End
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
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   1045
      Width           =   1275
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   2
      Left            =   7635
      Picture         =   "frmMCpbDet.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1045
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
      Height          =   300
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   700
      Width           =   615
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   5175
      Picture         =   "frmMCpbDet.frx":0996
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   700
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
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   355
      Width           =   975
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   7335
      Picture         =   "frmMCpbDet.frx":0B40
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   355
      Width           =   255
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      ForeColor       =   &H80000002&
      Height          =   1830
      Left            =   60
      TabIndex        =   36
      Top             =   1735
      Width           =   7875
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
         Index           =   5
         Left            =   1830
         TabIndex        =   10
         Top             =   1050
         Width           =   1155
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
         Top             =   1050
         Width           =   525
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
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Top             =   690
         Width           =   315
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   6135
         Picture         =   "frmMCpbDet.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   690
         Width           =   255
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
         Height          =   210
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   350
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
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   350
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "P&endientes"
         Height          =   375
         Index           =   4
         Left            =   3360
         Picture         =   "frmMCpbDet.frx":0E94
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
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
         Height          =   210
         Index           =   2
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdRtcPcp 
         Caption         =   "&Reten./Perce."
         Height          =   375
         Left            =   6600
         TabIndex        =   38
         Top             =   645
         Width           =   1155
      End
      Begin VB.CheckBox chkGnr_RP 
         Alignment       =   1  'Right Justify
         Caption         =   "Genera Reten/Percep."
         Height          =   200
         Left            =   5550
         TabIndex        =   37
         Top             =   1125
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpFeEDoc 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   1395
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         OLEDropMode     =   1
         Format          =   66519041
         CurrentDate     =   37959.8076041667
      End
      Begin MSComCtl2.DTPicker dtpFeVDoc 
         Height          =   300
         Left            =   3900
         TabIndex        =   12
         Top             =   1395
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   66519041
         CurrentDate     =   37962.5159027778
      End
      Begin MSComCtl2.DTPicker dtpFeRDoc 
         Height          =   300
         Left            =   6360
         TabIndex        =   13
         Top             =   1395
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   66519041
         CurrentDate     =   37962.5159722222
      End
      Begin VB.Label lblTexto 
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
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblTexto 
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
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1200
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
         Index           =   3
         Left            =   1620
         TabIndex        =   44
         Top             =   690
         Width           =   4515
      End
      Begin VB.Label lblTexto 
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
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   1425
         Width           =   1080
      End
      Begin VB.Label lblTexto 
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
         Index           =   8
         Left            =   2820
         TabIndex        =   42
         Top             =   1425
         Width           =   930
      End
      Begin VB.Label lblTexto 
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
         Index           =   9
         Left            =   5400
         TabIndex        =   41
         Top             =   1425
         Width           =   810
      End
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
      Height          =   300
      Index           =   3
      Left            =   6300
      TabIndex        =   24
      Top             =   5080
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
      Height          =   300
      Index           =   2
      Left            =   4500
      TabIndex        =   22
      Top             =   5080
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
      Height          =   300
      Index           =   1
      Left            =   6300
      TabIndex        =   23
      Top             =   4720
      Width           =   1575
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
      Height          =   300
      Index           =   0
      Left            =   4500
      TabIndex        =   21
      Top             =   4720
      Width           =   1575
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5080
      Width           =   675
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
      Index           =   6
      Left            =   1080
      TabIndex        =   14
      Top             =   3655
      Width           =   6435
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
      Height          =   300
      Index           =   8
      Left            =   3180
      TabIndex        =   20
      Top             =   5080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H80000002&
      Height          =   975
      Left            =   3405
      TabIndex        =   29
      Top             =   5440
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
         Index           =   3
         Left            =   2700
         TabIndex        =   33
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
         Index           =   2
         Left            =   900
         TabIndex        =   32
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
         Index           =   1
         Left            =   2700
         TabIndex        =   31
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
         Index           =   0
         Left            =   900
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   195
         Width           =   1755
      End
      Begin VB.Label lblTexto 
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
         Index           =   19
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lblTexto 
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
         Index           =   18
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.ComboBox cboTpoTCb 
      Height          =   315
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5080
      Width           =   915
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "&Auxiliar"
      Height          =   375
      Left            =   60
      TabIndex        =   28
      Top             =   4900
      Width           =   1215
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
      Index           =   10
      Left            =   1080
      TabIndex        =   16
      Top             =   4345
      Width           =   1695
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   5
      Left            =   5175
      Picture         =   "frmMCpbDet.frx":103E
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1390
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
      Height          =   300
      Index           =   9
      Left            =   840
      TabIndex        =   4
      Top             =   1390
      Width           =   615
   End
   Begin VB.CommandButton cmdMasFjo 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7635
      Picture         =   "frmMCpbDet.frx":11E8
      TabIndex        =   26
      Top             =   1390
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
      Height          =   300
      Index           =   7
      Left            =   1080
      TabIndex        =   15
      Top             =   4000
      Width           =   6435
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
      Index           =   11
      Left            =   2880
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpFehOpe 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   66519041
      CurrentDate     =   37924.6695138889
   End
   Begin VB.Label lblTexto 
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
      Index           =   0
      Left            =   120
      TabIndex        =   73
      Top             =   40
      Width           =   1515
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   72
      Top             =   1075
      Width           =   585
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
      Left            =   2100
      TabIndex        =   71
      Top             =   1045
      Width           =   5535
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   70
      Top             =   730
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
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   69
      Top             =   700
      Width           =   3735
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
      Left            =   1800
      TabIndex        =   68
      Top             =   355
      Width           =   5535
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   67
      Top             =   385
      Width           =   555
   End
   Begin VB.Label lblTexto 
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
      Index           =   17
      Left            =   4080
      TabIndex        =   66
      Top             =   5110
      Width           =   345
   End
   Begin VB.Label lblTexto 
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
      Index           =   16
      Left            =   4080
      TabIndex        =   65
      Top             =   4750
      Width           =   360
   End
   Begin VB.Label lblTexto 
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
      Index           =   15
      Left            =   6840
      TabIndex        =   64
      Top             =   4480
      Width           =   435
   End
   Begin VB.Label lblTexto 
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
      Index           =   14
      Left            =   5100
      TabIndex        =   63
      Top             =   4480
      Width           =   375
   End
   Begin VB.Label lblTexto 
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
      Index           =   12
      Left            =   1440
      TabIndex        =   62
      Top             =   4750
      Width           =   840
   End
   Begin VB.Label lblTexto 
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
      Index           =   10
      Left            =   120
      TabIndex        =   61
      Top             =   3685
      Width           =   465
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
      ForeColor       =   &H80000002&
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   60
      Top             =   4750
      Width           =   1410
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Referencias :"
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
      Index           =   11
      Left            =   120
      TabIndex        =   59
      Top             =   4375
      Width           =   975
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
      Index           =   4
      Left            =   1440
      TabIndex        =   58
      Top             =   1390
      Width           =   3735
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   57
      Top             =   1420
      Width           =   495
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
      Index           =   20
      Left            =   120
      TabIndex        =   56
      Top             =   4030
      Width           =   900
   End
End
Attribute VB_Name = "frmMCpbDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'sirve para darle el mes 00
'por defecto todo hira ahi
Public rcMesAct

Private pbNuevo As Boolean
Private pbValidada As Boolean
Private pnCta_IndCCo As Integer, _
        pnCta_IndAjD As Integer, _
        pnCta_IndDoc As Integer, _
        pnCta_IndAnl As Integer, _
        pnCta_IndFjo As Integer
Private pcCodCta_AjD_Deb As String, pcCodCta_AjD_Hab As String
Private pcCodCCo_AjD_Deb As String, pcCodCCo_AjD_Hab As String, pcCodCCo_Def As String
Private pnItemCpb As Integer
Public pnUltIte, pnTpoMon As Integer
Public pnNroIte As Integer
Public psGlosa As String, psGlosax As String
Public pbHayRtcPcp As Boolean
'[ Variables de retención / percepción
Public psTpoDocRP As String, _
       psSerDocRP As String, _
       psNroDocRP As String
Public pnImpMNRP As Double, _
       pnImpMERP As Double, _
       pnImpDcMNRP As Double, _
       pnImpDcMERP As Double
'[ Indicadores de flujo de caja
Private Const INDMASFJO_INI As Byte = 0, _
              INDMASFJO_MAS As Byte = 1, _
              INDMASFJO_DET As Byte = 2

Private Sub cmdMasFjo_Click()
  frmMCpbDetMasGrd.Show vbModal
End Sub

Private Sub Form_Load()
'rcMesAct = "00"
'2014-03-31 error de validacion fecha cance prov
'rcMesAct = "01"
rcMesAct = "12"

    pbValidada = False

   Me.KeyPreview = True
   
   With frmMCpbGrd                     'Cambiar Formulario de Grid.
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
      txtDato(0).MaxLength = .uorstMain_1![codcta].DefinedSize
      txtDato(1).MaxLength = .uorstMain_1![codcco].DefinedSize
      txtDato(2).MaxLength = .uorstMain_1![codaux].DefinedSize
      txtDato(3).MaxLength = .uorstMain_1![codtdc].DefinedSize
      txtDato(4).MaxLength = .uorstMain_1![serdoc].DefinedSize
      txtDato(5).MaxLength = .uorstMain_1![nrodoc].DefinedSize
      txtDato(10).MaxLength = .uorstMain_1![refdoc].DefinedSize
      txtDato(11).MaxLength = .uorstMain_1!pdocpr.DefinedSize
      txtDato(gsIdioma + 5).MaxLength = .uorstMain_1![GloIte].DefinedSize
      txtDato(8 - gsIdioma).MaxLength = .uorstMain_1![GloItex].DefinedSize
      txtDato(8).MaxLength = 7
      txtImporte(0).MaxLength = 14
      txtImporte(1).MaxLength = 14
      txtImporte(2).MaxLength = 14
      txtImporte(3).MaxLength = 14
      txtDeta(0).Text = Format(frmMCpbCab.txtDeta(0).Text, FORMATO_NUM_1)
      txtDeta(1).Text = Format(frmMCpbCab.txtDeta(1).Text, FORMATO_NUM_1)
      txtDeta(2).Text = Format(frmMCpbCab.txtDeta(2).Text, FORMATO_NUM_1)
      txtDeta(3).Text = Format(frmMCpbCab.txtDeta(3).Text, FORMATO_NUM_1)
      psGlosa = frmMCpbCab.txtDato(0).Text
      psGlosax = frmMCpbCab.txtDato(1).Text
      pnTpoMon = TPOMON_NAC_IND
      psTpoDocRP = "": psSerDocRP = "": psNroDocRP = ""
      pnImpMNRP = 0: pnImpMERP = 0: pnImpDcMNRP = 0: pnImpDcMERP = 0
'2014-03-31 error al elgir un comprobante con cta cte.
'      With dtpFehOpe
'         .MinDate = DateAdd("m", -5, CDate("01/" & rcgfMesAct(IIf(rcMesAct = "00", "01", rcMesAct)) & "/" & gsAnoAct))
'         .MaxDate = gfUltDia(CDate("01/" & rcgfMesAct(IIf(rcMesAct = "00", "01", rcMesAct)) & "/" & gsAnoAct))
'      End With
      With dtpFehOpe
         .MinDate = DateAdd("m", -5, CDate("01/" & rcgfMesAct(IIf(rcMesAct = "00", "12", rcMesAct)) & "/" & gsAnoAct))
         .MaxDate = gfUltDia(CDate("01/" & rcgfMesAct(IIf(rcMesAct = "00", "12", rcMesAct)) & "/" & gsAnoAct))
      End With
      dtpFehOpe.Value = frmMCpbCab.dtpFehCpb.Value

    ']
   End With
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(20, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Fecha de Operación:", "Cuenta:", "C.Costo:", "Auxiliar:", "F.Caja:", "Tipo Documento:", "NºDocumento:", "Fecha Emisión:", "Vencimiento:", "Recepción:", "Glosa:", "Referencias :", "Mon. Func.:", "Tipo de Cambio:", "Debe", "Haber", "M.N.:", "M.E.:", "Total M.N.:", "Total M.E.:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Operation Date:", "Account:", "C.Center:", "Auxiliary:", "Cash F.:", "Type Document:", "NºDocument:", "Issue Date:", "Due Date:", "Receiving:", "Gloss:", "References :", "Func. Curr.:", "Rate of Exchange:", "Debit", "Credit", "N.C.:", "F.C.:", "Total N.C.:", "Total F.C.:")
  Next nElemento
  fraDocumento.Caption = Choose(gsIdioma, " Documento ", " Document ")
  optTpoPvs(0).Caption = Choose(gsIdioma, "Pro&visión", "Pro&vision")
  optTpoPvs(1).Caption = Choose(gsIdioma, "&Cancelación", "&Cancelation")
  optTpoPvs(2).Caption = Choose(gsIdioma, "&Otro", "&Other")
  cmdDatoAyud(4).Caption = Choose(gsIdioma, "P&endientes", "O&utstanding")
  cmdRtcPcp.Caption = Choose(gsIdioma, "&Reten./Perce.", "&Withh./Prec.")
  chkGnr_RP.Caption = Choose(gsIdioma, "Genera Reten/Percep.", "Generate Withh./Precep.")
  cmdAuxiliar.Caption = Choose(gsIdioma, "&Auxiliar", "&Auxiliary")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']
   
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
   If Trim(txtDato(0).Text) <> "" And Trim(lblDatoDeta(0).Caption) <> "" Then
      ppAyuDet 0
      pnCta_IndDoc = frmMCpbGrd.uorstCoCta!IndDoc
      pnCta_IndAjD = frmMCpbGrd.uorstCoCta!IndAjD
      pnCta_IndCCo = frmMCpbGrd.uorstCoCta!indcco
      pnCta_IndAnl = frmMCpbGrd.uorstCoCta!TpoAnl
      pnCta_IndFjo = frmMCpbGrd.uorstCoCta!IndFjo
      pcCodCCo_Def = IIf(IsNull(frmMCpbGrd.uorstCoCta!codcco_def), "", frmMCpbGrd.uorstCoCta!codcco_def)
      pcCodCta_AjD_Deb = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCta_AjD_Deb), "", frmMCpbGrd.uorstCoCta!CodCta_AjD_Deb)
      pcCodCta_AjD_Hab = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCta_AjD_Hab), "", frmMCpbGrd.uorstCoCta!CodCta_AjD_Hab)
      pcCodCCo_AjD_Deb = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCCo_AjD_Deb), "", frmMCpbGrd.uorstCoCta!CodCCo_AjD_Deb)
      pcCodCCo_AjD_Hab = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCCo_AjD_Hab), "", frmMCpbGrd.uorstCoCta!CodCCo_AjD_Hab)
      ' Actualiza los datos de centro de costo
      txtDato(1).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(0).Enabled)
      cmdDatoAyud(1).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(0).Enabled)
   End If
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
   If txtDato(3).Text <> "" Then ppAyuDet 3
   If txtDato(9).Text <> "" Then ppAyuDet 9
 ']

   If Not pbNuevo Then
      If frmMCpbGrd.uorstMain_1.RecordCount > 0 And frmMCpbGrd.uorstMain_1!tpognr <> TPOGNR_DRO Then cmdCorregir.Enabled = False
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
   If frmMCpbGrd.uorstMain_1.RecordCount <> 0 Then
      frmMCpbGrd.uorstMain_1.CancelUpdate
   End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMCpbGrd.uorstMain_1, Me
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMCpbGrd.uorstMain_1, Me
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
   ' Cambio original Raul If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
   If Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
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
   ' Datos referente del flujo de caja
   txtDato(9).Enabled = (cmdMasFjo.Tag <> INDMASFJO_DET And pnCta_IndFjo = INDFJO_ACT)
   cmdMasFjo.Enabled = ((txtDato(9).Text = "" Or cmdMasFjo.Tag <> INDMASFJO_MAS) And pnCta_IndFjo = INDFJO_ACT)
   cmdDatoAyud(5).Enabled = (cmdMasFjo.Tag <> INDMASFJO_DET And pnCta_IndFjo = INDFJO_ACT)
 ']
End Sub

Public Sub cmdGrabar_Click()
 '  On Error GoTo Err

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
   If pnCta_IndAjD = INDAJD_ACT And pnCta_IndAnl = TPOANL_AUX And pnCta_IndDoc = INDAUX_ACT And Len(Trim(txtDato(2).Text)) = 0 Then
'     MsgBox "Debe registrar datos completos del documento.", vbExclamation
      MsgBox TEXT_6002, vbExclamation
      txtDato(2).SetFocus
      Exit Sub
   End If
   If pnCta_IndFjo = INDFJO_ACT And Len(Trim(txtDato(9).Text)) = 0 Then
      MsgBox TEXT_6002, vbExclamation
      txtDato(9).SetFocus
      Exit Sub
   End If
   
   If Len(Trim(txtDato(3).Text)) <> 0 And Len(Trim(txtDato(2).Text)) = 0 Then
'      MsgBox "Debe asignar el auxiliar para el documento registrado.", vbExclamation
      MsgBox TEXT_6002, vbExclamation
      txtDato(2).SetFocus
      Exit Sub
   End If
   If CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0 And CDec(txtImporte(2).Text) = 0 And CDec(txtImporte(3).Text) = 0 Then
      If MsgBox(Choose(gsIdioma, "Grabará el detalle sin asignar importes?", "Will safe detail without assign amounts?"), vbYesNo) = vbNo Then
          txtImporte(0).SetFocus
          Exit Sub
      Else
         dbSinImportes = True
      End If
   Else
      If CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0 Then
         If MsgBox(Choose(gsIdioma, "Grabará el detalle sin asignar importes en ", "Will safe detail without assign amounts in ") & TPOMON_NAC_TXT_2 & "?", vbYesNo) = 7 Then
             txtImporte(0).SetFocus
             Exit Sub
         End If
      End If
      If CDec(txtImporte(2).Text) = 0 And CDec(txtImporte(3).Text) = 0 Then
         If MsgBox(Choose(gsIdioma, "Grabará el detalle sin asignar importes en ", "Will safe detail without assign amounts in ") & TPOMON_EXT_TXT_2 & "?", vbYesNo) = 7 Then
             txtImporte(2).SetFocus
             Exit Sub
         End If
      End If
   End If
   If Not dbSinImportes Then
      If cboTpoMon.ListIndex = TPOMON_NAC_IND And (CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0) Then
         MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Nacional.", "You Must enter the amount in National Currency."), vbInformation
         txtImporte(0).SetFocus
         Exit Sub
      ElseIf cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtImporte(2).Text) = 0 And CDec(txtImporte(3).Text) = 0) Then
         MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Extranjera.", "You Must enter the amount in Foreign Currency."), vbInformation
         txtImporte(2).SetFocus
         Exit Sub
      End If
   End If
   
   If Len(Trim(txtDato(2).Text)) <> 0 And Len(Trim(txtDato(3).Text)) <> 0 And Len(Trim(txtDato(4).Text)) <> 0 And Len(Trim(txtDato(5).Text)) <> 0 Then
      With frmMCpbGrd.uorstcomacpbdet
         .Source = "SELECT CodAux, CodTDc, SerDoc, NroDoc, ImpMN, ImpME, ImpTCb, TpoTCb, TpoMon, TpoPvs, TpoCtb, CodCta, CodDro, NroCpb, MesPvs "
         '2014-03-31 debe mirar el diario .Source = .Source & "FROM comacpbdet "
         .Source = .Source & "FROM cocpbdet "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano ='" & gsAnoAct & "' "
         .Source = .Source & "AND CodAux ='" & txtDato(2).Text & "' "
         .Source = .Source & "AND CodCta ='" & txtDato(0).Text & "' "
         .Source = .Source & "AND CodTDc='" & txtDato(3).Text & "' "
         .Source = .Source & "AND SerDoc='" & txtDato(4).Text & "' "
         .Source = .Source & "AND NroDoc='" & txtDato(5).Text & "' "
         .Source = .Source & "AND TpoPvs<>'" & TPOPVS_OTR & "'"
         .Open
         dnImpMN = 0
         dnImpME = 0
         If optTpoPvs(0).Value Then
            frmMCpbGrd.uorstcomacpbdet.Find "TpoPvs='" & TPOPVS_PVS & "'"
            If Not frmMCpbGrd.uorstcomacpbdet.EOF Then
               If frmMCpbGrd.uorstcomacpbdet!coddro <> frmMCpbCab.txtLlave(0).Text Or frmMCpbGrd.uorstcomacpbdet!NroCpb <> frmMCpbCab.txtLlave(1).Text Then
                  MsgBox Choose(gsIdioma, "Ya está registrada la provision del documento.", "the provision of document is registered ."), vbExclamation
                  frmMCpbGrd.uorstcomacpbdet.Close
                  optTpoPvs(0).SetFocus
                  Exit Sub
               End If
            End If
         Else
            If pbNuevo And optTpoPvs(1).Value Then
               frmMCpbGrd.uorstcomacpbdet.Find "TpoPvs='" & TPOPVS_PVS & "'"
               If .EOF Then
                  MsgBox Choose(gsIdioma, "No está registrada la provisión del documento.", "the provision of document is not registered."), vbExclamation
                  frmMCpbGrd.uorstcomacpbdet.Close
                  optTpoPvs(1).SetFocus
                  Exit Sub
               Else
                  If frmMCpbGrd.uorstcomacpbdet!codcta <> txtDato(0).Text Then
                     MsgBox Choose(gsIdioma, "La cuenta de la cancelación no es igual a la de la provisión.", "The cancelation account is not the same of the provision."), vbExclamation
                     frmMCpbGrd.uorstcomacpbdet.Close
                     txtDato(0).SetFocus
                     Exit Sub
                  End If
                  If frmMCpbGrd.uorstcomacpbdet!TpoCtb = TPOCTB_DEB And (CDec(txtImporte(0).Text) > 0 Or CDec(txtImporte(2).Text) > 0) Then
                     MsgBox Choose(gsIdioma, "Revise la información. La provisión está registrada en el DEBE.", "You review information. The provision is registered in DEBIT."), vbExclamation
                     frmMCpbGrd.uorstcomacpbdet.Close
                     txtImporte(1).SetFocus
                     Exit Sub
                  End If
                  If frmMCpbGrd.uorstcomacpbdet!TpoCtb = TPOCTB_HAB And (CDec(txtImporte(1).Text) > 0 Or CDec(txtImporte(3).Text) > 0) Then
                     MsgBox Choose(gsIdioma, "Revise la información. La provisión está registrada en el HABER.", "You review information. The provision is registered in CREDIT."), vbExclamation
                     frmMCpbGrd.uorstcomacpbdet.Close
                     txtImporte(0).SetFocus
                     Exit Sub
                  End If
               End If
            End If
            If (Not .EOF) And optTpoPvs(1).Value Then
               dcTpoMon = frmMCpbGrd.uorstcomacpbdet!tpomon
               frmMCpbGrd.uorstcomacpbdet.MoveFirst
               Do
                  If frmMCpbGrd.uorstcomacpbdet!coddro & frmMCpbGrd.uorstcomacpbdet!NroCpb & frmMCpbGrd.uorstcomacpbdet!mespvs <> frmMCpbCab.txtLlave(0).Text & frmMCpbCab.txtLlave(1).Text & rcMesAct Then
                     dnImpMN = dnImpMN + IIf(frmMCpbGrd.uorstcomacpbdet!TpoPvs = TPOPVS_PVS, frmMCpbGrd.uorstcomacpbdet!ImpMN, frmMCpbGrd.uorstcomacpbdet!ImpMN * (-1))
                     dnImpME = dnImpME + IIf(frmMCpbGrd.uorstcomacpbdet!TpoPvs = TPOPVS_PVS, frmMCpbGrd.uorstcomacpbdet!ImpME, frmMCpbGrd.uorstcomacpbdet!ImpME * (-1))
                  End If
                  frmMCpbGrd.uorstcomacpbdet.MoveNext
               Loop Until .EOF
               If dcTpoMon = TPOMON_NAC Then
                  If CDec(txtImporte(0).Text) > 0 Then
                     If dnImpMN < CDec(txtImporte(0).Text) Then
                        MsgBox Choose(gsIdioma, "El monto de la cancelación es mayor al de la provisión.", "The cancelation amount is more  than provision."), vbExclamation
                        frmMCpbGrd.uorstcomacpbdet.Close
                        txtImporte(0).SetFocus
                        Exit Sub
                     End If
                  End If
                  If CDec(txtImporte(1).Text) > 0 Then
                     If dnImpMN < CDec(txtImporte(1).Text) Then
                        MsgBox Choose(gsIdioma, "El monto de la cancelación es mayor al de la provisión.", "The cancelation amount is more  than provision."), vbExclamation
                        frmMCpbGrd.uorstcomacpbdet.Close
                        txtImporte(1).SetFocus
                        Exit Sub
                     End If
                  End If
               Else
                  If CDec(txtImporte(2).Text) > 0 Then
                     If dnImpME < CDec(txtImporte(2).Text) Then
                        MsgBox Choose(gsIdioma, "El monto de la cancelación es mayor al de la provisión.", "The cancelation amount is more  than provision."), vbExclamation
                        frmMCpbGrd.uorstcomacpbdet.Close
                        txtImporte(2).SetFocus
                        Exit Sub
                     End If
                  End If
                  If CDec(txtImporte(3).Text) > 0 Then
                     If dnImpME < CDec(txtImporte(3).Text) Then
                        MsgBox Choose(gsIdioma, "El monto de la cancelación es mayor al de la provisión.", "The cancelation amount is more  than provision."), vbExclamation
                        frmMCpbGrd.uorstcomacpbdet.Close
                        txtImporte(3).SetFocus
                        Exit Sub
                     End If
                  End If
               End If
            End If
         End If
         frmMCpbGrd.uorstcomacpbdet.Close
      End With
   End If
      
   With frmMCpbGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain_1.AddNew
      Else
         .uorstMain_1.Find "cLlave='" & frmMCpbCab.txtLlave(0).Text & frmMCpbCab.txtLlave(1) & Trim(Str(pnUltIte)) & "'"
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
     ' [Propio del formulario.
     ' Actualizo los flujos de efectivo
     UpDatosFlujo
     'Generación por Percepciones/Retenciones.
     UpDatosRtcPcp
     
     'Ajustes por Diferencia de Cambio.
      If pnCta_IndAjD = INDAJD_ACT And optTpoPvs(1).Value Then
        'Eliminación de Ajuste.
         If Not pbNuevo Then
            With .uorstMain_1
               dvRegistro = .Bookmark
               If Not .RecordCount = 0 Then
                  .MoveFirst
                  Do
                     If !coddro = frmMCpbCab.txtLlave(0).Text And !NroCpb = frmMCpbCab.txtLlave(1).Text And !blqite = pnNroIte And !tpognr = TPOGNR_DCA Then
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
         frmMCpbCab.upDatosGrid
         
         If .uorstMain_1.RecordCount <> 0 Then
          '[Búsqueda de llave actual.     'Cambiar.
            .uorstMain_1.MoveFirst
            .uorstMain_1.Find "cLlave='" & frmMCpbCab.txtLlave(0).Text & frmMCpbCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
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
            .uorstMain_1.Find "cLlave='" & frmMCpbCab.txtLlave(0).Text & frmMCpbCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
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
      ' Inicializo el numero de item
      pnItemCpb = 0
   End With
      
   Exit Sub
Err:
   gpErrores
    ' Inicializo el numero de item
    pnItemCpb = 0
   frmMCpbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
    '[Propio del formulario.
   With frmMCpbCab
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
   Case 0, 1, 2, 3, 5
      If ((pnCta_IndCCo = INDCCO_ACT And Index = 1) Or (pnCta_IndFjo = INDFJO_ACT And Index = 5)) Or Index <> 1 Then
         txtDato(IIf(Index = 5, 9, Index)).SetFocus
      End If
    Case 4  ' Inserto los documentos agrupados a la tabla tempolral
        txtDato(Index).SetFocus
   End Select
   If ((pnCta_IndCCo = INDCCO_ACT And Index = 1) Or (pnCta_IndFjo = INDFJO_ACT And Index = 5)) Or Index <> 1 Then
      ppAyuBus IIf(Index = 5, 9, Index)
   End If
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
   If Index = 8 Then
      txtDato(Index).SelStart = 0
      txtDato(Index).SelLength = txtDato(Index).MaxLength + 1
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
'[ Se agrega para la eliminacion del dato de flujo de caja digitado directamente
  If Index = 0 Or Index = 9 Then
    If txtDato(9).Text = "" Then
      If (Not pbNuevo) And cmdMasFjo.Tag <> INDMASFJO_INI Then
        ' Elimino el flujo del detalle
        txtDato(Index).Tag = "DELETE FROM " & ps_Prefijo & "tmpcomacpbdetFjo "
        txtDato(Index).Tag = txtDato(Index).Tag & "WHERE codemp='" & frmMCpbGrd.uorstMain_1!codemp & "' "
        txtDato(Index).Tag = txtDato(Index).Tag & "AND pdoano='" & frmMCpbGrd.uorstMain_1!pdoano & "' "
        txtDato(Index).Tag = txtDato(Index).Tag & "AND MesPvs='" & frmMCpbGrd.uorstMain_1!mespvs & "' "
        txtDato(Index).Tag = txtDato(Index).Tag & "AND CodDro='" & frmMCpbGrd.uorstMain_1!coddro & "' "
        txtDato(Index).Tag = txtDato(Index).Tag & "AND NroCpb='" & frmMCpbGrd.uorstMain_1!NroCpb & "' "
        txtDato(Index).Tag = txtDato(Index).Tag & "AND NroIte=" & frmMCpbGrd.uorstMain_1!NroIte
        frmMCpbGrd.uocnnMain.Execute txtDato(Index).Tag
      End If
      cmdMasFjo.Tag = INDMASFJO_INI
    End If
    cmdMasFjo.Tag = IIf(txtDato(9).Text <> "" And cmdMasFjo.Tag = INDMASFJO_INI, INDMASFJO_MAS, cmdMasFjo.Tag)
    cmdMasFjo.Enabled = ((txtDato(Index).Text = "" Or cmdMasFjo.Tag <> INDMASFJO_MAS) And pnCta_IndFjo = INDFJO_ACT)
    txtDato(9).Enabled = (cmdMasFjo.Tag <> INDMASFJO_DET And pnCta_IndFjo = INDFJO_ACT)
    cmdDatoAyud(5).Enabled = (cmdMasFjo.Tag <> INDMASFJO_DET And pnCta_IndFjo = INDFJO_ACT)
  End If
']
   
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
        pnCta_IndDoc = frmMCpbGrd.uorstCoCta!IndDoc
        pnCta_IndAjD = frmMCpbGrd.uorstCoCta!IndAjD
        pnCta_IndCCo = frmMCpbGrd.uorstCoCta!indcco
        pnCta_IndAnl = frmMCpbGrd.uorstCoCta!TpoAnl
        pnCta_IndFjo = frmMCpbGrd.uorstCoCta!IndFjo
        pcCodCCo_Def = IIf(IsNull(frmMCpbGrd.uorstCoCta!codcco_def), "", frmMCpbGrd.uorstCoCta!codcco_def)
        pcCodCta_AjD_Deb = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCta_AjD_Deb), "", frmMCpbGrd.uorstCoCta!CodCta_AjD_Deb)
        pcCodCta_AjD_Hab = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCta_AjD_Hab), "", frmMCpbGrd.uorstCoCta!CodCta_AjD_Hab)
        pcCodCCo_AjD_Deb = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCCo_AjD_Deb), "", frmMCpbGrd.uorstCoCta!CodCCo_AjD_Deb)
        pcCodCCo_AjD_Hab = IIf(IsNull(frmMCpbGrd.uorstCoCta!CodCCo_AjD_Hab), "", frmMCpbGrd.uorstCoCta!CodCCo_AjD_Hab)
        
        ' Actualizo los datos adicionales
        txtDato(1).Text = IIf(txtDato(1).Text = "", pcCodCCo_Def, txtDato(1).Text)
        txtDato(1).Text = IIf(pnCta_IndCCo = INDCCO_ACT, txtDato(1).Text, "")
        lblDatoDeta(1).Caption = IIf(pnCta_IndCCo = INDCCO_ACT, lblDatoDeta(1).Caption, "")
        txtDato(1).Enabled = (pnCta_IndCCo = INDCCO_ACT)
        cmdDatoAyud(1).Enabled = (pnCta_IndCCo = INDCCO_ACT)
        txtDato(9).Enabled = (pnCta_IndFjo = INDFJO_ACT)
        cmdDatoAyud(5).Enabled = (pnCta_IndFjo = INDFJO_ACT)
        cmdMasFjo.Enabled = (pnCta_IndFjo = INDFJO_ACT)
      End If
   Case 1, 2, 3, 9
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   Case 8
      txtDato(Index).Text = Format(CDec(IIf(txtDato(Index).Text = "", 0, txtDato(Index).Text)), FORMATO_NUM_2)
   End Select
 
 '[Propio del formulario. - Agregado por Angel, jalado de formulario anterior
   If Index = 0 Or Index = 8 Then
       If Val(txtDato(8)) = 0 Then
         txtDato(8).Tag = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
         With frmMCpbGrd.uorstTGTCb
            If .RecordCount <> 0 Then
               .MoveFirst
               .Find "FehTCb = '" & IIf(pnCta_IndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmMCpbDet.dtpFehOpe, frmMCpbDet.dtpFeEDoc), frmMCpbDet.dtpFehOpe) & "'"
         ' [Adicional Agregado por Angel
               If .EOF Then
                  MsgBox TEXT_9015, vbExclamation
                  txtDato(8).Text = Format(0, FORMATO_NUM_2)
                  Index = Index - 1
                  txtDato(0).SetFocus
               Else
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
      ' Cambio original de Raul If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
      If Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
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
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
           ' Elimino los documentos del temporal
      cmdDatoAyud(tnIndex).Tag = "WHERE codemp='" & gsCodEmp & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND pdoano='" & gsAnoAct & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND CodCta='" & txtDato(0).Text & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND CodAux='" & txtDato(2).Text & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND UsrCre='" & gsAbvUsr & "'"
      frmTCpbGrd.uocnnMain.Execute "DELETE FROM codoctmp1 " & cmdDatoAyud(tnIndex).Tag
      frmTCpbGrd.uocnnMain.Execute "DELETE FROM codoctmp2 " & cmdDatoAyud(tnIndex).Tag
      ' Inserto los documentos  pendientes
      cmdDatoAyud(tnIndex).Tag = "INSERT INTO codoctmp1 (codemp, pdoano, codcta, codaux, codtdc, serdoc, nrodoc, usrcre, impdmn, imphmn, impdme, imphme, impsmn, impsme) "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT codemp, pdoano, CodCta, CodAux, CodTDc, SerDoc, NroDoc, '" & gsAbvUsr & "', "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END)), 2) AS DebeMN, "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END)), 2) AS HaberMN, "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END)), 2) AS DebeME, "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END)), 2) AS HaberME, "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END) - (CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END))), 2) AS SaldoMN, "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END) - (CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END))), 2) AS SaldoME "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM cocpbdet "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE codemp= '" & gsCodEmp & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND pdoano= '" & gsAnoAct & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND CodCta = '" & txtDato(0).Text & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND CodAux = '" & txtDato(2).Text & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND MesPvs <= '" & gsMesAct & "' "
      If ps_Plataforma = pSrvMySql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND (FeEDoc) <= '" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "' "
      ElseIf ps_Plataforma = pSrvSql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND FeEDoc <= CONVERT(smalldatetime, '" & Format(dtpFehOpe.Value, "dd/mm/yyyy") & "', 103) "
      End If
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "GROUP BY codemp, pdoano, CodCta, CodAux, CodTDc, SerDoc, NroDoc "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ORDER BY CodTDc, SerDoc, NroDoc"
      frmTCpbGrd.uocnnMain.Execute cmdDatoAyud(tnIndex).Tag
       
      ' Inserto los documentos provisión
      cmdDatoAyud(tnIndex).Tag = "INSERT INTO codoctmp2 (codemp, pdoano, mespvs, codcta, codaux, codtdc, serdoc, nrodoc, tpomon, feedoc, usrcre) "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT DISTINCT codemp, pdoano, mespvs, codcta, codaux, codtdc, serdoc, nrodoc, tpomon, feedoc, '" & gsAbvUsr & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM cocpbdet "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE codemp='" & gsCodEmp & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND pdoano='" & gsAnoAct & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND mespvs<= '" & gsMesAct & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND codcta='" & txtDato(0).Text & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND codaux='" & txtDato(2).Text & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND tpopvs='" & TPOPVS_PVS & "' "
      If ps_Plataforma = pSrvMySql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND (FeEDoc) <= '" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "' "
      ElseIf ps_Plataforma = pSrvSql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND FeEDoc <= CONVERT(smalldatetime, '" & Format(dtpFehOpe.Value, "dd/mm/yyyy") & "', 103) "
      End If
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ORDER BY codtdc, serdoc, nrodoc"
      frmTCpbGrd.uocnnMain.Execute cmdDatoAyud(tnIndex).Tag
       
      ' Filtro de seleccion
      cmdDatoAyud(tnIndex).Tag = "a.UsrCre = '" & gsAbvUsr & "' "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND (CASE b.TpoMon WHEN '" & TPOMON_NAC & "' THEN a.ImpSMN ELSE a.ImpSME END) <> 0 "
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND b.MesPvs <= '" & gsMesAct & "' "
      If ps_Plataforma = pSrvMySql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND b.FeEDoc <= ('" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "')"
      ElseIf ps_Plataforma = pSrvSql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND b.FeEDoc <= CONVERT(smalldatetime, '" & Format(dtpFehOpe.Value, "dd/mm/yyyy") & "', 103)"
      End If
     
      modAyuBus.Sal_Doc cmdDatoAyud(tnIndex).Tag, txtDato(3).Text & txtDato(4).Text & txtDato(5).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      ' Elimino los datos de la tabla temporal
      frmTCpbGrd.uocnnMain.Execute "DELETE FROM CODocTmp1 WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND UsrCre = '" & gsAbvUsr & "'"
      frmTCpbGrd.uocnnMain.Execute "DELETE FROM CODocTmp2 WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND UsrCre = '" & gsAbvUsr & "'"
      
      txtDato(3).Text = Left(frmOAyuBus.uvDato1, 2)
      txtDato(4).Text = Mid(frmOAyuBus.uvDato1, 3, pLenSerDoc)
      txtDato(5).Text = Mid(frmOAyuBus.uvDato1, 3 + pLenSerDoc)
      ' Obtengo los datos por default del documento
      With frmMCpbGrd.uorstCOTCbMes
        .Source = "SELECT FeEDoc, FeVDoc, FeRDoc, TpoMon "
        .Source = .Source & "FROM comacpbdet "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND CodCta='" & txtDato(0).Text & "' "
        .Source = .Source & "AND CodAux='" & txtDato(2).Text & "' "
        .Source = .Source & "AND CodTDc='" & txtDato(3).Text & "' "
        .Source = .Source & "AND SerDoc='" & txtDato(4).Text & "' "
        .Source = .Source & "AND NroDoc='" & txtDato(5).Text & "' "
        .Source = .Source & "AND TpoPvs='" & TPOPVS_PVS & "'"
        .Open
        If .RecordCount <> 0 Then
          dtpFeEDoc = Format(!feedoc, "dd/mm/yyyy")
          dtpFeVDoc = Format(!fevdoc, "dd/mm/yyyy")
          dtpFeRDoc = Format(!ferdoc, "dd/mm/yyyy")
          cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
        End If
        .Close
      End With
      If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
         txtImporte(IIf(frmOAyuBus.uvDato2 < 0, 0, 1)).Text = Abs(Val(frmOAyuBus.uvDato2))
      Else
         txtImporte(IIf(Val(frmOAyuBus.uvDato2) < 0, 2, 3)).Text = Abs(Val(frmOAyuBus.uvDato2))
      End If
      ' Cambio original de Raul If (txtDato(3).Text = "01" Or txtDato(3).Text = "07") And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
      If Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value Then
         cmdRtcPcp.Enabled = True
      Else
         cmdRtcPcp.Enabled = False
      End If
   Case 9                              'Cambiar (añadir índices).
      modAyuBus.Fjo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodFjo)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(4).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub
Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCpbGrd.uorstCoCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmMCpbGrd.uorstCoCta!detcta
         End If
      End With
   Case 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCpbGrd.uorstCoCCo
         .MoveFirst
         .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmMCpbGrd.uorstCoCCo!detcco
         End If
      End With
   Case 2
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCpbGrd.uorstTGAux
         .MoveFirst
         .Find "CodAux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmMCpbGrd.uorstTGAux!razAux
         End If
      End With
   Case 3
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCpbGrd.uorstTGTDc
         .MoveFirst
         .Find "CodTDc='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmMCpbGrd.uorstTGTDc!dettdc
         End If
      End With
   Case 9
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(4).Caption = ""
         Exit Function
      End If
      With frmMCpbGrd.uorstCOFjo
         .MoveFirst
         .Find "CodFjo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(4).Caption = " " & frmMCpbGrd.uorstCOFjo!DetFjo
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   Dim dnContador As Integer
   On Error GoTo Err

   With frmMCpbGrd                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain_1!codemp = gsCodEmp
            .uorstMain_1!pdoano = gsAnoAct
            .uorstMain_1!mespvs = rcMesAct
            .uorstMain_1!coddro = frmMCpbCab.txtLlave(0).Text
            .uorstMain_1!NroCpb = frmMCpbCab.txtLlave(1).Text
            ' Obtengo el numero de Item
            pnItemCpb = frmMCpbGrd.pfNumItemCpb(gsAnoAct, rcMesAct, frmMCpbCab.txtLlave(0).Text, frmMCpbCab.txtLlave(1).Text)
            pnNroIte = pnItemCpb
            .uorstMain_1!NroIte = pnNroIte
            .uorstMain_1!blqite = pnNroIte
         End If
        
        'Datos.
         .uorstMain_1!fehope = dtpFehOpe.Value
         .uorstMain_1!codcta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
         .uorstMain_1!codcco = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
         .uorstMain_1!codaux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
         .uorstMain_1!indfjo_det = cmdMasFjo.Tag
         .uorstMain_1!codtdc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
         .uorstMain_1!serdoc = txtDato(4).Text
         .uorstMain_1!nrodoc = txtDato(5).Text
         .uorstMain_1!refdoc = txtDato(10).Text
         .uorstMain_1!pdocpr = txtDato(11).Text
         .uorstMain_1!feedoc = dtpFeEDoc.Value
         .uorstMain_1!fevdoc = dtpFeVDoc.Value
         .uorstMain_1!ferdoc = dtpFeRDoc.Value
         .uorstMain_1!GloIte = IIf(txtDato(gsIdioma + 5).Text = "", Null, txtDato(gsIdioma + 5).Text)
         .uorstMain_1!GloItex = IIf(txtDato(8 - gsIdioma).Text = "", Null, txtDato(8 - gsIdioma).Text)
         psGlosa = txtDato(6).Text
         psGlosax = txtDato(7).Text
         .uorstMain_1!TpoPvs = IIf(optTpoPvs(0).Value, TPOPVS_PVS, IIf(optTpoPvs(1).Value, TPOPVS_CAN, TPOPVS_OTR))
         .uorstMain_1!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
         pnTpoMon = cboTpoMon.ListIndex
         .uorstMain_1!ImpTCb = txtDato(8).Text
         .uorstMain_1!ImpMN = CDec(IIf(txtImporte(0).Text <> 0, txtImporte(0).Text, txtImporte(1).Text))
         .uorstMain_1!ImpME = CDec(IIf(txtImporte(2).Text <> 0, txtImporte(2).Text, txtImporte(3).Text))
         'cambio  .uorstMain_1!TpoCtb = IIf(txtImporte(0).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
         .uorstMain_1!TpoCtb = IIf(txtImporte(0).Text = 0 And txtImporte(2).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
         .uorstMain_1!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
         .uorstMain_1!tpognr = TPOGNR_DRO
         .uorstMain_1!IndGnr_RP = IIf(chkGnr_RP.Value = vbUnchecked, 0, 1)
      Else
        'Llaves.
'         txtLlave(0).Text = .uorstMain_1!CodSvc
        
        'Datos.
         If .uorstMain_1.EOF Then .uorstMain_1.MoveLast
         pnUltIte = .uorstMain_1!NroIte
         pnNroIte = .uorstMain_1!blqite
         dtpFehOpe.Value = .uorstMain_1!fehope
         txtDato(0).Text = IIf(IsNull(.uorstMain_1!codcta), "", .uorstMain_1!codcta)
         txtDato(1).Text = IIf(IsNull(.uorstMain_1!codcco), "", .uorstMain_1!codcco)
         txtDato(2).Text = IIf(IsNull(.uorstMain_1!codaux), "", .uorstMain_1!codaux)
         txtDato(3).Text = IIf(IsNull(.uorstMain_1!codtdc), "", .uorstMain_1!codtdc)
         txtDato(4).Text = IIf(IsNull(.uorstMain_1!serdoc), "", .uorstMain_1!serdoc)
         txtDato(5).Text = IIf(IsNull(.uorstMain_1!nrodoc), "", .uorstMain_1!nrodoc)
         txtDato(10).Text = IIf(IsNull(.uorstMain_1!refdoc), "", .uorstMain_1!refdoc)
         txtDato(11).Text = IIf(IsNull(.uorstMain_1!pdocpr), "", .uorstMain_1!pdocpr)
         dtpFeEDoc.Value = IIf(IsNull(.uorstMain_1!feedoc), .uorstMain_1!fehope, .uorstMain_1!feedoc)
         dtpFeVDoc.Value = IIf(IsNull(.uorstMain_1!fevdoc), .uorstMain_1!fehope, .uorstMain_1!fevdoc)
         dtpFeRDoc.Value = IIf(IsNull(.uorstMain_1!ferdoc), .uorstMain_1!fehope, .uorstMain_1!ferdoc)
         txtDato(gsIdioma + 5).Text = IIf(IsNull(.uorstMain_1!GloIte), "", .uorstMain_1!GloIte)
         txtDato(8 - gsIdioma).Text = IIf(IsNull(.uorstMain_1!GloItex), "", .uorstMain_1!GloItex)
         optTpoPvs(0).Value = IIf(.uorstMain_1!TpoPvs = TPOPVS_PVS, TPOPVS_PVS_VER, TPOPVS_PVS_FAL)
         optTpoPvs(1).Value = IIf(.uorstMain_1!TpoPvs = TPOPVS_CAN, TPOPVS_CAN_VER, TPOPVS_CAN_FAL)
         optTpoPvs(2).Value = IIf(.uorstMain_1!TpoPvs = TPOPVS_OTR, TPOPVS_OTR_VER, TPOPVS_OTR_FAL)
         cboTpoMon.ListIndex = IIf(.uorstMain_1!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         cboTpoTCb.ListIndex = IIf(.uorstMain_1!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
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
               .Item(dnContador).Text = Format(frmMCpbCab.txtDeta.Item(dnContador).Text, FORMATO_NUM_1)
            Next
         End With
         ']
         '[ Mostrar flujo de caja
         cmdMasFjo.Tag = .uorstMain_1!indfjo_det
         txtDato(9).Text = ppRetornaFlujo(.uorstMain_1!NroIte)
         cmdRtcPcp.Enabled = False
         ' Actualizar indicador de flujo de caja de la cuenta
         pnCta_IndFjo = frmMCpbGrd.uorstCoCta!IndFjo
         ']
         '[ Mostrar los datos de retención y percepción
         ppRetornaRtcPcp (.uorstMain_1!NroIte)
         chkGnr_RP.Value = .uorstMain_1!IndGnr_RP
         ']
         ppAyuDet (0)
         ppAyuDet (1)
         ppAyuDet (2)
         ppAyuDet (3)
         ppAyuDet (9)
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

Private Sub UpDatosFlujo()
    
  Static sWhere As String, sSentencia As String
  Static nRegistros As Double
  
  ' Elimino e inserto los flujos de caja
  sWhere = "WHERE codemp='" & gsCodEmp & "' "
  sWhere = sWhere & "AND pdoano='" & gsAnoAct & "' "
  sWhere = sWhere & "AND MesPvs='" & rcMesAct & "' "
  sWhere = sWhere & "AND CodDro='" & frmMCpbCab.txtLlave(0).Text & "' "
  sWhere = sWhere & "AND NroCpb='" & frmMCpbCab.txtLlave(1).Text & "' "
  sWhere = sWhere & "AND NroIte=" & frmMCpbGrd.uorstMain_1!NroIte
  frmMCpbGrd.uocnnMain.Execute "DELETE FROM comacpbdetFjo " & sWhere, nRegistros
  If frmMCpbGrd.uorstMain_1!indfjo_det = INDMASFJO_MAS Then
    sSentencia = "SELECT * FROM comacpbdetFjo "
    sSentencia = sSentencia & sWhere & " ORDER BY NroOrd"
     With frmMCpbGrd.uorstCOFjoDet
       If .State = adStateOpen Then .Close
       .Source = sSentencia
       .Open
       .AddNew
       !codemp = gsCodEmp
       !pdoano = gsAnoAct
       !mespvs = rcMesAct
       !coddro = frmMCpbCab.txtLlave(0).Text
       !NroCpb = frmMCpbCab.txtLlave(1).Text
       !NroIte = frmMCpbGrd.uorstMain_1!NroIte
       !NroOrd = 1
       !CodFjo = frmMCpbDet.txtDato(9).Text
       !codcta = frmMCpbDet.txtDato(0).Text
       !TpoCtb = frmMCpbGrd.uorstMain_1!TpoCtb
       !ImpMN = frmMCpbGrd.uorstMain_1!ImpMN
       !ImpME = frmMCpbGrd.uorstMain_1!ImpME
       !UsrCre = gsAbvUsr
       !FyHCre = Now
       If Not pbNuevo Then
         !UsrMdf = gsAbvUsr
         !FyHMdf = Now
       End If
       .Update
    End With
  ElseIf frmMCpbGrd.uorstMain_1!indfjo_det = INDMASFJO_DET Then
    sSentencia = "INSERT INTO comacpbdetFjo "
    sSentencia = sSentencia & "SELECT * FROM " & ps_Prefijo & "tmpcomacpbdetFjo "
    sSentencia = sSentencia & sWhere & " ORDER BY NroOrd"
    frmMCpbGrd.uocnnMain.Execute sSentencia, nRegistros
  End If

End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
   Dim dnContador As Integer

  'Llaves.
'   txtLlave(0).Text = ""

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
         If dnContador = 6 Then .Item(dnContador).Text = psGlosa
         If dnContador = 7 Then .Item(dnContador).Text = psGlosax
         If dnContador = 8 Then .Item(dnContador).Tag = ""
      Next
   End With
   txtDato(8).Text = Format(0, FORMATO_NUM_2)
   cboTpoMon.ListIndex = pnTpoMon
   cboTpoTCb.ListIndex = TPOTCB_VTA_IND
   optTpoPvs.Item(0) = TPOPVS_PVS_FAL
   optTpoPvs.Item(1) = TPOPVS_CAN_FAL
   optTpoPvs.Item(2) = TPOPVS_OTR_VER
   dtpFeEDoc.Value = frmMCpbCab.dtpFehCpb.Value
   dtpFeVDoc.Value = frmMCpbCab.dtpFehCpb.Value
   dtpFeRDoc.Value = frmMCpbCab.dtpFehCpb.Value
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
   ' Indicador de flujo
   cmdMasFjo.Tag = INDMASFJO_INI
   pbHayRtcPcp = False: psTpoDocRP = "": psNroDocRP = "": psSerDocRP = ""
   pnImpMNRP = 0: pnImpMERP = 0: pnImpDcMNRP = 0: pnImpDcMERP = 0
  
  ' Inicializo detalle de flujo
  frmMCpbGrd.uocnnMain.Execute "DELETE FROM " & ps_Prefijo & "tmpcomacpbdetFjo"


End Sub

Private Sub UpDatosRtcPcp()
  On Error GoTo Err
    
  Dim sWhere As String, sSentencia As String
  Dim nRegistros As Double
  Dim dvRegistro As Variant
  
  sWhere = "WHERE codemp='" & gsCodEmp & "' "
  sWhere = sWhere & "AND pdoano='" & gsAnoAct & "' "
  sWhere = sWhere & "AND MesPvs='" & rcMesAct & "' "
  sWhere = sWhere & "AND CodDro='" & frmMCpbCab.txtLlave(0).Text & "' "
  sWhere = sWhere & "AND NroCpb='" & frmMCpbCab.txtLlave(1).Text & "' "
  sWhere = sWhere & "AND NroIte=" & frmMCpbGrd.uorstMain_1!NroIte
  ' Elimino el documento de percepción o retención
  With frmMCpbGrd.uorstcomacpbdetRP
    If Not (.EOF And .BOF) Or .RecordCount > 0 Then .MoveFirst
    .Find "cLlave='" & rcMesAct & frmMCpbCab.txtLlave(0).Text & frmMCpbCab.txtLlave(1).Text & frmMCpbGrd.uorstMain_1!NroIte & "'"
    If Not .EOF Then .Delete
  End With
  frmMCpbGrd.uocnnMain.Execute "DELETE FROM comacpbdetRP " & sWhere, nRegistros
  ' Inserto el documento de percepción o retención
  If pbHayRtcPcp Then
    ' Obtengo importe del impuesto
    pnImpMNRP = CDec(IIf(txtImporte(0).Text <> 0, txtImporte(0).Text, txtImporte(1).Text))
    pnImpMERP = CDec(IIf(txtImporte(2).Text <> 0, txtImporte(2).Text, txtImporte(3).Text))
    pnImpDcMNRP = IIf((gsIndPcp = "S" Or gsIndRtc = "S" Or (gsIndRtc = "N" And chkGnr_RP.Value = vbChecked)), pnImpMNRP, pnImpDcMNRP)
    pnImpDcMERP = IIf((gsIndPcp = "S" Or gsIndRtc = "S" Or (gsIndPcp = "N" And chkGnr_RP.Value = vbChecked)), pnImpMERP, pnImpDcMERP)
    ' Si es numero de comprobante R/P y no es agente y no es el diario
    If ((psTpoDocRP = gsCodTDc_Rtc) And (gsIndRtc = "S" Or (gsIndRtc = "N" And chkGnr_RP.Value = vbChecked))) Then
      pnImpMNRP = gfRedond((pnImpDcMNRP * (gnPctRtc / 100)), 2)
      pnImpMERP = gfRedond((pnImpDcMERP * (gnPctRtc / 100)), 2)
    ElseIf ((psTpoDocRP = gsCodTDc_Pcp) And (gsIndPcp = "S" Or (gsIndRtc = "N" And chkGnr_RP.Value = vbChecked))) Then
      pnImpMNRP = gfRedond((pnImpDcMNRP * (gnPctPcp / 100)), 2)
      pnImpMERP = gfRedond((pnImpDcMERP * (gnPctPcp / 100)), 2)
    End If
    With frmMCpbGrd.uorstcomacpbdetRP
      .AddNew
      !codemp = gsCodEmp
      !pdoano = gsAnoAct
      !mespvs = rcMesAct
      !coddro = frmMCpbCab.txtLlave(0).Text
      !NroCpb = frmMCpbCab.txtLlave(1).Text
      !NroIte = frmMCpbGrd.uorstMain_1!NroIte
      !CodTDc_RtcPcp = psTpoDocRP
      !SerDoc_RtcPcp = psSerDocRP
      !NroDoc_RtcPcp = psNroDocRP
      !feEDoc_RtcPcp = dtpFehOpe.Value
      !ImpMN_RtcPcp = CDec(pnImpMNRP)
      !ImpME_RtcPcp = CDec(pnImpMERP)
      !codaux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      !codcta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !codtdc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
      !serdoc = txtDato(4).Text
      !nrodoc = txtDato(5).Text
      !ImpMN = CDec(pnImpDcMNRP)
      !ImpME = CDec(pnImpDcMERP)
      !IndRtcPcp = IIf(psTpoDocRP = gsCodTDc_Rtc, gsIndRtc, gsIndPcp)
      !UsrCre = gsAbvUsr
      !FyHCre = Now
      If Not pbNuevo Then
        !UsrMdf = gsAbvUsr
        !FyHMdf = Now
      End If
      .Update
    End With
  End If
  ' Elimino e inserto la cuenta de retención y percepción
  With frmMCpbGrd                     'Cambiar Formulario de Grid.
    If Not pbNuevo Then
      With .uorstMain_1
        dvRegistro = .Bookmark
        If Not .RecordCount = 0 Then
          .MoveFirst
          .Find "cLlave='" & frmMCpbCab.txtLlave(0).Text & frmMCpbCab.txtLlave(1).Text & Trim(pnNroIte) & "'"
          If Not .EOF Then
            Do
              If !coddro = frmMCpbCab.txtLlave(0).Text And !NroCpb = frmMCpbCab.txtLlave(1).Text And !blqite = pnNroIte And !tpognr = TPOGNR_DRP Then
                .Delete
              End If
              .MoveNext
              If .EOF Then Exit Do
            Loop Until (.EOF Or Not (!coddro = frmMCpbCab.txtLlave(0).Text And !NroCpb = frmMCpbCab.txtLlave(1).Text And !blqite = pnNroIte And !tpognr = TPOGNR_DRP))
          End If
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
    If pbHayRtcPcp Then
      ' Llaves.
      .uorstMain_1.AddNew
      .uorstMain_1!codemp = gsCodEmp
      .uorstMain_1!pdoano = gsAnoAct
      .uorstMain_1!mespvs = rcMesAct
      .uorstMain_1!coddro = frmMCpbCab.txtLlave(0).Text
      .uorstMain_1!NroCpb = frmMCpbCab.txtLlave(1).Text
      .uorstMain_1!blqite = pnNroIte
      ' Obtengo o incremento el numero de item
      pnItemCpb = frmMCpbGrd.pfNumItemCpb(gsAnoAct, rcMesAct, frmMCpbCab.txtLlave(0).Text, frmMCpbCab.txtLlave(1).Text)
      .uorstMain_1!NroIte = pnItemCpb
      'Datos.
      .uorstMain_1!fehope = dtpFehOpe.Value
      .uorstMain_1!codcta = IIf(psTpoDocRP = gsCodTDc_Rtc, gsCodCta_Rtc, gsCodCta_Pcp)
      .uorstMain_1!codcco = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
      .uorstMain_1!codaux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      .uorstMain_1!indfjo_det = INDMASFJO_INI
      .uorstMain_1!codtdc = IIf(psTpoDocRP = "", Null, psTpoDocRP)
      .uorstMain_1!serdoc = psSerDocRP
      .uorstMain_1!nrodoc = psNroDocRP
      .uorstMain_1!refdoc = txtDato(10).Text
      .uorstMain_1!pdocpr = txtDato(11).Text
      .uorstMain_1!feedoc = dtpFeEDoc.Value
      .uorstMain_1!fevdoc = dtpFeVDoc.Value
      .uorstMain_1!ferdoc = dtpFeRDoc.Value
      .uorstMain_1!GloIte = IIf(txtDato(gsIdioma + 5).Text = "", Null, txtDato(gsIdioma + 5).Text)
      .uorstMain_1!GloItex = IIf(txtDato(8 - gsIdioma).Text = "", Null, txtDato(8 - gsIdioma).Text)
      .uorstMain_1!TpoPvs = TPOPVS_PVS
      .uorstMain_1!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      .uorstMain_1!ImpTCb = txtDato(8).Text
      .uorstMain_1!ImpMN = CDec(pnImpMNRP)
      .uorstMain_1!ImpME = CDec(pnImpMERP)
      .uorstMain_1!TpoCtb = IIf(txtImporte(0).Text = 0 And txtImporte(2).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
      .uorstMain_1!TpoCtb = IIf(psTpoDocRP = gsCodTDc_Rtc, IIf(txtImporte(0).Text = 0 And txtImporte(2).Text = 0, TPOCTB_DEB, TPOCTB_HAB), IIf(txtImporte(0).Text = 0 And txtImporte(2).Text = 0, TPOCTB_HAB, TPOCTB_DEB))
      .uorstMain_1!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
      .uorstMain_1!tpognr = TPOGNR_DRP
      .uorstMain_1!UsrCre = gsAbvUsr
      .uorstMain_1!FyHCre = Now
      If Not pbNuevo Then
        .uorstMain_1!UsrMdf = gsAbvUsr
        .uorstMain_1!FyHMdf = Now
      End If
      .uorstMain_1.Update
    End If
  End With
   
  Exit Sub
Err:
  gpErrores

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
  chkGnr_RP.Visible = (gsIndRtc = "N" Or gsIndPcp = "N")
  chkGnr_RP.Enabled = tbHabilitar
  cmdMasFjo.Enabled = False
End Sub

'[Código propio del formulario.
Private Sub cmdAuxiliar_Click()
  frmMAuxGrd.Show vbModal
  frmMCpbGrd.uorstTGAux.Requery
End Sub

Private Sub cmdRtcPcp_Click()
  frmMCpbDetRet.Show vbModal
End Sub

Private Sub dtpFeEDoc_LostFocus()
  Dim dcValIndDoc As Integer
  
  If IsNull(frmMCpbDet.dtpFeEDoc) Then
    MsgBox Choose(gsIdioma, "Verifique activación y registro de la Fecha de Emisión del documento.", "Check activation and register of Issue Date of document."), vbExclamation
    dtpFeEDoc.Enabled = Not dtpFeEDoc.Enabled
    dtpFeEDoc.SetFocus
    Exit Sub
  End If
  dcValIndDoc = frmMCpbGrd.uorstCoCta!IndDoc
  With frmMCpbGrd.uorstTGTCb
    If .RecordCount <> 0 Then
      .MoveFirst
      .Find "FehTCb = '" & IIf(dcValIndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmMCpbDet.dtpFehOpe, frmMCpbDet.dtpFeEDoc), frmMCpbDet.dtpFehOpe) & "'"
      
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
  
  With frmMCpbCab
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
  If Not frmMCpbGrd.uorstCoCta.EOF Then
    If Not optTpoPvs(1).Value Then
      txtDato(3).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      cmdDatoAyud(3).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      txtDato(4).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      txtDato(5).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      txtDato(9).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      dtpFeEDoc.Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      dtpFeVDoc.Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      dtpFeRDoc.Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      optTpoPvs(0).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      optTpoPvs(1).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
      optTpoPvs(2).Enabled = (frmMCpbGrd.uorstCoCta!IndDoc = INDDOC_ACT)
    End If
  End If
End Sub

Private Sub CboTpoTCb_LostFocus()
  Dim dcValIndDoc As Integer
   
  dcValIndDoc = frmMCpbGrd.uorstCoCta!IndDoc
  
  With frmMCpbGrd.uorstTGTCb
    If .RecordCount <> 0 Then
      .MoveFirst
      .Find "FehTCb = '" & IIf(dcValIndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmMCpbDet.dtpFehOpe, frmMCpbDet.dtpFeEDoc), frmMCpbDet.dtpFehOpe) & "'"
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

   With frmMCpbGrd.uorstcomacpbdet
    .Source = "SELECT CodAux, CodTDc, SerDoc, NroDoc, ImpMN, ImpME, ImpTCb, TpoTCb, TpoMon, TpoPvs, "
    .Source = .Source & "TpoCtb, CodCta, FehOpe, FeEDoc, FeVDoc, FeRDoc, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "Concat(CodAux,CodTDc,SerDoc,NroDoc)", "(CodAux+CodTDc+SerDoc+NroDoc)") & " AS cLlave "
    .Source = .Source & "FROM comacpbdet "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND CodAux ='" & txtDato(2).Text & "' "
    .Source = .Source & "AND CodCta ='" & txtDato(0).Text & "' "
    .Source = .Source & "AND CodTDc='" & txtDato(3).Text & "' "
    .Source = .Source & "AND SerDoc='" & txtDato(4).Text & "' "
    .Source = .Source & "AND NroDoc='" & txtDato(5).Text & "' "
    .Source = .Source & "AND TpoPvs='" & TPOPVS_PVS & "'"
    .Open
      If Not .EOF Then
         If .RecordCount > 1 Then
            MsgBox Choose(gsIdioma, "Existe más de una provisión para el documento generado.", "The generated document has more than a provision") & Chr(13) & "No se generará el Ajuste por Tipo de Cambio. Revise y hágalo manualmente.", vbExclamation
            .Close
            Exit Sub
         Else
            dcCodCta_Pvs = !codcta
            dcTpoCtb_Pvs = !TpoCtb
            dcTpoMon_Pvs = !tpomon
            dcTpoTCb_Pvs = !TpoTcb
            dnImpTCb_Pvs = CDec(!ImpTCb)
            dnImpMN_Pvs = CDec(!ImpMN)
            dnImpME_Pvs = CDec(!ImpME)
            .MoveFirst
            .Find "cLlave='" & txtDato(2).Text & txtDato(3).Text & txtDato(4).Text & txtDato(5).Text & "'"
            If Month(!fehope) <> Month(dtpFehOpe.Value) Then
               dcMes = gfCeros(Str(Month(dtpFehOpe.Value)), 2, -1, "0")
               With frmMCpbGrd.uorstCOTCbMes
                  .Source = "SELECT ImpTCb_Cpr, ImpTCb_Vta "
                  .Source = .Source & "FROM COTCbMes "
                  .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
                  .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
                  .Source = .Source & "AND MesPvs='" & dcMes & "'"
                  .Open
                  dnImpTCb_Pvs = CDec(IIf(dcTpoTCb_Pvs = TPOTCB_VTA, !ImpTCb_Vta, !ImpTCb_Cpr))
                  .Close
               End With
            End If
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
            
            If gnIndMNE = INDMNE_ACT Then
''              If (dcTpoMon_Pvs = TPOMON_EXT And gsTpoMon_Fnc = TPOMON_NAC) Or (dcTpoMon_Pvs = TPOMON_NAC And gsTpoMon_Fnc = TPOMON_EXT) Then
               If (dcTpoMon_Pvs = TPOMON_EXT And cboTpoMon.ListIndex = TPOMON_NAC_IND) _
                  Or (dcTpoMon_Pvs = TPOMON_NAC And cboTpoMon.ListIndex = TPOMON_EXT_IND) _
                  Or (dcTpoMon_Pvs = TPOMON_EXT And cboTpoMon.ListIndex = TPOMON_EXT_IND And CDec(txtDato(8).Text) <> dnImpTCb_Pvs) Or dnImpTot_AjD > 0 _
                  Or (dcTpoMon_Pvs = TPOMON_NAC And cboTpoMon.ListIndex = TPOMON_NAC_IND And CDec(txtDato(8).Text) <> dnImpTCb_Pvs) Or dnImpTot_AjD > 0 Then
                  dnImpTCb_Can = CDec(txtDato(8).Text)
'[REVISAR. Cambiado (21/3/04).
'                  If CDec(txtImporte(0).Text) > 0 Or CDec(txtImporte(1).Text) > 0 Then
             '     If dcTpoMon_Pvs = TPOMON_EXT Then
']
              '       dnImpTot_Can = IIf(CDec(txtImporte(0).Text) = 0, CDec(txtImporte(1).Text), CDec(txtImporte(0).Text))
              '       dcTpoCtb_Can = IIf(CDec(txtImporte(0).Text) = 0, TPOCTB_HAB, TPOCTB_DEB)
              '       dnImpTot_AjD = CDec(CDec(IIf(CDec(txtImporte(2).Text) = 0, CDec(txtImporte(3).Text), CDec(txtImporte(2).Text)) * dnImpMN_Pvs) / dnImpME_Pvs)
              '       dcTpoCtb_AjD = IIf(dnImpTot_Can > dnImpTot_AjD, IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_DEB, TPOCTB_HAB), IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
              '       dnImpTot_AjD = Abs((dnImpTot_Can - dnImpTot_AjD))
              '       dcTpoMon_Can = TPOMON_NAC
              '    Else
              '       dnImpTot_Can = IIf(CDec(txtImporte(2).Text) = 0, CDec(txtImporte(3).Text), CDec(txtImporte(2).Text))
              '       dcTpoCtb_Can = IIf(CDec(txtImporte(2).Text) = 0, TPOCTB_HAB, TPOCTB_DEB)
              '       dnImpTot_AjD = CDec(CDec(IIf(CDec(txtImporte(0).Text) = 0, CDec(txtImporte(1).Text), CDec(txtImporte(0).Text)) * dnImpME_Pvs) / dnImpMN_Pvs)
              '       dcTpoCtb_AjD = IIf(dnImpTot_Can > dnImpTot_AjD, IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_DEB, TPOCTB_HAB), IIf(dcTpoCtb_Pvs = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
              '       dnImpTot_AjD = Abs((dnImpTot_Can - dnImpTot_AjD))
              '       dcTpoMon_Can = TPOMON_EXT
              '    End If
'                  If pbNuevo And dnImpTot_AjD > 0 Then
                  If dnImpTot_AjD > 0 Then
                    'Generación de Item 1/2.
                     With frmMCpbGrd.uorstMain_1
                        .AddNew
                        !codemp = gsCodEmp
                        !pdoano = gsAnoAct
                        !mespvs = rcMesAct
                        !coddro = frmMCpbCab.txtLlave(0).Text
                        !NroCpb = frmMCpbCab.txtLlave(1).Text
                        ' Obtengo o incremento el numero de item
                        pnItemCpb = frmMCpbGrd.pfNumItemCpb(gsAnoAct, rcMesAct, frmMCpbCab.txtLlave(0).Text, frmMCpbCab.txtLlave(1).Text)
                        !NroIte = pnItemCpb
                        !blqite = pnNroIte
                        !fehope = dtpFehOpe.Value
                        !codcta = IIf(dcTpoCtb_AjD = TPOCTB_DEB, pcCodCta_AjD_Deb, pcCodCta_AjD_Hab)
                        !codcco = IIf(dcTpoCtb_AjD = TPOCTB_DEB, IIf(pcCodCCo_AjD_Deb = "", Null, pcCodCCo_AjD_Deb), IIf(pcCodCCo_AjD_Hab = "", Null, pcCodCCo_AjD_Hab))
                        !GloIte = "Ajuste por Diferencia de Cambio"
                        !GloItex = "Adjustment by Defference of Exchange"
                        !TpoPvs = TPOPVS_OTR
                        !feedoc = dtpFeEDoc.Value
                        !fevdoc = dtpFeVDoc.Value
                        !ferdoc = dtpFeRDoc.Value
                        !tpomon = dcTpoMon_Can
                        !ImpTCb = dnImpTCb_Can
                        !ImpMN = IIf(dcTpoMon_Pvs = TPOMON_EXT, dnImpTot_AjD, 0)
                        !ImpME = IIf(dcTpoMon_Pvs = TPOMON_NAC, dnImpTot_AjD, 0)
                        !TpoCtb = IIf(dcTpoCtb_AjD = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
                        !TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
                        !tpognr = TPOGNR_DCA
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                        .Update
                        
                       'Generación de Item 2/2.
                        .AddNew
                        !codemp = gsCodEmp
                        !pdoano = gsAnoAct
                        !coddro = frmMCpbCab.txtLlave(0).Text
                        !NroCpb = frmMCpbCab.txtLlave(1).Text
                        !mespvs = rcMesAct
                        ' Incremento el numero de item
                        pnItemCpb = frmMCpbGrd.pfNumItemCpb(gsAnoAct, rcMesAct, frmMCpbCab.txtLlave(0).Text, frmMCpbCab.txtLlave(1).Text)
                        !NroIte = pnItemCpb
                        !blqite = pnNroIte
                        !fehope = dtpFehOpe.Value
                        !codcta = dcCodCta_Pvs
                        !codcco = IIf(pcCodCCo_Def = "", Null, pcCodCCo_Def)
                        !codaux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
                        !codtdc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
                        !serdoc = txtDato(4).Text
                        !nrodoc = txtDato(5).Text
                        !refdoc = IIf(txtDato(10).Text = "", Null, txtDato(10).Text)
                        !pdocpr = IIf(txtDato(11).Text = "", Null, txtDato(11).Text)
                        !GloIte = "Ajuste por Diferencia de Cambio"
                        !GloItex = "Adjustment by Defference of Exchange"
                        !feedoc = dtpFeEDoc.Value
                        !fevdoc = dtpFeVDoc.Value
                        !ferdoc = dtpFeRDoc.Value
                        !TpoPvs = TPOPVS_OTR
                        !tpomon = dcTpoMon_Can
                        !ImpTCb = dnImpTCb_Can
                        !ImpMN = IIf(dcTpoMon_Pvs = TPOMON_EXT, dnImpTot_AjD, 0)
                        !ImpME = IIf(dcTpoMon_Pvs = TPOMON_NAC, dnImpTot_AjD, 0)
                        !TpoCtb = dcTpoCtb_AjD
                        !TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
                        !tpognr = TPOGNR_DCA
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                        .Update
                     End With
                  End If
               End If
            End If
         End If
      End If
      frmMCpbGrd.uorstcomacpbdet.Close
   End With
End Sub

Private Sub dtpFehOpe_LostFocus()
  If optTpoPvs(0).Value Then
    dtpFeEDoc.Value = dtpFehOpe.Value
    dtpFeRDoc.Value = dtpFehOpe.Value
    dtpFeVDoc.Value = dtpFehOpe.Value
  End If

  If Month(dtpFehOpe.Value) <> Val(rcMesAct) Or Year(dtpFehOpe.Value) <> Val(gsAnoAct) Then
    If Not ((Format(dtpFehOpe.Value, "yyyymmdd") < Format(dtpFehOpe.MinDate, "yyyymmdd")) Or (Format(dtpFehOpe.Value, "yyyymmdd") > Format(dtpFehOpe.MaxDate, "yyyymmdd"))) Then Exit Sub
    MsgBox Choose(gsIdioma, "La fecha debe ser del Rango permitido que se provisiona.", "The date must be in permited range that provision."), vbExclamation
    dtpFehOpe.SetFocus
  End If
End Sub

Private Sub optTpoPvs_Click(Index As Integer)
  cmdDatoAyud(4).Enabled = (cmdGrabar.Enabled And Index = 1)
  dtpFeEDoc.Enabled = (Index = 0)
  dtpFeVDoc.Enabled = (Index = 0)
  dtpFeRDoc.Enabled = (Index = 0)
  If optTpoPvs(1).Value Then
    With frmMCpbGrd.uorstTGTCb
      If .RecordCount <> 0 Then
        .MoveFirst
        .Find "FehTCb = '" & frmMCpbDet.dtpFehOpe & "'"
        If .EOF Then
          MsgBox TEXT_9015, vbExclamation
          frmMCpbDet.txtDato(8).Text = Format(0, FORMATO_NUM_2)
          Index = Index - 1
          frmMCpbDet.txtDato(0).SetFocus
        Else
          frmMCpbDet.txtDato(8).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
        End If
        ']
      Else
        frmMCpbDet.txtDato(8).Text = Format(0, FORMATO_NUM_2)
      End If
    End With
  End If
  ' Cambio original de Raul cmdRtcPcp.Enabled = ((txtDato(3).Text = CODTDC_FAC Or txtDato(3).Text = CODTDC_NCR) And Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value)
  cmdRtcPcp.Enabled = (Trim(txtDato(4).Text) <> "" And Trim(txtDato(5).Text) <> "" And optTpoPvs(1).Value)

End Sub

Private Function ppRetornaFlujo(nNroItem As Integer) As String
  Static sSentencia As String
  Static uorstCOTmp As ADODB.Recordset
  Set uorstCOTmp = New ADODB.Recordset
  
  ppRetornaFlujo = ""
  sSentencia = "INSERT INTO " & ps_Prefijo & "tmpcomacpbdetFjo "
  sSentencia = sSentencia & "SELECT * FROM comacpbdetFjo "
  sSentencia = sSentencia & "WHERE codemp='" & frmMCpbGrd.uorstMain_1!codemp & "' "
  sSentencia = sSentencia & "AND pdoano='" & frmMCpbGrd.uorstMain_1!pdoano & "' "
  sSentencia = sSentencia & "AND MesPvs='" & frmMCpbGrd.uorstMain_1!mespvs & "' "
  sSentencia = sSentencia & "AND CodDro='" & frmMCpbGrd.uorstMain_1!coddro & "' "
  sSentencia = sSentencia & "AND NroCpb='" & frmMCpbGrd.uorstMain_1!NroCpb & "' "
  sSentencia = sSentencia & "AND NroIte=" & nNroItem
  frmMCpbGrd.uocnnMain.Execute "DELETE FROM " & ps_Prefijo & "tmpcomacpbdetFjo"
  frmMCpbGrd.uocnnMain.Execute sSentencia
  With uorstCOTmp
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmMCpbGrd.uocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Source = "SELECT CodFjo FROM " & ps_Prefijo & "tmpcomacpbdetFjo "
    .Source = .Source & "ORDER BY NroIte, NroOrd"
    .Open
    If Not (.EOF And .BOF) Then ppRetornaFlujo = !CodFjo
    .Close
  End With
  Set uorstCOTmp = Nothing
  
End Function

Private Sub ppRetornaRtcPcp(nNroItem As Integer)
  Static sSentencia As String
  Static uorstCOTmp As ADODB.Recordset
  Set uorstCOTmp = New ADODB.Recordset
  
  psTpoDocRP = "": psSerDocRP = "": psNroDocRP = ""
  pnImpMNRP = 0: pnImpMERP = 0: pnImpDcMNRP = 0: pnImpDcMERP = 0
  sSentencia = "SELECT DISTINCT CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp, "
  sSentencia = sSentencia & "ImpMN_RtcPcp, ImpME_RtcPcp, ImpMN, ImpME, NroIte "
  sSentencia = sSentencia & "FROM comacpbdetRP "
  sSentencia = sSentencia & "WHERE codemp='" & frmMCpbGrd.uorstMain_1!codemp & "' "
  sSentencia = sSentencia & "AND pdoano='" & frmMCpbGrd.uorstMain_1!pdoano & "' "
  sSentencia = sSentencia & "AND MesPvs='" & frmMCpbGrd.uorstMain_1!mespvs & "' "
  sSentencia = sSentencia & "AND CodDro='" & frmMCpbGrd.uorstMain_1!coddro & "' "
  sSentencia = sSentencia & "AND NroCpb='" & frmMCpbGrd.uorstMain_1!NroCpb & "' "
  sSentencia = sSentencia & "AND NroIte=" & nNroItem & " "
  sSentencia = sSentencia & "ORDER BY NroIte"
  frmMCpbGrd.uocnnMain.Execute sSentencia
  With uorstCOTmp
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmMCpbGrd.uocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Source = sSentencia
    .Open
    If Not (.EOF And .BOF) Then
      psTpoDocRP = !CodTDc_RtcPcp
      psSerDocRP = !SerDoc_RtcPcp
      psNroDocRP = !NroDoc_RtcPcp
      pnImpMNRP = Format(IIf(IsNull(!ImpMN_RtcPcp), 0, !ImpMN_RtcPcp), FORMATO_NUM_1)
      pnImpMERP = Format(IIf(IsNull(!ImpME_RtcPcp), 0, !ImpME_RtcPcp), FORMATO_NUM_1)
      pnImpDcMNRP = Format(IIf(IsNull(!ImpMN), 0, !ImpMN), FORMATO_NUM_1)
      pnImpDcMERP = Format(IIf(IsNull(!ImpME), 0, !ImpME), FORMATO_NUM_1)
    End If
    .Close
  End With
  pbHayRtcPcp = (psTpoDocRP <> "")
  Set uorstCOTmp = Nothing
  
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


