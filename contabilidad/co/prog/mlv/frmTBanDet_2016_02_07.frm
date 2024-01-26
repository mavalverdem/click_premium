VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTBanDet 
   Caption         =   "[Entidad]"
   ClientHeight    =   5325
   ClientLeft      =   2025
   ClientTop       =   1500
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5325
   ScaleMode       =   0  'User
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Pago a Proveedores"
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
      Height          =   615
      Left            =   3240
      TabIndex        =   56
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox cboTpoCta 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   4
         Left            =   2880
         Picture         =   "frmTBanDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   280
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   4
         Left            =   90
         TabIndex        =   57
         Top             =   240
         Width           =   405
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
         Left            =   480
         TabIndex        =   58
         Top             =   240
         Width           =   2385
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   1
      Left            =   7575
      Picture         =   "frmTBanDet.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   990
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   990
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoBan 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   3960
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   2940
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3960
      Width           =   675
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
      Index           =   10
      Left            =   4680
      TabIndex        =   30
      Top             =   3960
      Width           =   735
   End
   Begin VB.ComboBox cboTpoTCb 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3960
      Width           =   915
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   8
      Left            =   1080
      TabIndex        =   23
      Top             =   2985
      Width           =   6435
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   9
      Left            =   1080
      TabIndex        =   25
      Top             =   3315
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H80000002&
      Height          =   975
      Left            =   3405
      TabIndex        =   48
      Top             =   4305
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   550
         Width           =   1755
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
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   705
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
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   16
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   7
      Left            =   1095
      TabIndex        =   21
      Top             =   2655
      Width           =   6435
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
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   5955
      TabIndex        =   35
      Top             =   3630
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
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   5955
      TabIndex        =   36
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      ForeColor       =   &H00400000&
      Height          =   945
      Left            =   60
      TabIndex        =   11
      Top             =   1650
      Width           =   7875
      Begin VB.CheckBox chkPvsDoc 
         Caption         =   "Provisionar"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6480
         TabIndex        =   19
         Top             =   615
         Width           =   1230
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "P&endientes"
         Height          =   280
         Index           =   5
         Left            =   6480
         Picture         =   "frmTBanDet.frx":0354
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   3
         Left            =   6120
         Picture         =   "frmTBanDet.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   225
         Width           =   280
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   3
         Left            =   1320
         TabIndex        =   13
         Top             =   225
         Width           =   405
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   5
         Left            =   1320
         TabIndex        =   17
         Top             =   555
         Width           =   525
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   6
         Left            =   1830
         TabIndex        =   18
         Top             =   555
         Width           =   1155
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
         Left            =   1710
         TabIndex        =   14
         Top             =   225
         Width           =   4425
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
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1200
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
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   570
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   7575
      Picture         =   "frmTBanDet.frx":06A8
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   660
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   660
      Width           =   975
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   2
      Left            =   7575
      Picture         =   "frmTBanDet.frx":0852
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1320
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4560
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
         Picture         =   "frmTBanDet.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Picture         =   "frmTBanDet.frx":0BA6
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "frmTBanDet.frx":0D50
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Picture         =   "frmTBanDet.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Picture         =   "frmTBanDet.frx":0F9C
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Picture         =   "frmTBanDet.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComCtl2.DTPicker dtpFehOpe 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   63897601
      CurrentDate     =   37924.6695138889
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
      Index           =   1
      Left            =   2340
      TabIndex        =   7
      Top             =   990
      Width           =   5235
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
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   960
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   10
      Left            =   2940
      TabIndex        =   26
      Top             =   3645
      Width           =   840
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
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   11
      Left            =   4020
      TabIndex        =   28
      Top             =   3645
      Width           =   1410
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   975
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   3330
      Width           =   975
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Operación"
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
      Height          =   210
      Index           =   9
      Left            =   1080
      TabIndex        =   31
      Top             =   3645
      Width           =   750
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe"
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
      Height          =   210
      Index           =   12
      Left            =   6435
      TabIndex        =   32
      Top             =   3390
      Width           =   780
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   13
      Left            =   5535
      TabIndex        =   33
      Top             =   3645
      Width           =   360
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   14
      Left            =   5535
      TabIndex        =   34
      Top             =   3975
      Width           =   345
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   675
      Width           =   960
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
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   660
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
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   5895
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1335
      Width           =   975
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1515
   End
End
Attribute VB_Name = "frmTBanDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean
Private pnCta_IndCCo As Integer
Private pnCta_IndDoc As Integer
Private pnCta_IndAjD As Integer
Private pnCta_IndAnl As Integer
Private pnCta_IndBco As Integer

Private psCodCCo_Def As String
Private pnItemBan As Integer
Public pnUltIte, pnTpoMon As Integer
Public pnNroIte As Integer
Public psGlosa As String, psGlosax As String
Private sProcesoFyH As String

Private Sub CboTpoTCb_LostFocus()
Dim dcValIndDoc As Integer
  
With frmTBanGrd.uorstTGTCb
    If .RecordCount <> 0 Then
      .MoveFirst
      .Find "FehTCb = '" & frmTBanCab.dtpFehBan & "'"
      If .EOF Then
        MsgBox TEXT_9015, vbExclamation
        txtDato(10).Text = Format(0, FORMATO_NUM_2)
        txtDato(10).SetFocus
      Else
        txtDato(10).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
      End If
    Else
      txtDato(10).Text = Format(0, FORMATO_NUM_2)
    End If
End With
  
End Sub
Private Sub chkPvsDoc_Click()
  Dim sSentencia As String
  If pbNuevo Then
    If chkPvsDoc.Value = vbChecked Then
      ' Elimino los documentos del temporal
      sSentencia = "DELETE FROM codoctmp1 "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND usrcre='" & gsAbvUsr & "' "
      sSentencia = sSentencia & "AND fyhcre='" & sProcesoFyH & "'"
      frmTBanGrd.uocnnMain.Execute sSentencia
    End If
    cmdDatoAyud(5).Enabled = (chkPvsDoc.Value = vbUnchecked)
    txtImporte(0).Enabled = (chkPvsDoc.Value = vbChecked)
    txtImporte(1).Enabled = (chkPvsDoc.Value = vbChecked)
  End If
End Sub
Private Sub Form_Load()

  pbValidada = False
  
  Me.KeyPreview = True
  With frmTBanGrd                     'Cambiar Formulario de Grid.
    '[Datos.                           'Cambiar.
    With cboTpoBan
      .AddItem TPOBAN_ING_TXT, TPOBAN_ING
      .AddItem TPOBAN_EGR_TXT, TPOBAN_EGR
    End With
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
    End With
    With cboTpoTCb
      .AddItem TPOTCB_VTA_TXT, TPOTCB_VTA_IND
      .AddItem TPOTCB_CPR_TXT, TPOTCB_CPR_IND
    End With
    With cbotpocta
      .AddItem TPOCTA_COR_TXT_2, TPOCTA_COR_IND
      .AddItem TPOCTA_AHO_TXT_2, TPOCTA_AHO_IND
      .AddItem TPOCTA_MAE_TXT_2, TPOCTA_MAE_IND
      .AddItem TPOCTA_SIN_TXT_2, TPOCTA_SIN_IND
    End With
    
    If proceso = False Then
        txtDato(0).MaxLength = .uorstMain_1!CodCta.DefinedSize
        txtDato(1).MaxLength = .uorstMain_1!codaux.DefinedSize
        txtDato(2).MaxLength = .uorstMain_1!codcco.DefinedSize
        txtDato(3).MaxLength = .uorstMain_1!codtdc.DefinedSize
        
        txtDato(5).MaxLength = .uorstMain_1!serdoc.DefinedSize
        txtDato(6).MaxLength = .uorstMain_1!nrodoc.DefinedSize
        
        txtDato(gsIdioma + 6).MaxLength = .uorstMain_1!GloIte.DefinedSize
        txtDato(9 - gsIdioma).MaxLength = .uorstMain_1!GloItex.DefinedSize
        
        txtDato(9).MaxLength = .uorstMain_1!RefDoc.DefinedSize
        txtDato(10).MaxLength = 8
    Else
        txtDato(0).MaxLength = .uorstMain_1Fil!CodCta.DefinedSize
        txtDato(1).MaxLength = .uorstMain_1Fil!codaux.DefinedSize
        txtDato(2).MaxLength = .uorstMain_1Fil!codcco.DefinedSize
        txtDato(3).MaxLength = .uorstMain_1Fil!codtdc.DefinedSize
        
        txtDato(5).MaxLength = .uorstMain_1Fil!serdoc.DefinedSize
        txtDato(6).MaxLength = .uorstMain_1Fil!nrodoc.DefinedSize
        
        txtDato(gsIdioma + 6).MaxLength = .uorstMain_1Fil!GloIte.DefinedSize
        txtDato(9 - gsIdioma).MaxLength = .uorstMain_1Fil!GloItex.DefinedSize
        
        txtDato(9).MaxLength = .uorstMain_1Fil!RefDoc.DefinedSize
        txtDato(10).MaxLength = 8
    
    End If
    
    txtImporte(0).MaxLength = 14
    txtImporte(1).MaxLength = 14
    txtDeta(0).Text = Format(frmTBanCab.txtDeta(0).Text, FORMATO_NUM_1)
    txtDeta(1).Text = Format(frmTBanCab.txtDeta(1).Text, FORMATO_NUM_1)
    txtDeta(2).Text = Format(frmTBanCab.txtDeta(2).Text, FORMATO_NUM_1)
    txtDeta(3).Text = Format(frmTBanCab.txtDeta(3).Text, FORMATO_NUM_1)
    psGlosa = frmTBanCab.txtDato(0).Text
    psGlosax = frmTBanCab.txtDato(1).Text
    pnTpoMon = frmTBanCab.cboTpoMon.ListIndex
    
    If proceso = False Then
        txtDato(4).MaxLength = .uorstMain_1!codbco.DefinedSize
    Else
        txtDato(4).MaxLength = .uorstMain_1Fil!codbco.DefinedSize
    End If
    
    With dtpFehOpe
      .MinDate = DateAdd("m", -5, CDate("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct))
      .MaxDate = gfUltDia(CDate("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct))
    End With
    dtpFehOpe.Value = frmTBanCab.dtpFehBan.Value
    ']
  End With
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(17, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Fecha de Operación:", "Cuenta:", "Auxiliar :", "C.Costo:", "Tipo Documento:", "NºDocumento:", "Glosa:", "Traducción : ", "Referencias :", "Operación :", "Mon. Func.:", "Tipo de Cambio:", "Importe :", "M.N.:", "M.E.:", "Total M.N.:", "Total M.E.:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Operation Date:", "Account:", "Auxiliary :", "C.Center:", "Type Document:", "NºDocument:", "Gloss:", "Translation : ", "References :", "Operation :", "Func. Curr.:", "Rate of Exchange:", "Amount :", "N.C.:", "F.C.:", "Total N.C.:", "Total F.C.:")
  Next nElemento
  fraDocumento.Caption = Choose(gsIdioma, " Documento ", " Document ")
  cmdDatoAyud(5).Caption = Choose(gsIdioma, "P&endientes", "O&utstanding")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']
  
  If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
  End If
  cmdGrabar.Enabled = pbNuevo
  cmdDeshacer.Enabled = False
  upHabilitacion pbNuevo
  
  If pbNuevo Then
    txtDato(4).Text = frmTBanCab.txtDato(9).Text
    cbotpocta.ListIndex = 1
  End If
  
End Sub

Private Sub Form_Activate()
  '[Busca detalle de códigos.           'Cambiar (habilitar/deshabilitar).
  If Trim(txtDato(0).Text) <> "" And Trim(lblDatoDeta(0).Caption) <> "" Then
    ppAyuDet 0
    pnCta_IndDoc = frmTBanGrd.uorstCoCta!IndDoc
    pnCta_IndAjD = frmTBanGrd.uorstCoCta!IndAjD
    pnCta_IndAnl = frmTBanGrd.uorstCoCta!TpoAnl
    pnCta_IndCCo = frmTBanGrd.uorstCoCta!indcco
    psCodCCo_Def = IIf(IsNull(frmTBanGrd.uorstCoCta!codcco_def), "", frmTBanGrd.uorstCoCta!codcco_def)
    ' Actualiza los datos de centro de costo
    txtDato(2).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(0).Enabled)
    cmdDatoAyud(2).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(0).Enabled)
  End If
  If txtDato(1).Text <> "" Then ppAyuDet 1
  If txtDato(2).Text <> "" Then ppAyuDet 2
  If txtDato(3).Text <> "" Then ppAyuDet 3
  If txtDato(4).Text <> "" Then ppAyuDet 4
  ']
  
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   frmTFacGrd.uorstMain_1.CancelUpdate 'Cambiar Formulario de Grid.
If proceso = False Then
   If frmTBanGrd.uorstMain_1.RecordCount <> 0 Then
      frmTBanGrd.uorstMain_1.CancelUpdate
   End If
Else
   If frmTBanGrd.uorstMain_1Fil.RecordCount <> 0 Then
      frmTBanGrd.uorstMain_1Fil.CancelUpdate
   End If
End If
End Sub

Private Sub cmdRetroceder_Click()
If proceso = False Then
   gpTUe_Retroceder frmTBanGrd.uorstMain_1, Me
Else
   gpTUe_Retroceder frmTBanGrd.uorstMain_1Fil, Me
End If
End Sub

Private Sub cmdAvanzar_Click()
If proceso = False Then
   gpTUe_Avanzar frmTBanGrd.uorstMain_1, Me
Else
   gpTUe_Avanzar frmTBanGrd.uorstMain_1Fil, Me
End If
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
  '[Dato con el foco al corregir.       'Cambiar.
  cmdDatoAyud(5).Enabled = False
  txtDato(0).SetFocus
End Sub

Public Sub cmdGrabar_Click()
'   On Error GoTo Err

  '[No pertenece al Formulario - Agregado por Angel
  Dim dnNroIte As Integer
  Dim dnImpMN, dnImpME As Double
  Dim dcTpoMon, dcTpoCtb As String, sSqlexe As String
  Dim dvRegistro As Variant
  Dim nRegistro As Long
  
  ' Obtengo documentos seleccionados
  nRegistro = 0
  With frmTBanGrd.uorstCOTCbMes
    If .State = adStateOpen Then .Close
    .Source = "SELECT  COUNT(*) AS nDocuSele "
    .Source = .Source & "FROM codoctmp1 "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND usrcre='" & gsAbvUsr & "' "
    .Source = .Source & "AND fyhcre='" & sProcesoFyH & "' "
    .Source = .Source & "AND indsel='" & INDPREGEN_ACT & "'"
    .Open
    If .RecordCount <> 0 Then nRegistro = CLng(!ndocusele)
    .Close
  End With

  '[Validacion de Datos segun Indicadores de Cuenta.
  If nRegistro = 0 Then
    
    If xIndicador = 2 Then
        If cbotpocta.Text = "" Then MsgBox "Ingresar Tipo de Cuenta", vbExclamation: txtDato(0).SetFocus: Exit Sub
    End If
    
    If Len(Trim(txtDato(0).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(0).SetFocus: Exit Sub
    If pnCta_IndCCo = INDCCO_ACT And Len(Trim(txtDato(2).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(2).SetFocus: Exit Sub
    If pnCta_IndDoc = INDDOC_ACT And (Len(Trim(txtDato(3).Text)) = 0 Or Len(Trim(txtDato(5).Text)) = 0 Or Len(Trim(txtDato(6).Text)) = 0) Then MsgBox TEXT_6002, vbExclamation: txtDato(5).SetFocus: Exit Sub
    ' valida cta+auxiliar
    If pnCta_IndAjD = INDAJD_ACT And pnCta_IndAnl = TPOANL_AUX And pnCta_IndDoc = INDDOC_ACT And Len(Trim(txtDato(1).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(2).SetFocus: Exit Sub
    
    If cboTpoMon.ListIndex = TPOMON_NAC_IND And (CDec(txtImporte(0).Text) = 0) Then
      MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Nacional.", "You Must enter the amount in National Currency."), vbInformation
      txtImporte(0).SetFocus
      Exit Sub
    ElseIf cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtImporte(1).Text) = 0) Then
      MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Extranjera.", "You Must enter the amount in Foreign Currency."), vbInformation
      txtImporte(1).SetFocus
      Exit Sub
    End If
   
    If Len(Trim(txtDato(3).Text)) <> 0 And Len(Trim(txtDato(5).Text)) <> 0 And Len(Trim(txtDato(6).Text)) <> 0 Then
      With frmTBanGrd.uorstCOBanDet
        .Source = "SELECT codaux, codtdc, serdoc, nrodoc, impmn, impme, imptcb, tpotcb, tpomon, TpoPvs, TpoCtb, CodCta, CodDro, NroCpb, MesPvs "
        .Source = .Source & "FROM CoCpbDet "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND CodCta='" & txtDato(0).Text & "' "
        .Source = .Source & "AND CodAux='" & txtDato(1).Text & "' "
        .Source = .Source & "AND CodTDc='" & txtDato(3).Text & "' "
        .Source = .Source & "AND SerDoc='" & txtDato(5).Text & "' "
        .Source = .Source & "AND NroDoc='" & txtDato(6).Text & "' "
        .Source = .Source & "AND TpoPvs<>'" & TPOPVS_OTR & "' "
        .Source = .Source & "AND (coddro<>'" & frmTBanCab.txtLlave(0).Text & "' "
        .Source = .Source & "AND nrocpb<>'" & frmTBanCab.txtLlave(1).Text & "') "
        .Source = .Source & "UNION "
        .Source = .Source & "SELECT codaux, codtdc, serdoc, nrodoc, impmn, impme, imptcb, tpotcb, tpomon, "
        .Source = .Source & "(CASE pvsdoc WHEN " & INDPREGEN_ACT & " THEN '" & TPOPVS_PVS & "' ELSE '" & TPOPVS_CAN & "' END) AS tpopvs, "
        .Source = .Source & "(CASE tpoban WHEN " & TPOBAN_EGR & " THEN '" & TPOCTB_DEB & "' ELSE '" & TPOCTB_HAB & "' END) AS tpoctb, "
        .Source = .Source & "codcta, coddro, nroban AS nrocpb, mespvs "
        .Source = .Source & "FROM cobandet "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND codcta='" & txtDato(0).Text & "' "
        .Source = .Source & "AND codaux ='" & txtDato(1).Text & "' "
        .Source = .Source & "AND codtdc='" & txtDato(3).Text & "' "
        .Source = .Source & "AND serdoc='" & txtDato(5).Text & "' "
        .Source = .Source & "AND nrodoc='" & txtDato(6).Text & "' "
        .Source = .Source & "AND coddro='" & frmTBanCab.txtLlave(0).Text & "' "
        .Source = .Source & "AND nroban='" & frmTBanCab.txtLlave(1).Text & "'"
        .Open
        dnImpMN = 0: dnImpME = 0
        ' Valido la provision
        If chkPvsDoc.Value = vbChecked Then
          frmTBanGrd.uorstCOBanDet.Find "TpoPvs='" & TPOPVS_PVS & "'"
          If Not frmTBanGrd.uorstCOBanDet.EOF Then
            If (frmTBanGrd.uorstCOBanDet!coddro <> frmTBanCab.txtLlave(0).Text Or frmTBanGrd.uorstCOBanDet!NroCpb <> frmTBanCab.txtLlave(1).Text) Then
              MsgBox Choose(gsIdioma, "Ya está registrada la provision del documento.", "the provision of document is registered ."), vbExclamation
              frmTBanGrd.uorstCOBanDet.Close
              txtDato(5).SetFocus
              Exit Sub
            End If
          End If
        Else
          ' Valido la cancelacion
          If pbNuevo And chkPvsDoc.Value = INDPREGEN_INA Then
            frmTBanGrd.uorstCOBanDet.Find "tpopvs='" & TPOPVS_PVS & "'"
            If .EOF Then
              MsgBox Choose(gsIdioma, "No está registrada la provisión del documento.", "the provision of document is not registered."), vbExclamation
              frmTBanGrd.uorstCOBanDet.Close
              txtDato(3).SetFocus
              Exit Sub
            Else
              If frmTBanGrd.uorstCOBanDet!CodCta <> txtDato(0).Text Then
                MsgBox Choose(gsIdioma, "La cuenta de la cancelación no es igual a la de la provisión.", "The cancelation account is not the same of the provision."), vbExclamation
                frmTBanGrd.uorstCOBanDet.Close
                txtDato(0).SetFocus
                Exit Sub
              End If
              If frmTBanGrd.uorstCOBanDet!TpoCtb = TPOCTB_DEB And (CDec(txtImporte(0).Text) > 0 Or CDec(txtImporte(1).Text) > 0) Then
                MsgBox Choose(gsIdioma, "Revise la información. La provisión está registrada en el DEBE.", "You review information. The provision is registered in DEBIT."), vbExclamation
                frmTBanGrd.uorstCOBanDet.Close
                txtImporte(1).SetFocus
                Exit Sub
              End If
              If frmTBanGrd.uorstCOBanDet!TpoCtb = TPOCTB_HAB And (CDec(txtImporte(0).Text) > 0 Or CDec(txtImporte(1).Text) > 0) Then
                MsgBox Choose(gsIdioma, "Revise la información. La provisión está registrada en el HABER.", "You review information. The provision is registered in CREDIT."), vbExclamation
                frmTBanGrd.uorstCOBanDet.Close
                txtImporte(0).SetFocus
                Exit Sub
              End If
            End If
          End If
              
          If Not .EOF Then
            dcTpoMon = frmTBanGrd.uorstCOBanDet!tpomon
            frmTBanGrd.uorstCOBanDet.MoveFirst
            Do
              If ((frmTBanGrd.uorstCOBanDet!coddro & frmTBanGrd.uorstCOBanDet!NroCpb & frmTBanGrd.uorstCOBanDet!mespvs) <> (frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1).Text & gsMesAct)) Then
                dnImpMN = dnImpMN + IIf(frmTBanGrd.uorstCOBanDet!TpoPvs = TPOPVS_PVS, frmTBanGrd.uorstCOBanDet!ImpMN, frmTBanGrd.uorstCOBanDet!ImpMN * (-1))
                dnImpME = dnImpME + IIf(frmTBanGrd.uorstCOBanDet!TpoPvs = TPOPVS_PVS, frmTBanGrd.uorstCOBanDet!ImpME, frmTBanGrd.uorstCOBanDet!ImpME * (-1))
              End If
              frmTBanGrd.uorstCOBanDet.MoveNext
            Loop Until .EOF
            If dcTpoMon = TPOMON_NAC Then
              If CDec(txtImporte(0).Text) > 0 Then
                If dnImpMN < CDec(txtImporte(0).Text) Then
                  MsgBox Choose(gsIdioma, "El monto de la cancelación es mayor al de la provisión.", "The cancelation amount is more  than provision."), vbExclamation
                  frmTBanGrd.uorstCOBanDet.Close
                  txtImporte(0).SetFocus
                  Exit Sub
                End If
              End If
            Else
              If CDec(txtImporte(1).Text) > 0 Then
                If dnImpME < CDec(txtImporte(1).Text) Then
                  MsgBox Choose(gsIdioma, "El monto de la cancelación es mayor al de la provisión.", "The cancelation amount is more  than provision."), vbExclamation
                  frmTBanGrd.uorstCOBanDet.Close
                  txtImporte(1).SetFocus
                  Exit Sub
                End If
              End If
            End If
          End If
        End If
        frmTBanGrd.uorstCOBanDet.Close
      End With
    End If
  End If
  With frmTBanGrd                     'Cambiar Formulario de Grid.
    .uocnnMain.BeginTrans            'INICIA TRANSACCION.
    If nRegistro >= 1 And pbNuevo Then
      upDatosDetalle
    Else
    
    If proceso = False Then
      If pbNuevo Then
        .uorstMain_1.AddNew
      Else
        '' corregido error 11/09/2009
        '.uorstMain_1.Find "cLlave='" & frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1) & Trim(Str(pnUltIte)) & "'"
        .uorstMain_1.Find "NroItem='" & Trim(Str(pnUltIte)) & "'"
      End If
    Else
      If pbNuevo Then
        .uorstMain_1Fil.AddNew
      Else
        .uorstMain_1Fil.Find "cLlave='" & frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1) & Trim(Str(pnUltIte)) & "'"
      End If
    End If
      
      upDatosDesconectados 0
      ' Actualizo la cabecera
      If proceso = False Then
        .uorstMain_0.Update
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
      Else
              .uorstMain_0Fil.Update
        With .uorstMain_1Fil
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
    End If
    ' Elimino los documentos del temporal
    sSqlexe = "DELETE FROM codoctmp1 "
    sSqlexe = sSqlexe & "WHERE codemp='" & gsCodEmp & "' "
    sSqlexe = sSqlexe & "AND pdoano='" & gsAnoAct & "' "
    sSqlexe = sSqlexe & "AND usrcre='" & gsAbvUsr & "' "
    sSqlexe = sSqlexe & "AND fyhcre='" & sProcesoFyH & "'"
    .uocnnMain.Execute sSqlexe
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      
    If proceso = False Then
    If pbNuevo Then
      dnNroIte = .uorstMain_1!nroitem
      .uorstMain_1.Requery
      frmTBanCab.upDatosGrid
      If .uorstMain_1.RecordCount <> 0 Then
        '[Búsqueda de llave actual.     'Cambiar.
        .uorstMain_1.MoveFirst
        'Modificado 11/09/2009
        '.uorstMain_1.Find "cLlave='" & frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
        .uorstMain_1.Find "NroItem='" & Trim(Str(dnNroIte)) & "'"
      End If
      upDatosPredeterminados
      '[Dato con el foco al añadir.   'Cambiar.
      txtDato(0).SetFocus
    Else
      If .uorstMain_1.RecordCount <> 0 Then
        dnNroIte = .uorstMain_1!nroitem
        '[Búsqueda de llave actual.     'Cambiar.
        .uorstMain_1.MoveFirst
        'Modificado 11/09/2009
        '.uorstMain_1.Find "cLlave='" & frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
        .uorstMain_1.Find "NroItem='" & Trim(Str(dnNroIte)) & "'"
        
        If .uorstMain_1.EOF Then .uorstMain_1.MoveFirst
      End If
      cmdRetroceder.Enabled = True
      cmdAvanzar.Enabled = True
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      upHabilitacion False
    End If
    Else
    
    If pbNuevo Then
      dnNroIte = .uorstMain_1Fil!nroitem
      .uorstMain_1Fil.Requery
      frmTBanCab.upDatosGrid
      If .uorstMain_1Fil.RecordCount <> 0 Then
        '[Búsqueda de llave actual.     'Cambiar.
        .uorstMain_1Fil.MoveFirst
        .uorstMain_1Fil.Find "cLlave='" & frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
      End If
      upDatosPredeterminados
      '[Dato con el foco al añadir.   'Cambiar.
      txtDato(0).SetFocus
    Else
      If .uorstMain_1Fil.RecordCount <> 0 Then
        dnNroIte = .uorstMain_1Fil!nroitem
        '[Búsqueda de llave actual.     'Cambiar.
        .uorstMain_1Fil.MoveFirst
        .uorstMain_1Fil.Find "cLlave='" & frmTBanCab.txtLlave(0).Text & frmTBanCab.txtLlave(1) & Trim(Str(dnNroIte)) & "'"
        If .uorstMain_1Fil.EOF Then .uorstMain_1Fil.MoveFirst
      End If
      cmdRetroceder.Enabled = True
      cmdAvanzar.Enabled = True
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      upHabilitacion False
    End If
    
    End If
    
    ' Inicializo el numero de item
    pnItemBan = 0
  End With

'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
fEstMayUpd
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion


  Exit Sub
Err:
  gpErrores
  ' Inicializo el numero de item
  pnItemBan = 0
  frmTBanGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
  
End Sub

Public Sub cmdDeshacer_Click()
    '[Propio del formulario.
   With frmTBanCab
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
   Case 0, 1, 2, 3, 4
    If (pnCta_IndCCo = INDCCO_ACT And Index = 1) Or Index <> 1 Then
      txtDato(Index).SetFocus
    End If
   Case 5  ' Inserto los documentos agrupados a la tabla tempolral
    txtDato(Index).SetFocus
   End Select
  If (pnCta_IndCCo = INDCCO_ACT And Index = 1) Or Index <> 1 Then ppAyuBus Index
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
    ppAyuBus Index
  End If
End Sub
Private Sub txtDato_LostFocus(Index As Integer)
  If Index = 10 Then
    txtDato(Index).Text = IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)
    txtDato(Index).Text = Format(CDec(txtDato(Index).Text), FORMATO_NUM_2)
  End If
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)

  On Error GoTo Err
  
  'Completa con ceros a la izquierda.
  Select Case Index
   Case 3, 5, 6                            'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength And IsNumeric(txtDato(Index).Text) Then
      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
    End If
  End Select
  
  'Busca el dato en su tabla principal.
  Select Case Index
   Case 0                              'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    If lblDatoDeta(Index).Caption <> "" Then
      
      pnCta_IndCCo = frmTBanGrd.uorstCoCta!indcco
      pnCta_IndDoc = frmTBanGrd.uorstCoCta!IndDoc
      pnCta_IndAjD = frmTBanGrd.uorstCoCta!IndAjD
      pnCta_IndAnl = frmTBanGrd.uorstCoCta!TpoAnl
      psCodCCo_Def = IIf(IsNull(frmTBanGrd.uorstCoCta!codcco_def), "", frmTBanGrd.uorstCoCta!codcco_def)
      
      ' Actualizo los datos adicionales
      txtDato(2).Text = IIf(txtDato(2).Text = "", psCodCCo_Def, txtDato(2).Text)
      txtDato(2).Text = IIf(pnCta_IndCCo = INDCCO_ACT, txtDato(2).Text, "")
      lblDatoDeta(2).Caption = IIf(pnCta_IndCCo = INDCCO_ACT, lblDatoDeta(2).Caption, "")
      txtDato(2).Enabled = (pnCta_IndCCo = INDCCO_ACT)
      cmdDatoAyud(2).Enabled = (pnCta_IndCCo = INDCCO_ACT)
    End If
   Case 1, 2, 3, 4
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
   Case 10
    txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_2)
  End Select
  ']
  '[ Propio del formulario
  If Index = 0 Or Index = 10 Then
    If Val(txtDato(10).Text) = 0 Then
      txtDato(10).Tag = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
      With frmTBanGrd.uorstTGTCb
        If .RecordCount <> 0 Then
          .MoveFirst
          .Find "FehTCb = '" & frmTBanCab.dtpFehBan & "'"
          ' [Adicional Agregado por Angel
          If .EOF Then
            MsgBox TEXT_9015, vbExclamation
            txtDato(10).Text = Format(0, FORMATO_NUM_2)
            txtDato(0).SetFocus
          Else
            txtDato(10).Text = Format(IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
          End If
          ']
        Else
          txtDato(10).Text = Format(0, FORMATO_NUM_2)
        End If
      End With
    End If
  End If
  
  Exit Sub
Err:
  gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  
  Dim nNumeroRecord As Long
  
  Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
    modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1                              'Cambiar (añadir índices).
    modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2                              'Cambiar (añadir índices).
    modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 3                              'Cambiar (añadir índices).
    modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 5                              'Cambiar (añadir índices).
    ' Primer paso - Elimino documentos del temporal y provisión
    cmdDatoAyud(tnIndex).Tag = "WHERE codemp='" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND pdoano='" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND usrcre='" & gsAbvUsr & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND fyhcre='" & sProcesoFyH & "'"
    frmTBanGrd.uocnnMain.Execute "DELETE FROM codoctmp1 " & cmdDatoAyud(tnIndex).Tag, nNumeroRecord
    sProcesoFyH = Format(Now, s_FmtFeHoMysql_0)
    
    ' Segundo paso - Genero temporal de acumulado documentos
    frmTBanGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocupen", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 11)='#tmpdocupen') DROP TABLE #tmpdocupen")
    
    cmdDatoAyud(tnIndex).Tag = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdocupen ", "")
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT det.codemp, det.pdoano, det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END)), 2) AS DebeMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END)), 2) AS HaberMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END)), 2) AS DebeME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END)), 2) AS HaberME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END) - (CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END))), 2) AS SaldoMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END) - (CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END))), 2) AS SaldoME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END) - (CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END))), 2) AS CanceMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END) - (CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END))), 2) AS CanceME "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpdocupen ")
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM cocpbdet det "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "LEFT JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.codcta=cta.codcta "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE det.codemp = '" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.pdoano = '" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.tposdo = '" & TPOSDO_INV & "' "
    If txtDato(0).Text <> "" Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.codcta = '" & txtDato(0).Text & "' "
    Else
      If frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 1) >= '1' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 2) <= '31' "
      Else
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 2) >= '33' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 2) <= '49' "
      End If
    End If
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.codaux = '" & txtDato(1).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.codtdc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.serdoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.nrodoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.mespvs <= '" & gsMesAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.inddoc = '" & INDDOC_ACT & "' "
    If ps_Plataforma = pSrvMySql Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND (det.feedoc) <= '" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.feedoc <= CONVERT(smalldatetime, '" & Format(dtpFehOpe.Value, "dd/mm/yyyy") & "', 103) "
    End If
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND NOT (det.mespvs='" & gsMesAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.coddro='" & frmTBanCab.txtLlave(0).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.nrocpb='" & frmTBanCab.txtLlave(1).Text & "') "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "GROUP BY det.codemp, det.pdoano, det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "UNION "
    ' Documentos de comprobante de diario activo
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT ban.codemp, ban.pdoano, ban.codcta, ban.codaux, ban.codtdc, ban.serdoc, ban.nrodoc, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE ban.tpoban WHEN " & TPOBAN_EGR & " THEN ban.impmn ELSE 0 END)), 2) AS DebeMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE ban.tpoban WHEN " & TPOBAN_ING & " THEN ban.impmn ELSE 0 END)), 2) AS HaberMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE ban.tpoban WHEN " & TPOBAN_EGR & " THEN ban.impme ELSE 0 END)), 2) AS DebeME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM((CASE ban.tpoban WHEN " & TPOBAN_ING & " THEN ban.impme ELSE 0 END)), 2) AS HaberME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE ban.tpoban WHEN " & TPOBAN_EGR & " THEN ban.impmn ELSE 0 END) - (CASE ban.tpoban WHEN " & TPOBAN_ING & " THEN ban.impmn ELSE 0 END))), 2) AS SaldoMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE ban.tpoban WHEN " & TPOBAN_EGR & " THEN ban.impme ELSE 0 END) - (CASE ban.tpoban WHEN " & TPOBAN_ING & " THEN ban.impme ELSE 0 END))), 2) AS SaldoME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE ban.tpoban WHEN " & TPOBAN_EGR & " THEN ban.impmn ELSE 0 END) - (CASE ban.tpoban WHEN " & TPOBAN_ING & " THEN ban.impmn ELSE 0 END))), 2) AS CanceMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(((CASE ban.tpoban WHEN " & TPOBAN_EGR & " THEN ban.impme ELSE 0 END) - (CASE ban.tpoban WHEN " & TPOBAN_ING & " THEN ban.impme ELSE 0 END))), 2) AS CanceME "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM cobandet ban "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "LEFT JOIN cocta cta ON ban.codemp=cta.codemp AND ban.pdoano=cta.pdoano AND ban.codcta=cta.codcta "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE ban.codemp = '" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.pdoano = '" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.tposdo = '" & TPOSDO_INV & "' "
    If txtDato(0).Text <> "" Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.codcta = '" & txtDato(0).Text & "' "
    Else
      If frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 1) >= '1' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 2) <= '31' "
      Else
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 2) >= '33' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 2) <= '49' "
      End If
    End If
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.codaux = '" & txtDato(1).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ban.codtdc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ban.serdoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ban.nrodoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.mespvs ='" & gsMesAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.inddoc ='" & INDDOC_ACT & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.coddro ='" & frmTBanCab.txtLlave(0).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.nroban ='" & frmTBanCab.txtLlave(1).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "GROUP BY ban.codemp, ban.pdoano, ban.codcta, ban.codaux, ban.codtdc, ban.serdoc, ban.nrodoc "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ORDER BY codcta, codaux, codtdc, serdoc, nrodoc"
    frmTBanGrd.uocnnMain.Execute cmdDatoAyud(tnIndex).Tag, nNumeroRecord
    
    ' Elimino los datos de documento y provisión tabla temporal al inicio para los casos que se cuelga y no se borra
    cmdDatoAyud(tnIndex).Tag = "WHERE codemp='" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND pdoano='" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND usrcre='" & gsAbvUsr & "' "
    frmTBanGrd.uocnnMain.Execute "DELETE FROM codoctmp1 " & cmdDatoAyud(tnIndex).Tag

    ' Tercer paso - Inserto los documentos  pendientes
    cmdDatoAyud(tnIndex).Tag = "INSERT INTO codoctmp1 (codemp, pdoano, codcta, codaux, codtdc, serdoc, nrodoc, impdmn, imphmn, impdme, imphme, impsmn, impsme, tpomon, codcco, indsel, imppmn, imppme, usrcre, fyhcre) "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT det.codemp, det.pdoano, det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(DebeMN), 2) AS DebeMN, ROUND(SUM(HaberMN), 2) AS HaberMN, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(DebeME), 2) AS DebeME, ROUND(SUM(HaberME), 2) AS HaberME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(DebeMN-HaberMN), 2) AS SaldoMN, ROUND(SUM(DebeME-HaberME), 2) AS SaldoME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "cta.tpomon, cta.codcco_def, '" & INDPREGEN_INA & "' AS cIndSel, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ROUND(SUM(DebeMN-HaberMN), 2) AS CanceMN, ROUND(SUM(DebeME-HaberME), 2) AS CanceME, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "'" & gsAbvUsr & "' AS usrcre, '" & sProcesoFyH & "' AS fyhcre "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM " & ps_Prefijo & "tmpdocupen det "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "LEFT JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.codcta=cta.codcta "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE det.codemp = '" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.pdoano = '" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "GROUP BY det.codemp, det.pdoano, det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ORDER BY codcta, codaux, codtdc, serdoc, nrodoc"
    frmTBanGrd.uocnnMain.Execute cmdDatoAyud(tnIndex).Tag
    
    ' Inserto provisión de documentos
    cmdDatoAyud(tnIndex).Tag = "INSERT INTO codoctmp2 (codemp, pdoano, mespvs, codcta, codaux, codtdc, serdoc, nrodoc, tpomon, feedoc, usrcre, fyhcre) "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT DISTINCT det.codemp, det.pdoano, det.mespvs, det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc, det.tpomon, det.feedoc, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "'" & gsAbvUsr & "' AS usrcre, '" & sProcesoFyH & "' AS fyhcre "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM cocpbdet det "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "LEFT JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.codcta=cta.codcta "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE det.codemp = '" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.pdoano = '" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.tposdo = '" & TPOSDO_INV & "' "
    If txtDato(0).Text <> "" Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.codcta = '" & txtDato(0).Text & "' "
    Else
      If frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 1) >= '1' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 2) <= '31' "
      Else
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 2) >= '33' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(det.codcta, 2) <= '49' "
      End If
    End If
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.codaux = '" & txtDato(1).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.codtdc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.serdoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.nrodoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.mespvs <= '" & gsMesAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.tpopvs='" & TPOPVS_PVS & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.inddoc = '" & INDDOC_ACT & "' "
    If ps_Plataforma = pSrvMySql Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND (det.feedoc) <= '" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.feedoc <= CONVERT(smalldatetime, '" & Format(dtpFehOpe.Value, "dd/mm/yyyy") & "', 103) "
    End If
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND NOT (det.mespvs='" & gsMesAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.coddro='" & frmTBanCab.txtLlave(0).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND det.nrocpb='" & frmTBanCab.txtLlave(1).Text & "') "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "UNION "
    ' Documentos de comprobante de diario activo
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "SELECT DISTINCT ban.codemp, ban.pdoano, ban.mespvs, ban.codcta, ban.codaux, ban.codtdc, ban.serdoc, ban.nrodoc, ban.tpomon, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "'" & Format(dtpFehOpe.Value, "yyyy-mm-dd") & "' AS feedoc, "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "'" & gsAbvUsr & "' AS usrcre, '" & sProcesoFyH & "' AS fyhcre "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "FROM cobandet ban "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "LEFT JOIN cocta cta ON ban.codemp=cta.codemp AND ban.pdoano=cta.pdoano AND ban.codcta=cta.codcta "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "WHERE ban.codemp = '" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.pdoano = '" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.tposdo = '" & TPOSDO_INV & "' "
    If txtDato(0).Text <> "" Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.codcta = '" & txtDato(0).Text & "' "
    Else
      If frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 1) >= '1' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 2) <= '31' "
      Else
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 2) >= '33' "
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND LEFT(ban.codcta, 2) <= '49' "
      End If
    End If
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.codaux = '" & txtDato(1).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ban.codtdc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ban.serdoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ban.nrodoc, '') <> '' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.mespvs ='" & gsMesAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.pvsdoc='" & INDDOC_ACT & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND cta.inddoc ='" & INDDOC_ACT & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.coddro ='" & frmTBanCab.txtLlave(0).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND ban.nroban ='" & frmTBanCab.txtLlave(1).Text & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "ORDER BY codcta, codaux, codtdc, serdoc, nrodoc"
    frmTBanGrd.uocnnMain.Execute cmdDatoAyud(tnIndex).Tag, nNumeroRecord
    
    ' Elimino temporal de acumulado documentos
    frmTBanGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocupen", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 11)='#tmpdocupen') DROP TABLE #tmpdocupen")
    
    ' Filtro de seleccion
    cmdDatoAyud(tnIndex).Tag = "codoctmp1.usrcre='" & gsAbvUsr & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND codoctmp1.fyhcre='" & sProcesoFyH & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND (CASE b.tpomon WHEN '" & TPOMON_NAC & "' THEN codoctmp1.ImpSMN ELSE codoctmp1.ImpSME END) <> 0 "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND b.MesPvs <= '" & gsMesAct & "' "
    If ps_Plataforma = pSrvMySql Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND b.FeEDoc <= ('" & Format(dtpFehOpe.Value, "yyyy/mm/dd") & "')"
    ElseIf ps_Plataforma = pSrvSql Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND b.FeEDoc <= CONVERT(smalldatetime, '" & Format(dtpFehOpe.Value, "dd/mm/yyyy") & "', 103)"
    End If
    modAyuBus.Sel_Doc cmdDatoAyud(tnIndex).Tag, txtDato(3).Text & txtDato(4).Text & txtDato(5).Text, 3690, 7080, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    ' Elimino los datos de la tabla temporal
    cmdDatoAyud(tnIndex).Tag = "WHERE codemp='" & gsCodEmp & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND pdoano='" & gsAnoAct & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND usrcre='" & gsAbvUsr & "' "
    cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND fyhcre='" & sProcesoFyH & "' "
    ' Si no acepto la seleccion elimino todo y provisión
    If frmOSelPen.uaWhere = INDPREGEN_ACT Then
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND indsel='" & INDPREGEN_INA & "'"
    End If
    frmTBanGrd.uocnnMain.Execute "DELETE FROM codoctmp1 " & cmdDatoAyud(tnIndex).Tag
    
    txtDato(3).Text = Left(frmOSelPen.uvDato1, 2)
    txtDato(5).Text = Mid(frmOSelPen.uvDato1, 3, pLenSerDoc)
    txtDato(6).Text = Mid(frmOSelPen.uvDato1, 3 + pLenSerDoc)
    ' Obtengo los datos por default del documento
    With frmTBanGrd.uorstCOTCbMes
      If .State = adStateOpen Then .Close
      .Source = "SELECT ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(imppmn, 0)), 2) AS imppmn, "
      .Source = .Source & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(imppme, 0)), 2) AS imppme "
      .Source = .Source & "FROM codoctmp1 "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND usrcre='" & gsAbvUsr & "' "
      .Source = .Source & "AND fyhcre='" & sProcesoFyH & "' "
      .Source = .Source & "AND indsel='" & INDPREGEN_ACT & "'"
      .Open
      If .RecordCount <> 0 Then
        txtImporte(0).Text = CDec(IIf(IsNull(!imppmn), 0, !imppmn))
        txtImporte(1).Text = CDec(IIf(IsNull(!imppme), 0, !imppme))
      End If
      .Close
    End With
    Case 4
    
    modAyuBus.Bco_Cod "", txtDato(4).Text, 0, 0, Me.Top + txtDato(4).Top + txtDato(4).Height, Me.Left + txtDato(4).Left
    txtDato(4).Text = frmOAyuBus.uvDato1
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
    With frmTBanGrd.uorstCoCta
      .MoveFirst
      .Find "CodCta='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & frmTBanGrd.uorstCoCta!detcta
      End If
    End With
   Case 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTBanGrd.uorstTGAux
         .MoveFirst
         .Find "codaux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTBanGrd.uorstTGAux!razAux
         End If
      End With
   Case 2
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTBanGrd.uorstCoCCo
         .MoveFirst
         .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTBanGrd.uorstCoCCo!detcco
         End If
      End With
   Case 3
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmTBanGrd.uorstTGTDc
      .MoveFirst
      .Find "CodTDc='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & frmTBanGrd.uorstTGTDc!dettdc
      End If
    End With
   Case 4
    If txtDato(4).Text = "" Then
      lblDatoDeta(4).Caption = ""
      Exit Function
    End If
    With frmTBanGrd.uorstCoBco
      .MoveFirst
      .Find "Codbco='" & txtDato(4).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        lblDatoDeta(4).Caption = ""
        txtDato(4).Text = ""
        ppAyuDet = True
      Else
        lblDatoDeta(4).Caption = " " & frmTBanGrd.uorstCoBco!detbco
      End If
    End With
    
End Select
End Function

Private Sub upDatosDetalle()
  Dim dnContador As Integer, nTransaccion As Integer
  Dim sMoneda As String, s_Moneda As String, sCodRegistro As String
  Dim nImporte As Double
  Dim nImporteMN As Double, nImporteME As Double
  
  On Error GoTo Err
  
  frmTBanGrd.uorstCOTCbMes.Source = "SELECT codcta, codaux, codtdc, serdoc, nrodoc, codcco, tpomon, imppmn, imppme "
  frmTBanGrd.uorstCOTCbMes.Source = frmTBanGrd.uorstCOTCbMes.Source & "FROM codoctmp1 "
  frmTBanGrd.uorstCOTCbMes.Source = frmTBanGrd.uorstCOTCbMes.Source & "WHERE codemp='" & gsCodEmp & "' "
  frmTBanGrd.uorstCOTCbMes.Source = frmTBanGrd.uorstCOTCbMes.Source & "AND pdoano='" & gsAnoAct & "' "
  frmTBanGrd.uorstCOTCbMes.Source = frmTBanGrd.uorstCOTCbMes.Source & "AND usrcre='" & gsAbvUsr & "' "
  frmTBanGrd.uorstCOTCbMes.Source = frmTBanGrd.uorstCOTCbMes.Source & "AND fyhcre='" & sProcesoFyH & "' "
  frmTBanGrd.uorstCOTCbMes.Source = frmTBanGrd.uorstCOTCbMes.Source & "AND indsel='" & INDPREGEN_ACT & "'"
  frmTBanGrd.uorstCOTCbMes.Open
  
  If frmTBanGrd.uorstCOTCbMes.RecordCount <> 0 Then
    pnItemBan = frmTBanGrd.pfNumItemBan(gsAnoAct, gsMesAct, frmTBanCab.txtLlave(0).Text, frmTBanCab.txtLlave(1).Text)
    pnNroIte = pnItemBan
    psGlosa = txtDato(6).Text
    psGlosax = txtDato(7).Text
    sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
    If proceso = False Then
      While Not frmTBanGrd.uorstCOTCbMes.EOF
        frmTBanGrd.uorstMain_1.AddNew
        frmTBanGrd.uorstMain_1!codemp = gsCodEmp
        frmTBanGrd.uorstMain_1!pdoano = gsAnoAct
        frmTBanGrd.uorstMain_1!mespvs = gsMesAct
        frmTBanGrd.uorstMain_1!coddro = frmTBanCab.txtLlave(0).Text
        frmTBanGrd.uorstMain_1!nroban = frmTBanCab.txtLlave(1).Text
        frmTBanGrd.uorstMain_1!nroitem = pnNroIte
        frmTBanGrd.uorstMain_1!CodCta = frmTBanGrd.uorstCOTCbMes!CodCta
        sCodRegistro = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!codaux), "", frmTBanGrd.uorstCOTCbMes!codaux)
        frmTBanGrd.uorstMain_1!codaux = IIf(sCodRegistro = "", Null, sCodRegistro)
        sCodRegistro = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!codcco), "", frmTBanGrd.uorstCOTCbMes!codcco)
        frmTBanGrd.uorstMain_1!codcco = IIf(sCodRegistro = "", Null, sCodRegistro)
        sCodRegistro = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!codtdc), "", frmTBanGrd.uorstCOTCbMes!codtdc)
        frmTBanGrd.uorstMain_1!codtdc = IIf(sCodRegistro = "", Null, sCodRegistro)
        frmTBanGrd.uorstMain_1!serdoc = frmTBanGrd.uorstCOTCbMes!serdoc
        frmTBanGrd.uorstMain_1!nrodoc = frmTBanGrd.uorstCOTCbMes!nrodoc
        frmTBanGrd.uorstMain_1!GloIte = IIf(txtDato(gsIdioma + 6).Text = "", Null, txtDato(gsIdioma + 6).Text)
        frmTBanGrd.uorstMain_1!GloItex = IIf(txtDato(9 - gsIdioma).Text = "", Null, txtDato(9 - gsIdioma).Text)
        frmTBanGrd.uorstMain_1!RefDoc = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
        s_Moneda = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!tpomon), sMoneda, frmTBanGrd.uorstCOTCbMes!tpomon)
        nImporte = CDec(IIf(s_Moneda = TPOMON_NAC, frmTBanGrd.uorstCOTCbMes!imppmn, frmTBanGrd.uorstCOTCbMes!imppme))
        nTransaccion = IIf(frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING, IIf(nImporte < 0, TPOBAN_EGR, TPOBAN_ING), IIf(nImporte > 0, TPOBAN_ING, TPOBAN_EGR))
        frmTBanGrd.uorstMain_1!tpoban = nTransaccion
        frmTBanGrd.uorstMain_1!tpomon = s_Moneda
        frmTBanGrd.uorstMain_1!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
        frmTBanGrd.uorstMain_1!ImpTCb = CDec(txtDato(10).Text)
        frmTBanGrd.uorstMain_1!ImpMN = CDec(Abs(frmTBanGrd.uorstCOTCbMes!imppmn))
        frmTBanGrd.uorstMain_1!ImpME = CDec(Abs(frmTBanGrd.uorstCOTCbMes!imppme))
        frmTBanGrd.uorstMain_1!UsrCre = gsAbvUsr
        frmTBanGrd.uorstMain_1!FyHCre = Now
        
        frmTBanGrd.uorstMain_1!TpoCTA = Choose(cbotpocta.ListIndex + 1, TPOCTA_AHO_IND, TPOCTA_COR_IND, TPOCTA_MAE_IND, TPOCTA_SIN_IND)
        frmTBanGrd.uorstMain_1!codbco = txtDato(4).Text
        
        frmTBanGrd.uorstMain_1.Update
        
        ' actualizo los datos de la cabecera de bancos
        nImporteMN = CDec(frmTBanGrd.uorstCOTCbMes!imppmn)
        nImporteME = CDec(frmTBanGrd.uorstCOTCbMes!imppme)
        If frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING Then
          nImporteMN = nImporteMN * Choose(nTransaccion + 1, 1, -1)
          nImporteME = nImporteME * Choose(nTransaccion + 1, 1, -1)
        Else
          nImporteMN = nImporteMN * Choose(nTransaccion + 1, -1, 1)
          nImporteME = nImporteME * Choose(nTransaccion + 1, -1, 1)
        End If
      
        frmTBanGrd.uorstMain_0!ImpMN = Round(frmTBanGrd.uorstMain_0!ImpMN + nImporteMN, 2)
        frmTBanGrd.uorstMain_0!ImpME = Round(frmTBanGrd.uorstMain_0!ImpME + nImporteME, 2)
        frmTBanGrd.uorstMain_0.Update
        
        pnNroIte = pnNroIte + 1
        frmTBanGrd.uorstCOTCbMes.MoveNext
      Wend
    Else
    
      While Not frmTBanGrd.uorstCOTCbMes.EOF
        frmTBanGrd.uorstMain_1Fil.AddNew
        frmTBanGrd.uorstMain_1Fil!codemp = gsCodEmp
        frmTBanGrd.uorstMain_1Fil!pdoano = gsAnoAct
        frmTBanGrd.uorstMain_1Fil!mespvs = gsMesAct
        frmTBanGrd.uorstMain_1Fil!coddro = frmTBanCab.txtLlave(0).Text
        frmTBanGrd.uorstMain_1Fil!nroban = frmTBanCab.txtLlave(1).Text
        frmTBanGrd.uorstMain_1Fil!nroitem = pnNroIte
        frmTBanGrd.uorstMain_1Fil!CodCta = frmTBanGrd.uorstCOTCbMes!CodCta
        sCodRegistro = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!codaux), "", frmTBanGrd.uorstCOTCbMes!codaux)
        frmTBanGrd.uorstMain_1Fil!codaux = IIf(sCodRegistro = "", Null, sCodRegistro)
        sCodRegistro = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!codcco), "", frmTBanGrd.uorstCOTCbMes!codcco)
        frmTBanGrd.uorstMain_1Fil!codcco = IIf(sCodRegistro = "", Null, sCodRegistro)
        sCodRegistro = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!codtdc), "", frmTBanGrd.uorstCOTCbMes!codtdc)
        frmTBanGrd.uorstMain_1Fil!codtdc = IIf(sCodRegistro = "", Null, sCodRegistro)
        frmTBanGrd.uorstMain_1Fil!serdoc = frmTBanGrd.uorstCOTCbMes!serdoc
        frmTBanGrd.uorstMain_1Fil!nrodoc = frmTBanGrd.uorstCOTCbMes!nrodoc
        frmTBanGrd.uorstMain_1Fil!GloIte = IIf(txtDato(gsIdioma + 6).Text = "", Null, txtDato(gsIdioma + 6).Text)
        frmTBanGrd.uorstMain_1Fil!GloItex = IIf(txtDato(9 - gsIdioma).Text = "", Null, txtDato(9 - gsIdioma).Text)
        frmTBanGrd.uorstMain_1Fil!RefDoc = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
        s_Moneda = IIf(IsNull(frmTBanGrd.uorstCOTCbMes!tpomon), sMoneda, frmTBanGrd.uorstCOTCbMes!tpomon)
        nImporte = CDec(IIf(s_Moneda = TPOMON_NAC, frmTBanGrd.uorstCOTCbMes!imppmn, frmTBanGrd.uorstCOTCbMes!imppme))
        nTransaccion = IIf(frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING, IIf(nImporte < 0, TPOBAN_EGR, TPOBAN_ING), IIf(nImporte > 0, TPOBAN_ING, TPOBAN_EGR))
        frmTBanGrd.uorstMain_1Fil!tpoban = nTransaccion
        frmTBanGrd.uorstMain_1Fil!tpomon = s_Moneda
        frmTBanGrd.uorstMain_1Fil!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
        frmTBanGrd.uorstMain_1Fil!ImpTCb = CDec(txtDato(10).Text)
        frmTBanGrd.uorstMain_1Fil!ImpMN = CDec(Abs(frmTBanGrd.uorstCOTCbMes!imppmn))
        frmTBanGrd.uorstMain_1Fil!ImpME = CDec(Abs(frmTBanGrd.uorstCOTCbMes!imppme))
        frmTBanGrd.uorstMain_1Fil!UsrCre = gsAbvUsr
        frmTBanGrd.uorstMain_1Fil!FyHCre = Now
        
        frmTBanGrd.uorstMain_1Fil!TpoCTA = Choose(cbotpocta.ListIndex + 1, TPOCTA_AHO_IND, TPOCTA_COR_IND, TPOCTA_MAE_IND, TPOCTA_SIN_IND)
        frmTBanGrd.uorstMain_1Fil!codbco = txtDato(4).Text
      
        frmTBanGrd.uorstMain_1Fil.Update
        
        ' actualizo los datos de la cabecera de bancos
        nImporteMN = CDec(frmTBanGrd.uorstCOTCbMes!imppmn)
        nImporteME = CDec(frmTBanGrd.uorstCOTCbMes!imppme)
        If frmTBanCab.cboTpoBan.ListIndex = TPOBAN_ING Then
          nImporteMN = nImporteMN * Choose(nTransaccion + 1, 1, -1)
          nImporteME = nImporteME * Choose(nTransaccion + 1, 1, -1)
        Else
          nImporteMN = nImporteMN * Choose(nTransaccion + 1, -1, 1)
          nImporteME = nImporteME * Choose(nTransaccion + 1, -1, 1)
        End If
        frmTBanGrd.uorstMain_0Fil!ImpMN = Round(frmTBanGrd.uorstMain_0Fil!ImpMN + nImporteMN, 2)
        frmTBanGrd.uorstMain_0Fil!ImpME = Round(frmTBanGrd.uorstMain_0Fil!ImpME + nImporteME, 2)
        frmTBanGrd.uorstMain_0Fil.Update
        
        pnNroIte = pnNroIte + 1
        frmTBanGrd.uorstCOTCbMes.MoveNext
      Wend
    End If
  End If
  frmTBanGrd.uorstCOTCbMes.Close
  Exit Sub
Err:
  gpErrores

  Resume

End Sub

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  Dim dnContador As Integer
  Dim nImporteMN As Double, nImporteME As Double
  
  On Error GoTo Err
  With frmTBanGrd                     'Cambiar Formulario de Grid.
    If tnFase = 0 Then
      'Llaves.
      If Not proceso Then
        If pbNuevo Then
          .uorstMain_1!codemp = gsCodEmp
          .uorstMain_1!pdoano = gsAnoAct
          .uorstMain_1!mespvs = gsMesAct
          .uorstMain_1!coddro = frmTBanCab.txtLlave(0).Text
          .uorstMain_1!nroban = frmTBanCab.txtLlave(1).Text
          ' Obtengo el numero de Item
          pnItemBan = frmTBanGrd.pfNumItemBan(gsAnoAct, gsMesAct, frmTBanCab.txtLlave(0).Text, frmTBanCab.txtLlave(1).Text)
          pnNroIte = pnItemBan
          .uorstMain_1!nroitem = pnNroIte
        End If
      Else
        If pbNuevo Then
          .uorstMain_1Fil!codemp = gsCodEmp
          .uorstMain_1Fil!pdoano = gsAnoAct
          .uorstMain_1Fil!mespvs = gsMesAct
          .uorstMain_1Fil!coddro = frmTBanCab.txtLlave(0).Text
          .uorstMain_1Fil!nroban = frmTBanCab.txtLlave(1).Text
          ' Obtengo el numero de Item
          pnItemBan = frmTBanGrd.pfNumItemBan(gsAnoAct, gsMesAct, frmTBanCab.txtLlave(0).Text, frmTBanCab.txtLlave(1).Text)
          pnNroIte = pnItemBan
          .uorstMain_1Fil!nroitem = pnNroIte
        End If
      End If
      
      ' Reemplazo los caracteres
      txtDato(7).Text = gfSacaEntRetApos(txtDato(7).Text)
      txtDato(8).Text = gfSacaEntRetApos(txtDato(8).Text)
      
      If Not proceso Then
        'Datos.
        .uorstMain_1!CodCta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
        .uorstMain_1!codaux = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
        .uorstMain_1!codcco = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
        .uorstMain_1!codtdc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
        .uorstMain_1!serdoc = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
        .uorstMain_1!nrodoc = IIf(txtDato(6).Text = "", Null, txtDato(6).Text)
        .uorstMain_1!GloIte = IIf(txtDato(gsIdioma + 6).Text = "", Null, txtDato(gsIdioma + 6).Text)
        .uorstMain_1!GloItex = IIf(txtDato(9 - gsIdioma).Text = "", Null, txtDato(9 - gsIdioma).Text)
        .uorstMain_1!RefDoc = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
        .uorstMain_1!pvsdoc = IIf(chkPvsDoc.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
        psGlosa = txtDato(7).Text
        psGlosax = txtDato(8).Text
        dnContador = IIf(IsNull(.uorstMain_1!tpoban), cboTpoBan.ListIndex, .uorstMain_1!tpoban)
        .uorstMain_1!tpoban = cboTpoBan.ListIndex
        .uorstMain_1!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
        pnTpoMon = cboTpoMon.ListIndex
        .uorstMain_1!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
        .uorstMain_1!ImpTCb = CDec(txtDato(10).Text)
        txtImporte(0).Tag = CDec(.uorstMain_1!ImpMN)
        txtImporte(1).Tag = CDec(.uorstMain_1!ImpME)
        .uorstMain_1!ImpMN = CDec(txtImporte(0).Text)
        .uorstMain_1!ImpME = CDec(txtImporte(1).Text)
        .uorstMain_1!codbco = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
        .uorstMain_1!TpoCTA = Choose(cbotpocta.ListIndex + 1, TPOCTA_AHO_IND, TPOCTA_COR_IND, TPOCTA_MAE_IND, TPOCTA_SIN_IND)
      Else
          'Datos.
        .uorstMain_1Fil!CodCta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
        .uorstMain_1Fil!codaux = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
        .uorstMain_1Fil!codcco = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
        .uorstMain_1Fil!codtdc = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
        .uorstMain_1Fil!serdoc = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
        .uorstMain_1Fil!nrodoc = IIf(txtDato(6).Text = "", Null, txtDato(6).Text)
        .uorstMain_1Fil!GloIte = IIf(txtDato(gsIdioma + 6).Text = "", Null, txtDato(gsIdioma + 6).Text)
        .uorstMain_1Fil!GloItex = IIf(txtDato(9 - gsIdioma).Text = "", Null, txtDato(9 - gsIdioma).Text)
        .uorstMain_1Fil!RefDoc = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
        .uorstMain_1Fil!pvsdoc = IIf(chkPvsDoc.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
        psGlosa = txtDato(7).Text
        psGlosax = txtDato(8).Text
        dnContador = IIf(IsNull(.uorstMain_1Fil!tpoban), cboTpoBan.ListIndex, .uorstMain_1Fil!tpoban)
        .uorstMain_1Fil!tpoban = cboTpoBan.ListIndex
        .uorstMain_1Fil!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
        pnTpoMon = cboTpoMon.ListIndex
        .uorstMain_1Fil!TpoTcb = IIf(cboTpoTCb.ListIndex = TPOTCB_VTA_IND, TPOTCB_VTA, TPOTCB_CPR)
        .uorstMain_1Fil!ImpTCb = CDec(txtDato(10).Text)
        txtImporte(0).Tag = CDec(.uorstMain_1Fil!ImpMN)
        txtImporte(1).Tag = CDec(.uorstMain_1Fil!ImpME)
        .uorstMain_1Fil!ImpMN = CDec(txtImporte(0).Text)
        .uorstMain_1Fil!ImpME = CDec(txtImporte(1).Text)
        .uorstMain_1Fil!codbco = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
        .uorstMain_1Fil!TpoCTA = Choose(cbotpocta.ListIndex + 1, TPOCTA_AHO_IND, TPOCTA_COR_IND, TPOCTA_MAE_IND, TPOCTA_SIN_IND)
      End If
      
      '.uorstMain_1!tpocta = cbotpocta.ListIndex
      ' Actualizo los importes de la cabecera
      nImporteMN = CDec(txtImporte(0).Tag) * Choose(dnContador + 1, 1, -1)
      nImporteME = CDec(txtImporte(1).Tag) * Choose(dnContador + 1, 1, -1)
      If Not proceso Then
        If Not pbNuevo Then
          .uorstMain_0!ImpMN = Round(.uorstMain_0!ImpMN - nImporteMN, 2)
          .uorstMain_0!ImpME = Round(.uorstMain_0!ImpME - nImporteME, 2)
        End If
      Else
        If Not pbNuevo Then
          .uorstMain_0Fil!ImpMN = Round(.uorstMain_0Fil!ImpMN - nImporteMN, 2)
          .uorstMain_0Fil!ImpME = Round(.uorstMain_0Fil!ImpME - nImporteME, 2)
        End If
      End If
      
      nImporteMN = CDec(txtImporte(0).Text) * Choose(cboTpoBan.ListIndex + 1, 1, -1)
      nImporteME = CDec(txtImporte(1).Text) * Choose(cboTpoBan.ListIndex + 1, 1, -1)
      If proceso = False Then
        .uorstMain_0!ImpMN = Round(.uorstMain_0!ImpMN + nImporteMN, 2)
        .uorstMain_0!ImpME = Round(.uorstMain_0!ImpME + nImporteME, 2)
      Else
        .uorstMain_0Fil!ImpMN = Round(.uorstMain_0Fil!ImpMN + nImporteMN, 2)
        .uorstMain_0Fil!ImpME = Round(.uorstMain_0Fil!ImpME + nImporteME, 2)
      End If
      
    Else
      'Datos.
      On Error GoTo Err

      If proceso = False Then
        If .uorstMain_1.EOF Then .uorstMain_1.MoveLast
        pnUltIte = .uorstMain_1!nroitem
        dtpFehOpe.Value = frmTBanCab.dtpFehBan
        txtDato(0).Text = IIf(IsNull(.uorstMain_1!CodCta), "", .uorstMain_1!CodCta)
        txtDato(1).Text = IIf(IsNull(.uorstMain_1!codaux), "", .uorstMain_1!codaux)
        txtDato(2).Text = IIf(IsNull(.uorstMain_1!codcco), "", .uorstMain_1!codcco)
        txtDato(3).Text = IIf(IsNull(.uorstMain_1!codtdc), "", .uorstMain_1!codtdc)
        txtDato(5).Text = IIf(IsNull(.uorstMain_1!serdoc), "", .uorstMain_1!serdoc)
        txtDato(6).Text = IIf(IsNull(.uorstMain_1!nrodoc), "", .uorstMain_1!nrodoc)
        txtDato(gsIdioma + 6).Text = IIf(IsNull(.uorstMain_1!GloIte), "", .uorstMain_1!GloIte)
        txtDato(9 - gsIdioma).Text = IIf(IsNull(.uorstMain_1!GloItex), "", .uorstMain_1!GloItex)
        txtDato(9).Text = IIf(IsNull(.uorstMain_1!RefDoc), "", .uorstMain_1!RefDoc)
        chkPvsDoc.Value = .uorstMain_1!pvsdoc
        cboTpoBan.ListIndex = .uorstMain_1!tpoban
        cboTpoMon.ListIndex = IIf(.uorstMain_1!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
        cboTpoTCb.ListIndex = IIf(.uorstMain_1!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
        cbotpocta.ListIndex = IIf(.uorstMain_1!TpoCTA = TPOCTA_AHO_IND, TPOCTA_AHO_IND, IIf(.uorstMain_1!TpoCTA = TPOCTA_COR_IND, TPOCTA_COR_IND, IIf(.uorstMain_1!TpoCTA = TPOCTA_MAE_IND, TPOCTA_MAE_IND, TPOCTA_SIN_IND)))
        txtDato(10).Text = Format(IIf(IsNull(.uorstMain_1!ImpTCb), 0, .uorstMain_1!ImpTCb), FORMATO_NUM_2)
        txtImporte(0).Text = Format(IIf(IsNull(.uorstMain_1!ImpMN), 0, .uorstMain_1!ImpMN), FORMATO_NUM_1)
        txtImporte(1).Text = Format(IIf(IsNull(.uorstMain_1!ImpME), 0, .uorstMain_1!ImpME), FORMATO_NUM_1)
        txtImporte(0).Tag = Format(txtImporte(0).Text, FORMATO_NUM_1)
        txtImporte(1).Tag = Format(txtImporte(1).Text, FORMATO_NUM_1)
        txtDato(4).Text = IIf(IsNull(.uorstMain_1!codbco), "", .uorstMain_1!codbco)
      Else
        If .uorstMain_1Fil.EOF Then .uorstMain_1Fil.MoveLast
        pnUltIte = .uorstMain_1Fil!nroitem
        dtpFehOpe.Value = frmTBanCab.dtpFehBan
        txtDato(0).Text = IIf(IsNull(.uorstMain_1Fil!CodCta), "", .uorstMain_1Fil!CodCta)
        txtDato(1).Text = IIf(IsNull(.uorstMain_1Fil!codaux), "", .uorstMain_1Fil!codaux)
        txtDato(2).Text = IIf(IsNull(.uorstMain_1Fil!codcco), "", .uorstMain_1Fil!codcco)
        txtDato(3).Text = IIf(IsNull(.uorstMain_1Fil!codtdc), "", .uorstMain_1Fil!codtdc)
        txtDato(5).Text = IIf(IsNull(.uorstMain_1Fil!serdoc), "", .uorstMain_1Fil!serdoc)
        txtDato(6).Text = IIf(IsNull(.uorstMain_1Fil!nrodoc), "", .uorstMain_1Fil!nrodoc)
        txtDato(gsIdioma + 6).Text = IIf(IsNull(.uorstMain_1Fil!GloIte), "", .uorstMain_1Fil!GloIte)
        txtDato(9 - gsIdioma).Text = IIf(IsNull(.uorstMain_1Fil!GloItex), "", .uorstMain_1Fil!GloItex)
        txtDato(9).Text = IIf(IsNull(.uorstMain_1Fil!RefDoc), "", .uorstMain_1Fil!RefDoc)
        chkPvsDoc.Value = .uorstMain_1Fil!pvsdoc
        cboTpoBan.ListIndex = .uorstMain_1Fil!tpoban
        cboTpoMon.ListIndex = IIf(.uorstMain_1Fil!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
        cboTpoTCb.ListIndex = IIf(.uorstMain_1Fil!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
        cbotpocta.ListIndex = IIf(.uorstMain_1Fil!TpoCTA = TPOCTA_AHO_IND, TPOCTA_AHO_IND, IIf(.uorstMain_1Fil!TpoCTA = TPOCTA_COR_IND, TPOCTA_COR_IND, IIf(.uorstMain_1Fil!TpoCTA = TPOCTA_MAE_IND, TPOCTA_MAE_IND, TPOCTA_SIN_IND)))
        txtDato(10).Text = Format(IIf(IsNull(.uorstMain_1Fil!ImpTCb), 0, .uorstMain_1Fil!ImpTCb), FORMATO_NUM_2)
        txtImporte(0).Text = Format(IIf(IsNull(.uorstMain_1Fil!ImpMN), 0, .uorstMain_1Fil!ImpMN), FORMATO_NUM_1)
        txtImporte(1).Text = Format(IIf(IsNull(.uorstMain_1Fil!ImpME), 0, .uorstMain_1Fil!ImpME), FORMATO_NUM_1)
        txtImporte(0).Tag = Format(txtImporte(0).Text, FORMATO_NUM_1)
        txtImporte(1).Tag = Format(txtImporte(1).Text, FORMATO_NUM_1)
        txtDato(4).Text = IIf(IsNull(.uorstMain_1Fil!codbco), "", .uorstMain_1Fil!codbco)
      End If
      '[ Para mostrar los totales
      With txtDeta
        For dnContador = 0 To .Count - 1
          .Item(dnContador).Text = Format(frmTBanCab.txtDeta.Item(dnContador).Text, FORMATO_NUM_1)
        Next
      End With
      ']
      ppAyuDet (0)
      ppAyuDet (1)
      ppAyuDet (2)
      ppAyuDet (3)
      ppAyuDet (4)
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
  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      If dnContador <> 4 Then
      .Item(dnContador).Text = ""
      If dnContador = 7 Then .Item(dnContador).Text = psGlosa
      If dnContador = 8 Then .Item(dnContador).Text = psGlosax
      If dnContador = 10 Then .Item(dnContador).Tag = ""
      End If
    Next
  End With
  
  chkPvsDoc.Value = INDPREGEN_INA
  txtDato(1).Text = frmTBanCab.txtDato(4).Text
  txtDato(9).Text = frmTBanCab.lblDatoDeta(10).Caption & "-" & frmTBanCab.txtDato(7).Text
  txtDato(10).Text = Format(frmTBanCab.txtDato(6).Text, FORMATO_NUM_2)
  cboTpoBan.ListIndex = frmTBanCab.cboTpoBan.ListIndex
  cboTpoMon.ListIndex = pnTpoMon
  cboTpoTCb.ListIndex = frmTBanCab.cboTpoTCb.ListIndex
  
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
  chkPvsDoc.Enabled = pbNuevo
  With txtImporte
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = (tbHabilitar And Not (chkPvsDoc.Value = vbUnchecked And pbNuevo))
    Next
  End With
  cboTpoBan.Enabled = tbHabilitar
  cboTpoMon.Enabled = tbHabilitar
  cboTpoTCb.Enabled = tbHabilitar
  dtpFehOpe.Enabled = tbHabilitar
  cbotpocta.Enabled = tbHabilitar
  cbotpocta.Enabled = tbHabilitar
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
End Sub

'[Código propio del formulario.

Private Sub txtImporte_GotFocus(Index As Integer)
   '[ Agregado por Angel
   If Val(txtDato(10).Text) = 0 Then
      txtDato(10).Text = Format(0, FORMATO_NUM_2)
      txtDato(10).SetFocus
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
      If cboTpoMon.ListIndex = TPOMON_NAC_IND And (txtImporte(1).Text = 0 Or CDec(txtImporte(0).Text) <> CDec(txtImporte(0).Tag)) Then
        txtImporte(1).Text = Format(gfRedond(CDec(txtImporte(0).Text) / CDec(txtDato(10).Text), 2), FORMATO_NUM_1)
      End If
    End If
   Case 1
    If CDec(txtImporte(Index).Text) <> 0 Then
      If cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtImporte(0).Text) = 0 Or CDec(txtImporte(1).Text) <> CDec(txtImporte(1).Tag)) Then
        txtImporte(0).Text = Format(gfRedond(CDec(txtImporte(1).Text) * CDec(txtDato(10).Text), 2), FORMATO_NUM_1)
      End If
    End If
  End Select
  
  With frmTBanCab
    .cmdCalcular_Click
    If pbNuevo Then
      If cboTpoBan.ListIndex = TPOBAN_ING Then
        txtDeta(0).Text = Format(CDec(.txtDeta(0).Text) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
        txtDeta(2).Text = Format(CDec(.txtDeta(2).Text) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
        .txtDeta(0).Text = Format(CDec(.txtDeta(0).Text) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
        .txtDeta(2).Text = Format(CDec(.txtDeta(2).Text) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
      Else
        txtDeta(1).Text = Format(CDec(.txtDeta(1).Text) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
        txtDeta(3).Text = Format(CDec(.txtDeta(3).Text) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
        .txtDeta(1).Text = Format(CDec(.txtDeta(1).Text) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
        .txtDeta(3).Text = Format(CDec(.txtDeta(3).Text) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
      End If
    Else
      If cboTpoBan.ListIndex = TPOBAN_ING Then
        txtDeta(0).Text = Format(CDec(.txtDeta(0).Text) - CDec(txtImporte(0).Tag) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
        txtDeta(2).Text = Format(CDec(.txtDeta(2).Text) - CDec(txtImporte(1).Tag) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
      Else
        txtDeta(1).Text = Format(CDec(.txtDeta(1).Text) - CDec(txtImporte(0).Tag) + CDec(txtImporte(0).Text), FORMATO_NUM_1)
        txtDeta(3).Text = Format(CDec(.txtDeta(3).Text) - CDec(txtImporte(1).Tag) + CDec(txtImporte(1).Text), FORMATO_NUM_1)
      End If
    End If
    .txtDeta(4).Text = Format(CDec(.txtDeta(0).Text) - CDec(.txtDeta(1).Text), FORMATO_NUM_1)
    .txtDeta(5).Text = Format(CDec(.txtDeta(2).Text) - CDec(.txtDeta(3).Text), FORMATO_NUM_1)
    txtImporte(Index).Tag = Format(CDec(txtImporte(Index).Text), FORMATO_NUM_1)
    txtImporte(Index).Text = Format(CDec(txtImporte(Index).Text), FORMATO_NUM_1)
  End With
  
End Sub

Private Sub dtpFehOpe_LostFocus()
  If Month(dtpFehOpe.Value) <> Val(gsMesAct) Or Year(dtpFehOpe.Value) <> Val(gsAnoAct) Then
    If Not ((Format(dtpFehOpe.Value, "yyyymmdd") < Format(dtpFehOpe.MinDate, "yyyymmdd")) Or (Format(dtpFehOpe.Value, "yyyymmdd") > Format(dtpFehOpe.MaxDate, "yyyymmdd"))) Then Exit Sub
    MsgBox Choose(gsIdioma, "La fecha debe ser del Rango permitido que se provisiona.", "The date must be in permited range that provision."), vbExclamation
    dtpFehOpe.SetFocus
  End If
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

