VERSION 5.00
Begin VB.Form frmMCta 
   Caption         =   "[Entidad]"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRangos 
      Caption         =   "Cuentas de Cierre "
      ForeColor       =   &H00800000&
      Height          =   840
      Index           =   2
      Left            =   60
      TabIndex        =   56
      Top             =   5550
      Width           =   9915
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
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   58
         Top             =   390
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   13
         Left            =   9555
         Picture         =   "frmMCta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   390
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
         Height          =   285
         Index           =   13
         Left            =   5010
         TabIndex        =   61
         Top             =   390
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   4665
         Picture         =   "frmMCta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   390
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
         Height          =   285
         Index           =   13
         Left            =   5970
         TabIndex        =   62
         Top             =   390
         Width           =   3600
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
         Left            =   1080
         TabIndex        =   59
         Top             =   390
         Width           =   3600
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Acreedora : "
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
         Index           =   20
         Left            =   5070
         TabIndex        =   60
         Top             =   195
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Deudora :"
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
         Index           =   17
         Left            =   180
         TabIndex        =   57
         Top             =   195
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Height          =   280
      Index           =   12
      Left            =   4800
      Picture         =   "frmMCta.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2115
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
      Index           =   12
      Left            =   705
      TabIndex        =   18
      Top             =   2115
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
      Height          =   285
      Index           =   11
      Left            =   5340
      TabIndex        =   21
      Top             =   2115
      Width           =   690
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Height          =   280
      Index           =   11
      Left            =   9690
      Picture         =   "frmMCta.frx":04FE
      Style           =   1  'Graphical
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   2115
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
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   930
      Width           =   6225
   End
   Begin VB.CheckBox chkIndCCo 
      Caption         =   "Solicitar &Centro de Costos"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3285
      TabIndex        =   14
      Top             =   1845
      Width           =   2700
   End
   Begin VB.ComboBox cboNatCta 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1305
      Width           =   2000
   End
   Begin VB.CheckBox chkIndMoe 
      Caption         =   "Cuenta de &Orden"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7275
      TabIndex        =   16
      Top             =   1575
      Width           =   2700
   End
   Begin VB.CheckBox chkIndPsp 
      Caption         =   "Cuenta de &Presupuesto"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7275
      TabIndex        =   15
      Top             =   1290
      Width           =   2700
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Cuentas de Destino"
      ForeColor       =   &H00800000&
      Height          =   1335
      Index           =   0
      Left            =   60
      TabIndex        =   23
      Top             =   2400
      Width           =   9915
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   5
         Left            =   9570
         Picture         =   "frmMCta.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   915
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
         Height          =   285
         Index           =   5
         Left            =   5220
         TabIndex        =   34
         Top             =   915
         Width           =   690
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   3
         Left            =   9570
         Picture         =   "frmMCta.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   390
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
         Height          =   285
         Index           =   3
         Left            =   5220
         TabIndex        =   28
         Top             =   390
         Width           =   690
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
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   390
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   2
         Left            =   4740
         Picture         =   "frmMCta.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   390
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
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   915
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   4
         Left            =   4740
         Picture         =   "frmMCta.frx":0BA6
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   915
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
         Height          =   285
         Index           =   5
         Left            =   5910
         TabIndex        =   35
         Top             =   915
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
         Height          =   285
         Index           =   3
         Left            =   5910
         TabIndex        =   29
         Top             =   390
         Width           =   3675
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo - Haber : "
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
         Left            =   5280
         TabIndex        =   33
         Top             =   720
         Width           =   1890
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo - Debe : "
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
         Left            =   5280
         TabIndex        =   27
         Top             =   195
         Width           =   1830
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Debe: "
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
         Left            =   180
         TabIndex        =   24
         Top             =   195
         Width           =   465
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Haber: "
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
         Left            =   180
         TabIndex        =   30
         Top             =   720
         Width           =   525
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
         Left            =   1080
         TabIndex        =   26
         Top             =   390
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
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   32
         Top             =   915
         Width           =   3675
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Diferencia de Cambio"
      ForeColor       =   &H00800000&
      Height          =   1755
      Index           =   1
      Left            =   60
      TabIndex        =   36
      Top             =   3765
      Width           =   9915
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
         Height          =   285
         Index           =   7
         Left            =   5220
         TabIndex        =   48
         Top             =   840
         Width           =   690
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   7
         Left            =   9570
         Picture         =   "frmMCta.frx":0D50
         Style           =   1  'Graphical
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   840
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
         Height          =   285
         Index           =   9
         Left            =   5220
         TabIndex        =   54
         Top             =   1380
         Width           =   690
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   9
         Left            =   9570
         Picture         =   "frmMCta.frx":0EFA
         Style           =   1  'Graphical
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1380
         Width           =   255
      End
      Begin VB.ComboBox cboTpoAnl 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   255
         Width           =   1575
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
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   6
         Left            =   4740
         Picture         =   "frmMCta.frx":10A4
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   840
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
         Index           =   8
         Left            =   120
         TabIndex        =   51
         Top             =   1380
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   8
         Left            =   4740
         Picture         =   "frmMCta.frx":124E
         Style           =   1  'Graphical
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1380
         Width           =   255
      End
      Begin VB.ComboBox cboTpoMon 
         Height          =   315
         Left            =   5820
         TabIndex        =   41
         Top             =   255
         Width           =   675
      End
      Begin VB.ComboBox cboTpoTCb 
         Height          =   315
         Left            =   7995
         TabIndex        =   43
         Top             =   255
         Width           =   975
      End
      Begin VB.CheckBox chkIndAjD 
         Caption         =   "A&plicar"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   315
         Width           =   795
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo - Ganancia : "
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
         Height          =   200
         Index           =   14
         Left            =   5280
         TabIndex        =   47
         Top             =   615
         Width           =   2145
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo - Perdida : "
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
         Height          =   200
         Index           =   16
         Left            =   5280
         TabIndex        =   53
         Top             =   1155
         Width           =   1995
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
         Index           =   7
         Left            =   5910
         TabIndex        =   49
         Top             =   840
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
         Height          =   285
         Index           =   9
         Left            =   5910
         TabIndex        =   55
         Top             =   1380
         Width           =   3675
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Análisis:"
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
         Left            =   1500
         TabIndex        =   38
         Top             =   315
         Width           =   1185
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
         Height          =   280
         Index           =   6
         Left            =   1080
         TabIndex        =   46
         Top             =   840
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
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   52
         Top             =   1380
         Width           =   3675
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Moneda:"
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
         Index           =   11
         Left            =   4620
         TabIndex        =   40
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio: "
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
         Left            =   6855
         TabIndex        =   42
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Ganancia: "
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
         Height          =   200
         Index           =   13
         Left            =   180
         TabIndex        =   44
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Pérdida: "
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
         Height          =   200
         Index           =   15
         Left            =   180
         TabIndex        =   50
         Top             =   1155
         Width           =   1185
      End
   End
   Begin VB.CheckBox chkIndFjo 
      Caption         =   "Solicitar &Flujo de Caja"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3285
      TabIndex        =   13
      Top             =   1575
      Width           =   2700
   End
   Begin VB.CheckBox chkIndDoc 
      Caption         =   "Solicitar &Documento (Cta.Cte.)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3285
      TabIndex        =   12
      Top             =   1290
      Width           =   2700
   End
   Begin VB.ComboBox cboTpoSdo 
      Height          =   315
      ItemData        =   "frmMCta.frx":13F8
      Left            =   1200
      List            =   "frmMCta.frx":13FA
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1665
      Width           =   2000
   End
   Begin VB.ComboBox cboTpoCta 
      Height          =   315
      Left            =   8625
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   570
      Width           =   1275
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
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   570
      Width           =   6225
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   60
      TabIndex        =   63
      Top             =   6675
      Width           =   1170
      Begin VB.CheckBox chkEstCta 
         Caption         =   "&Activa"
         ForeColor       =   &H00C00000&
         Height          =   200
         Left            =   120
         TabIndex        =   64
         Top             =   195
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   3135
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   6540
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
         Picture         =   "frmMCta.frx":13FC
         Style           =   1  'Graphical
         TabIndex        =   65
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
         Picture         =   "frmMCta.frx":15A6
         Style           =   1  'Graphical
         TabIndex        =   66
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
         Left            =   480
         Picture         =   "frmMCta.frx":1750
         Style           =   1  'Graphical
         TabIndex        =   67
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
         Left            =   1220
         Picture         =   "frmMCta.frx":189A
         Style           =   1  'Graphical
         TabIndex        =   68
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
         Left            =   1950
         Picture         =   "frmMCta.frx":199C
         Style           =   1  'Graphical
         TabIndex        =   69
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
         Left            =   2690
         Picture         =   "frmMCta.frx":1A9E
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   60
         Width           =   720
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
      Height          =   285
      Index           =   0
      Left            =   705
      TabIndex        =   1
      Top             =   90
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
      Height          =   285
      Index           =   12
      Left            =   1140
      TabIndex        =   19
      Top             =   2115
      Width           =   3675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Banco : "
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
      Index           =   19
      Left            =   60
      TabIndex        =   17
      Top             =   2115
      Width           =   600
   End
   Begin VB.Label lblTexto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Centro de Costo Default : "
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
      Index           =   18
      Left            =   7815
      TabIndex        =   20
      Top             =   1905
      Width           =   2070
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
      Index           =   11
      Left            =   6030
      TabIndex        =   22
      Top             =   2115
      Width           =   3675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducción:"
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
      Left            =   60
      TabIndex        =   6
      Top             =   990
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nat de Cuenta:"
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
      Left            =   60
      TabIndex        =   8
      Top             =   1365
      Width           =   1065
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Saldo:"
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
      Left            =   60
      TabIndex        =   10
      Top             =   1725
      Width           =   1020
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cuenta:"
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
      Index           =   2
      Left            =   7485
      TabIndex        =   4
      Top             =   630
      Width           =   1125
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
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
      Left            =   60
      TabIndex        =   2
      Top             =   630
      Width           =   900
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
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   9880
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmMCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

Private Sub chkIndCCo_Click()
  txtDato(11).Text = IIf(chkIndCCo.Value = vbUnchecked, "", txtDato(11).Text)
  lblDatoDeta(11).Caption = IIf(chkIndCCo.Value = vbUnchecked, "", lblDatoDeta(11).Caption)
  txtDato(11).Enabled = (chkIndCCo.Value = vbChecked And chkIndCCo.Enabled)
  cmdDatoAyud(11).Enabled = (chkIndCCo.Value = vbChecked And chkIndCCo.Enabled)
End Sub

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmMCtaGrd                     'Cambiar Formulario de Grid.
    '[Llaves.                          'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codcta.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
      With cboTpoCta
         .AddItem TPOCTA_TRA_TXT, 0
         .AddItem TPOCTA_TIT_TXT, 1
      End With
      With cboNatCta
         .AddItem NATCTA_DEU_TXT, 0
         .AddItem NATCTA_ACR_TXT, 1
      End With
      With cboTpoSdo
         .AddItem TPOSDO_INV_TXT, 0
         .AddItem TPOSDO_RES_TXT, 1
         .AddItem TPOSDO_FUN_TXT, 2
         .AddItem TPOSDO_NAT_TXT, 3
         .AddItem TPOSDO_AMB_TXT, 4
      End With
      With cboTpoAnl
         .AddItem TPOANL_CTA_TXT, 0
         .AddItem TPOANL_AUX_TXT, 1
         .AddItem TPOANL_DOC_TXT, 2
      End With
      With cboTpoMon
         .AddItem TPOMON_NAC_TXT_0, 0
         .AddItem TPOMON_EXT_TXT_0, 1
      End With
      With cboTpoTCb
         .AddItem TPOTCB_CPR_TXT, 0
         .AddItem TPOTCB_VTA_TXT, 1
      End With
'      mskDato(0).MaxLength = .uorstMain!Tf1Cta.DefinedSize + 1
      txtDato(gsIdioma - 1).MaxLength = .uorstMain!detcta.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain!DetCtax.DefinedSize
      txtDato(12).MaxLength = .uorstMain!codbco.DefinedSize
      txtDato(11).MaxLength = .uorstMain!codcco_def.DefinedSize
      
      txtDato(2).MaxLength = .uorstMain!codcta_dst_deb.DefinedSize
      txtDato(3).MaxLength = .uorstMain!codcco_dst_deb.DefinedSize
      txtDato(4).MaxLength = .uorstMain!codcta_dst_hab.DefinedSize
      txtDato(5).MaxLength = .uorstMain!codcco_dst_hab.DefinedSize
      
      txtDato(6).MaxLength = .uorstMain!CodCta_AjD_Deb.DefinedSize
      txtDato(7).MaxLength = .uorstMain!CodCCo_AjD_Deb.DefinedSize
      txtDato(8).MaxLength = .uorstMain!CodCta_AjD_Hab.DefinedSize
      txtDato(9).MaxLength = .uorstMain!CodCCo_AjD_Hab.DefinedSize
      
      txtDato(10).MaxLength = .uorstMain!codcta_crr_deu.DefinedSize
      txtDato(13).MaxLength = .uorstMain!codcta_crr_acr.DefinedSize
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
'[Propio del formulario.
   ppDeshabilitar_DC
']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(21, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuenta :", "Descripción :", "Tipo de Cuenta :", "Traducción :", "Nat de Cuenta :", "Tipo de Saldo :", "Debe :", "Centro de Costo - Debe :", "Haber :", "Centro de Costo - Haber :", "Tipo de Análisis :", "Tipo de Moneda :", "Tipo de Cambio :", "Cuenta Ganancia :", "Centro de Costo - Ganancia :", "Cuenta Pérdida :", "Centro de Costo - Pérdida :", "Deudora :", "Centro de Costo Default :", "Banco :", "Acreedora :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Account :", "Description :", "Tipo Account :", "Traduction :", "Natur Account :", "Type Balance :", "Debit :", "Cost Center - Debit :", "Credit :", "Cost Center - Credit :", "Tpy Analysis :", "Type Currency :", "Type Exchange :", "Profit Account :", "Cost Center - Profit :", "Loss Account:", "Cost Center - Loss :", "Debit :", "Cost Center Deafult :", "Bank :", "Credit :")
  Next nElemento
  chkIndDoc.Caption = Choose(gsIdioma, "Solicitar &Documento (Cta.Cte.)", "Request &Document (Current Acc.)")
  chkIndCCo.Caption = Choose(gsIdioma, "Solicitar &Centro de Costos", "Request &Cost Center")
  chkIndFjo.Caption = Choose(gsIdioma, "Solicitar &Flujo de Caja", "Request Cash &Flow")
  chkIndPsp.Caption = Choose(gsIdioma, "Cuenta de &Presupuesto", "Account of &Budget")
  chkIndMoe.Caption = Choose(gsIdioma, "Cuenta de &Orden", "Account of &Order")
  fraRangos(0).Caption = Choose(gsIdioma, "Cuentas de Destino", "Destiny Accounts")
  fraRangos(1).Caption = Choose(gsIdioma, "Diferencia de Cambio", "Difference of Exchange")
  fraRangos(2).Caption = Choose(gsIdioma, " Cuentas de Cierre ", " Closing Accounts ")
  chkIndAjD.Caption = Choose(gsIdioma, "A&plicar", "A&pply")
  chkEstCta.Caption = Choose(gsIdioma, "&Activa", "&Active")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']

End Sub

Private Sub Form_Activate()
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMCtaGrd.uorstMain.BOF And frmMCtaGrd.uorstMain.EOF) Then
    frmMCtaGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMCtaGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMCtaGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
'[Propio del formulario.
   If txtLlave(0).Text = CTAFZD_CTA Then
      MsgBox Choose(gsIdioma, "No es posible corregir los datos de esta cuenta.", " you can not correct the data of this account."), vbCritical
      Exit Sub
   End If
']

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

   With frmMCtaGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain
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
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain.Find "CodCta='" & txtLlave(0).Text & "'"
       ']
         cmdGrabar.Enabled = False
         upHabilitacion False
   
         upDatosPredeterminados
       '[Llave con el foco al añadir.  'Cambiar.
         txtLlave(0).SetFocus
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
      
   Exit Sub
Err:
   gpErrores
  
   frmMCtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 2 To 13
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
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

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
 '[Convierte a mayúsculas.
   If Index = 0 Then                   'Cambiar (añadir índices).
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
 ']
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
'         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
'      End If
'   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMCtaGrd.uorstMain        'Cambiar Formulario de Grid.
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodCta='" & txtLlave(0).Text & "'"
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
      upHabilitacion True
      pbValidada = True
   Else
      cmdGrabar.Enabled = False
      upHabilitacion False
      pbValidada = False
   End If
      
'[Propio del formulario.
'[REVISAR
   Dim lcad, lgs, i, j As Integer
   Dim scad, sgs, v As String
   Dim pvRegistroActual As Variant
   Dim A() As Integer
   
   frmMCtaGrd.uorstCoCta.Requery 'Para actualizar en caso se acabe de crear cuenta de nivel inferior.
   
   scad = txtLlave(0).Text
   lcad = CInt(Len(txtLlave(0).Text))
   sgs = gsNivCta
   lgs = CInt(Len(gsNivCta))
   
   ' redimensiono el arreglo
   ReDim A(lgs) As Integer
   
   If txtLlave(0).Text = "" Then Exit Sub
   For j = 1 To lgs
      A(j) = CInt(Mid(sgs, j, 1))
      If lcad = A(j) Then
         With frmMCtaGrd.uorstCoCta
            For i = j To 1 Step -1
               If i = 1 Then Exit Sub
               v = Mid(scad, 1, A(i - 1))
'               pvRegistroActual = .Bookmark
               .MoveFirst
               .Find "CodCta = '" & v & "'"
               If .EOF Then
                  Cancel = True
                  MsgBox Choose(gsIdioma, "La cuenta ", " The account ") & v & Choose(gsIdioma, " (nivel anterior) NO EXISTE.", " (previous level) NOT EXIST"), vbInformation + vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
                  txtLlave(0).SetFocus
                  txtLlave(Index).SelStart = 0
                  txtLlave(Index).SelLength = txtLlave(Index).MaxLength
               End If
'            .Bookmark = pvRegistroActual
            Next i
         End With
         txtDato(0).SetFocus
      End If
      If j = lgs Then
        MsgBox Choose(gsIdioma, "El nivel ", "The level ") & Str(lcad) & Choose(gsIdioma, " no está configurado.", "is not configurated"), vbInformation + vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
        Cancel = True
        txtLlave(0).SetFocus
      End If
   Next j
']REVISAR.
']

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
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 3                              'Cambiar (añadir índices).
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
  Select Case Index
   Case 2 To 13                         'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    ' Valido centro de costo por cuenta
    If (Index = 2 Or Index = 4 Or Index = 6 Or Index = 8) Then
      txtDato(Index + 1).Text = IIf(txtDato(Index).Text = "", "", txtDato(Index + 1).Text)
      lblDatoDeta(Index + 1).Caption = IIf(txtDato(Index).Text = "", "", lblDatoDeta(Index + 1).Caption)
      txtDato(Index + 1).Enabled = (txtDato(Index).Text = "")
      cmdDatoAyud(Index + 1).Enabled = (txtDato(Index).Text = "")
      If txtDato(Index).Text = "" Then Exit Sub
      If frmMCtaGrd.uorstCoCta.RecordCount > 0 Then
        If Not frmMCtaGrd.uorstCoCta.EOF Then
          txtDato(Index + 1).Text = IIf(frmMCtaGrd.uorstCoCta!indcco = INDCCO_ACT, txtDato(Index + 1).Text, "")
          lblDatoDeta(Index + 1).Caption = IIf(frmMCtaGrd.uorstCoCta!indcco = INDCCO_ACT, lblDatoDeta(Index + 1).Caption, "")
          txtDato(Index + 1).Enabled = (frmMCtaGrd.uorstCoCta!indcco = INDCCO_ACT)
          cmdDatoAyud(Index + 1).Enabled = (frmMCtaGrd.uorstCoCta!indcco = INDCCO_ACT)
        End If
      End If
    End If
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 2, 4, 6, 8, 10, 13                     'Cambiar (añadir índices).
    modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 3, 5, 7, 9, 11
    modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 12
    modAyuBus.Bco_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 2, 4, 6, 8, 10, 13
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCtaGrd.uorstCoCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
   Case 3, 5, 7, 9, 11
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCtaGrd.uorstCoCCo
         .MoveFirst
         .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
         End If
      End With
   Case 12
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCtaGrd.uorstCoBco
         .MoveFirst
         .Find "codbco='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detbco), "", !detbco)
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMCtaGrd.uorstMain           'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !codcta = txtLlave(0).Text
         End If

        'Datos.
         !TpoCTA = Choose(cboTpoCta.ListIndex + 1, TPOCTA_TRA, TPOCTA_TIT)
         !NatCta = Choose(cboNatCta.ListIndex + 1, NATCTA_DEU, NATCTA_ACR)
         !TpoSdo = Choose(cboTpoSdo.ListIndex + 1, TPOSDO_INV, TPOSDO_RES, TPOSDO_FUN, TPOSDO_NAT, TPOSDO_AMB)
         !TpoAnl = Choose(cboTpoAnl.ListIndex + 1, TPOANL_CTA, TPOANL_AUX, TPOANL_DOC)
         !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
         !TpoTcb = Choose(cboTpoTCb.ListIndex + 1, TPOTCB_CPR, TPOTCB_VTA)
         !IndAjD = IIf(chkIndAjD.Value = vbChecked, INDAJD_ACT, INDAJD_INA)
         !indcco = IIf(chkIndCCo.Value = vbChecked, INDCCO_ACT, INDCCO_INA)
         !IndDoc = IIf(chkIndDoc.Value = vbChecked, INDDOC_ACT, INDDOC_INA)
         !IndMoe = IIf(chkIndMoe.Value = vbChecked, INDMOE_ACT, INDMOE_INA)
         !IndPsp = IIf(chkIndPsp.Value = vbChecked, INDPSP_ACT, INDPSP_INA)
         !IndFjo = IIf(chkIndFjo.Value = vbChecked, INDFJO_ACT, INDFJO_INA)
         !EstCta = IIf(chkEstCta.Value = vbChecked, ESTCTA_ACT, ESTCTA_INA)
         
         !detcta = IIf(txtDato(gsIdioma - 1).Text = "", Null, txtDato(gsIdioma - 1).Text)
         !DetCtax = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
         !codbco = IIf(txtDato(12).Text = "", Null, txtDato(12).Text)
         !codcco_def = IIf(txtDato(11).Text = "", Null, txtDato(11).Text)
         !codcta_dst_deb = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
         !codcco_dst_deb = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
         !codcta_dst_hab = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
         !codcco_dst_hab = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
         !CodCta_AjD_Deb = IIf(txtDato(6).Text = "", Null, txtDato(6).Text)
         !CodCCo_AjD_Deb = IIf(txtDato(7).Text = "", Null, txtDato(7).Text)
         !CodCta_AjD_Hab = IIf(txtDato(8).Text = "", Null, txtDato(8).Text)
         !CodCCo_AjD_Hab = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
         !codcta_crr_deu = IIf(txtDato(10).Text = "", Null, txtDato(10).Text)
         !codcta_crr_acr = IIf(txtDato(13).Text = "", Null, txtDato(13).Text)
      Else
        'Llaves.
         txtLlave(0).Text = !codcta
      
        'Datos.
         cboTpoCta.ListIndex = IIf(!TpoCTA = TPOCTA_TRA, 0, 1)
         cboNatCta.ListIndex = IIf(!NatCta = NATCTA_DEU, 0, 1)
         Select Case !TpoAnl
         Case TPOANL_CTA
            cboTpoAnl.ListIndex = 0
         Case TPOANL_AUX
            cboTpoAnl.ListIndex = 1
         Case TPOANL_DOC
            cboTpoAnl.ListIndex = 2
         End Select
         Select Case !TpoSdo
         Case TPOSDO_INV
            cboTpoSdo.ListIndex = 0
         Case TPOSDO_RES
            cboTpoSdo.ListIndex = 1
         Case TPOSDO_FUN
            cboTpoSdo.ListIndex = 2
         Case TPOSDO_NAT
            cboTpoSdo.ListIndex = 3
         Case TPOSDO_AMB
            cboTpoSdo.ListIndex = 4
         End Select
         cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, 0, 1)
         cboTpoTCb.ListIndex = IIf(!TpoTcb = TPOTCB_CPR, 0, 1)
         If !IndAjD = INDAJD_ACT Then
            chkIndAjD.Value = vbChecked
         Else
            chkIndAjD.Value = vbUnchecked
         End If
         chkIndCCo.Value = IIf(!indcco = INDCCO_ACT, vbChecked, vbUnchecked)
         chkIndDoc.Value = IIf(!IndDoc = INDDOC_ACT, vbChecked, vbUnchecked)
         chkIndMoe.Value = IIf(!IndMoe = INDMOE_ACT, vbChecked, vbUnchecked)
         chkIndPsp.Value = IIf(!IndPsp = INDPSP_ACT, vbChecked, vbUnchecked)
         chkIndFjo.Value = IIf(!IndFjo = INDFJO_ACT, vbChecked, vbUnchecked)
         chkEstCta.Value = IIf(!EstCta = ESTCTA_ACT, vbChecked, vbUnchecked)
         txtDato(gsIdioma - 1).Text = IIf(IsNull(!detcta), "", !detcta)
         txtDato(2 - gsIdioma).Text = IIf(IsNull(!DetCtax), "", !DetCtax)
         txtDato(12).Text = IIf(IsNull(!codbco), "", !codbco)
         txtDato(11).Text = IIf(IsNull(!codcco_def), "", !codcco_def)
         txtDato(2).Text = IIf(IsNull(!codcta_dst_deb), "", !codcta_dst_deb)
         txtDato(3).Text = IIf(IsNull(!codcco_dst_deb), "", !codcco_dst_deb)
         txtDato(4).Text = IIf(IsNull(!codcta_dst_hab), "", !codcta_dst_hab)
         txtDato(5).Text = IIf(IsNull(!codcco_dst_hab), "", !codcco_dst_hab)
         txtDato(6).Text = IIf(IsNull(!CodCta_AjD_Deb), "", !CodCta_AjD_Deb)
         txtDato(7).Text = IIf(IsNull(!CodCCo_AjD_Deb), "", !CodCCo_AjD_Deb)
         txtDato(8).Text = IIf(IsNull(!CodCta_AjD_Hab), "", !CodCta_AjD_Hab)
         txtDato(9).Text = IIf(IsNull(!CodCCo_AjD_Hab), "", !CodCCo_AjD_Hab)
         txtDato(10).Text = IIf(IsNull(!codcta_crr_acr), "", !codcta_crr_acr)
         txtDato(13).Text = IIf(IsNull(!codcta_crr_deu), "", !codcta_crr_deu)
'TC
        ' txtDato(10).Text = IIf(IsNull(!codcta_crr_acr), "", !codcta_crr_deu)
        ' txtDato(13).Text = IIf(IsNull(!codcta_crr_deu), "", !codcta_crr_acr)



       '[Busca detalle de códigos.     'Cambiar (habilitar/deshabilitar).
         ppAyuDet 2: ppAyuDet 3: ppAyuDet 4
         ppAyuDet 5: ppAyuDet 6: ppAyuDet 7
         ppAyuDet 8: ppAyuDet 9: ppAyuDet 10
         ppAyuDet 11: ppAyuDet 12: ppAyuDet 13
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
   txtLlave(0).Text = ""

  'Datos.
   cboTpoCta.ListIndex = 0
   cboNatCta.ListIndex = 0
   cboTpoSdo.ListIndex = 0
   cboTpoAnl.ListIndex = 0
   cboTpoMon.ListIndex = 0
   cboTpoTCb.ListIndex = 0
   chkIndAjD.Value = vbUnchecked
   chkIndCCo.Value = vbUnchecked
   chkIndDoc.Value = vbUnchecked
   chkIndMoe.Value = vbUnchecked
   chkIndPsp.Value = vbUnchecked
   chkIndFjo.Value = vbUnchecked
   chkEstCta.Value = vbChecked
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With

  'Ayudas.
   lblDatoDeta(2).Caption = ""
   lblDatoDeta(3).Caption = ""
   lblDatoDeta(4).Caption = ""
   lblDatoDeta(5).Caption = ""
   lblDatoDeta(6).Caption = ""
   lblDatoDeta(7).Caption = ""
   lblDatoDeta(8).Caption = ""
   lblDatoDeta(9).Caption = ""
   lblDatoDeta(10).Caption = ""
   lblDatoDeta(11).Caption = ""
   lblDatoDeta(12).Caption = ""
   lblDatoDeta(13).Caption = ""

End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   cboTpoCta.Enabled = tbHabilitar
   cboNatCta.Enabled = tbHabilitar
   cboTpoSdo.Enabled = tbHabilitar
   chkIndAjD.Enabled = tbHabilitar
   chkIndCCo.Enabled = tbHabilitar
   chkIndDoc.Enabled = tbHabilitar
   chkIndMoe.Enabled = tbHabilitar
   chkIndPsp.Enabled = tbHabilitar
   chkIndFjo.Enabled = tbHabilitar
   chkEstCta.Enabled = tbHabilitar
   With txtDato
      For dnContador = 0 To 5 '.Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
      .Item(10).Enabled = tbHabilitar
      .Item(11).Enabled = tbHabilitar
      .Item(12).Enabled = tbHabilitar
      .Item(13).Enabled = tbHabilitar
   End With
   
'[Propio del formulario.
   chkIndAjD_Click
']

  'Ayudas.
   cmdDatoAyud(2).Enabled = tbHabilitar
   cmdDatoAyud(3).Enabled = tbHabilitar
   cmdDatoAyud(4).Enabled = tbHabilitar
   cmdDatoAyud(5).Enabled = tbHabilitar
   cmdDatoAyud(10).Enabled = tbHabilitar
   cmdDatoAyud(11).Enabled = tbHabilitar
   cmdDatoAyud(12).Enabled = tbHabilitar
   cmdDatoAyud(13).Enabled = tbHabilitar
   lblDatoDeta(2).Enabled = tbHabilitar
   lblDatoDeta(3).Enabled = tbHabilitar
   lblDatoDeta(4).Enabled = tbHabilitar
   lblDatoDeta(5).Enabled = tbHabilitar
   lblDatoDeta(10).Enabled = tbHabilitar
   lblDatoDeta(11).Enabled = tbHabilitar
   lblDatoDeta(12).Enabled = tbHabilitar
   lblDatoDeta(13).Enabled = tbHabilitar

End Sub

'[Código propio del formulario.

'Private Sub cboTpoCta_LostFocus()
'   chkIndCCo.Enabled = (cboTpoCta.ListIndex = TPOANL_NAP)
'End Sub

Private Sub chkIndAjD_Click()
   If chkIndAjD.Value = vbChecked And chkIndAjD.Enabled Then
      ppHabilitar_DC
   Else
      ppDeshabilitar_DC
   End If
End Sub

Private Sub ppHabilitar_DC()
   cboTpoAnl.Enabled = True
   cboTpoMon.Enabled = True
   cboTpoTCb.Enabled = True
   txtDato(6).Enabled = True: txtDato(7).Enabled = True
   txtDato(8).Enabled = True: txtDato(9).Enabled = True
   cmdDatoAyud(6).Enabled = True: cmdDatoAyud(7).Enabled = True
   cmdDatoAyud(8).Enabled = True: cmdDatoAyud(9).Enabled = True
   lblDatoDeta(6).Enabled = True: lblDatoDeta(7).Enabled = True
   lblDatoDeta(8).Enabled = True: lblDatoDeta(9).Enabled = True
End Sub

Private Sub ppDeshabilitar_DC()
   cboTpoAnl.Enabled = False
   cboTpoMon.Enabled = False
   cboTpoTCb.Enabled = False
   txtDato(6).Enabled = False: txtDato(7).Enabled = False
   txtDato(8).Enabled = False: txtDato(9).Enabled = False
   cmdDatoAyud(6).Enabled = False: cmdDatoAyud(7).Enabled = False
   cmdDatoAyud(8).Enabled = False: cmdDatoAyud(9).Enabled = False
   lblDatoDeta(6).Enabled = False: lblDatoDeta(7).Enabled = False
   lblDatoDeta(8).Enabled = False: lblDatoDeta(9).Enabled = False
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

