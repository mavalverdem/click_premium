VERSION 5.00
Begin VB.Form frmMTCbCie 
   Caption         =   "[Entidad]"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   5370
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkVerificar 
      Alignment       =   1  'Right Justify
      Caption         =   "Replicar en Empresas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   5070
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   240
      Width           =   5130
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   24
         Left            =   3945
         TabIndex        =   40
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   25
         Left            =   3945
         TabIndex        =   41
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   26
         Left            =   3945
         TabIndex        =   42
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   27
         Left            =   3945
         TabIndex        =   43
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   28
         Left            =   3945
         TabIndex        =   44
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   29
         Left            =   3945
         TabIndex        =   45
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   30
         Left            =   3945
         TabIndex        =   46
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   31
         Left            =   3945
         TabIndex        =   47
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   32
         Left            =   3945
         TabIndex        =   48
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   33
         Left            =   3945
         TabIndex        =   49
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   34
         Left            =   3945
         TabIndex        =   50
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   35
         Left            =   3945
         TabIndex        =   51
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   23
         Left            =   2820
         TabIndex        =   38
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   22
         Left            =   2820
         TabIndex        =   35
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   21
         Left            =   2820
         TabIndex        =   32
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   20
         Left            =   2820
         TabIndex        =   29
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   19
         Left            =   2820
         TabIndex        =   26
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   18
         Left            =   2820
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   17
         Left            =   2820
         TabIndex        =   20
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   16
         Left            =   2820
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   15
         Left            =   2820
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   14
         Left            =   2820
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   13
         Left            =   2820
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   12
         Left            =   2820
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   11
         Left            =   1680
         TabIndex        =   37
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   10
         Left            =   1680
         TabIndex        =   34
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   9
         Left            =   1680
         TabIndex        =   31
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   8
         Left            =   1680
         TabIndex        =   28
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   7
         Left            =   1680
         TabIndex        =   25
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   6
         Left            =   1680
         TabIndex        =   22
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   5
         Left            =   1680
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   4
         Left            =   1680
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   3
         Left            =   1680
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Factor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   15
         Left            =   3945
         TabIndex        =   39
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Enero"
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
         Left            =   540
         TabIndex        =   3
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Marzo"
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
         Left            =   540
         TabIndex        =   9
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Febrero"
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
         Left            =   540
         TabIndex        =   6
         Top             =   900
         Width           =   570
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Mayo"
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
         Left            =   540
         TabIndex        =   15
         Top             =   1980
         Width           =   390
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Junio"
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
         Left            =   540
         TabIndex        =   18
         Top             =   2340
         Width           =   375
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Abril"
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
         Left            =   540
         TabIndex        =   12
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Agosto"
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
         Left            =   540
         TabIndex        =   24
         Top             =   3060
         Width           =   525
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Setiembre"
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
         Left            =   540
         TabIndex        =   27
         Top             =   3420
         Width           =   720
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Julio"
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
         Left            =   540
         TabIndex        =   21
         Top             =   2700
         Width           =   315
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Noviembre"
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
         Left            =   540
         TabIndex        =   33
         Top             =   4140
         Width           =   765
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diciembre"
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
         Left            =   540
         TabIndex        =   36
         Top             =   4500
         Width           =   705
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Octubre"
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
         Left            =   540
         TabIndex        =   30
         Top             =   3780
         Width           =   585
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   14
         Left            =   2820
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   13
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1200
      ScaleHeight     =   690
      ScaleWidth      =   2895
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5685
      Width           =   2895
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
         Left            =   0
         Picture         =   "frmMTCbCie.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Left            =   720
         Picture         =   "frmMTCbCie.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   60
         Width           =   700
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
         Left            =   1425
         Picture         =   "frmMTCbCie.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   60
         Width           =   700
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
         Left            =   2130
         Picture         =   "frmMTCbCie.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   60
         Width           =   700
      End
   End
End
Attribute VB_Name = "frmMTCbCie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain As ADODB.Recordset
'Public dvValTCb As String
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer
Private pbNuevo As Boolean

Private Sub Form_Load()
Dim dnContador As Integer
   Me.KeyPreview = True
   pbNuevo = True
   psConnStrgSele = "SELECT MesPvs, ImpTCb_Cpr, ImpTCb_Vta, impfac_hpr, "
   psConnStrgSele = psConnStrgSele & "codemp, pdoano, "
   psConnStrgSele = psConnStrgSele & "UsrCre, FyHCre, UsrMdf, FyHMdf "
   psConnStrgSele = psConnStrgSele & "FROM CoTCbMes "
   psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
   psConnStrgSele = psConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
   psConnStrgOrde = "ORDER BY 1"
   
   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic 'adLockReadOnly
      .Open
      .Properties("Unique Table").Value = "COTCBMES"
   End With

   mostrardatos
   
  ' Visualización de atributos
  chkVerificar.Caption = Choose(gsIdioma, "Replicar en Empresas ", "Replicate in Companies ")
  chkVerificar.Visible = (gsNvlUsr = NvlUsr_Adm)
  
   If pbNuevo Then
'      cmdRetroceder.Enabled = False
'      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(15, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Mes", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Compra", "Venta", "Factor")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Month", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Purchase", "Sale", "Factor")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']

End Sub

Private Sub Form_Activate()
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'   Call gpTeclasData2(KeyAscii)
'End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   frmMCCoGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
   uorstMain.Close
'   uocnnMain.Close
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Private Sub cmdCorregir_Click()
'   cmdRetroceder.Enabled = False
'   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   
    '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err
   'uocnnMain.BeginTrans   'INICIA TRANSACCION.
   
   guardardatos
   
   uorstMain.Update
   uocnnMain.CommitTrans  'CONFIRMA TRANSACCION.
   cmdSalir.SetFocus
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdDeshacer_Click()
On Error GoTo Err
   Dim dvRegistroActual As Variant
   dvRegistroActual = uorstMain.Bookmark
   uorstMain.CancelUpdate
   uorstMain.Requery
   uorstMain.Bookmark = dvRegistroActual
   zbNuevo = False
   uorstMain.CancelUpdate
   mostrardatos
   txtDato(0).Refresh
'   If pbNuevo Then
'      cmdRetroceder.Enabled = False
'   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   cmdCorregir.Enabled = True
   upHabilitacion False
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
On Error GoTo Err
  If Not (uorstMain.BOF And uorstMain.EOF) Then
    uorstMain.CancelUpdate
  End If
  Unload Me
  Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub
Private Sub txtDato_GotFocus(Index As Integer)
   txtDato.Item(Index).SelStart = 0
   txtDato.Item(Index).SelLength = txtDato.Item(Index).MaxLength
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   Dim var As Integer
   Dim ent As Integer
   var = Len(txtDato.Item(Index).Text)
   ent = Int(CDbl(Val(txtDato.Item(Index).Text)))
   If ent > 999 Then
      Cancel = True
      MsgBox Choose(gsIdioma, "Solo acepta 3 enteros", "Only accept 3 integers"), vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
      txtDato.Item(Index).SetFocus
   Else
      If CDec(Val(txtDato.Item(Index).Text)) > 9999 Then
         Cancel = True
         MsgBox Choose(gsIdioma, "Solo acepta 4 decimales", "Only accept 4 decimals"), vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
         txtDato.Item(Index).SetFocus
      Else
         txtDato.Item(Index).Text = Format(Round(CDbl(Val(txtDato.Item(Index).Text)), 4), FORMATO_NUM_2)
         Exit Sub
      End If
   End If
End Sub


Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   chkVerificar.Enabled = tbHabilitar

  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
End Sub

'[Propios del formulario
Private Sub mostrardatos()
  Dim i As Integer
  Dim pvRegistroActual As Variant
  
  With uorstMain
    If Not .EOF Then
      For i = 1 To 12
        pvRegistroActual = .Bookmark
        .MoveFirst
        .Find "MesPvs = '" & Format(i, "00") & "'"
        If .EOF Then
          txtDato(i - 1).Text = Format(Round(0, 4), FORMATO_NUM_2)
          txtDato(i + 11).Text = Format(Round(0, 4), FORMATO_NUM_2)
          txtDato(i + 23).Text = Format(Round(0, 4), FORMATO_NUM_2)
          .Bookmark = pvRegistroActual
        Else
          txtDato(i - 1).Text = Format(Round(CDec(uorstMain!ImpTCb_Cpr), 4), FORMATO_NUM_2)
          txtDato(i + 11).Text = Format(Round(CDec(uorstMain!imptcb_vta), 4), FORMATO_NUM_2)
          txtDato(i + 23).Text = Format(Round(CDec(uorstMain!impfac_hpr), 4), FORMATO_NUM_2)
          .Bookmark = pvRegistroActual
        End If
      Next i
    Else
      For i = 1 To 12
        txtDato(i - 1).Text = Format(Round(0, 4), FORMATO_NUM_2)
        txtDato(i + 11).Text = Format(Round(0, 4), FORMATO_NUM_2)
        txtDato(i + 23).Text = Format(Round(0, 4), FORMATO_NUM_2)
      Next i
      uocnnMain.BeginTrans
    End If
  End With
  chkVerificar.Value = vbUnchecked

End Sub

Private Sub guardardatos()
  Dim i, dnContador As Integer
  Dim pvRegistroActual, pvRegTotal As Variant
  Dim dvFeCre, dvFeMdf
  Dim sExpresion As String
  
  If uorstMain.RecordCount() > 0 Then
    uocnnMain.BeginTrans
  End If
  With uorstMain
    pvRegTotal = .RecordCount()
    For i = 1 To 12
      If pvRegTotal = 0 Then
        pvRegistroActual = 1
        .AddNew
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !mespvs = Format(i, "00")
        !ImpTCb_Cpr = Format(Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 4), FORMATO_NUM_2)
        !imptcb_vta = Format(Round(CDbl(IIf(txtDato(i + 11).Text = "", 0, txtDato(i + 11).Text)), 4), FORMATO_NUM_2)
        !impfac_hpr = Format(Round(CDbl(IIf(txtDato(i + 23).Text = "", 0, txtDato(i + 23).Text)), 4), FORMATO_NUM_2)
        !UsrCre = gsAbvUsr
        !FyHCre = Now
      Else
        .MoveFirst
        .Find "MesPvs = '" & Format(i, "00") & "'"
        If Not .EOF And (!ImpTCb_Cpr <> Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 4) Or _
          !imptcb_vta <> Round(CDbl(IIf(txtDato(i + 11).Text = "", 0, txtDato(i + 11).Text)), 4) Or _
          !impfac_hpr <> Round(CDbl(IIf(txtDato(i + 23).Text = "", 0, txtDato(i + 23).Text)), 4)) Then
          !UsrMdf = gsAbvUsr
          !ImpTCb_Cpr = Format(Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 4), FORMATO_NUM_2)
          !imptcb_vta = Format(Round(CDbl(IIf(txtDato(i + 11).Text = "", 0, txtDato(i + 11).Text)), 4), FORMATO_NUM_2)
          !impfac_hpr = Format(Round(CDbl(IIf(txtDato(i + 23).Text = "", 0, txtDato(i + 23).Text)), 4), FORMATO_NUM_2)
        End If
        .Update
        ' Inserta registro masivos
        If (chkVerificar.Value = vbChecked And gsNvlUsr = NvlUsr_Adm) Then
          If pvRegTotal = 0 Then
            sExpresion = "INSERT INTO cotcbmes (codemp, pdoano, mespvs, imptcb_cpr, imptcb_vta, impfac_hpr, usrcre, fyhcre) "
            sExpresion = sExpresion & "SELECT codemp, '" & gsAnoAct & "' AS pdoano, '" & Format(i, "00") & "' AS mespvs, "
            sExpresion = sExpresion & Format(Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 4), FORMATO_NUM_2) & " AS imptcb_cpr, "
            sExpresion = sExpresion & Format(Round(CDbl(IIf(txtDato(i + 11).Text = "", 0, txtDato(i + 11).Text)), 4), FORMATO_NUM_2) & " AS imptcb_vta, "
            sExpresion = sExpresion & Format(Round(CDbl(IIf(txtDato(i + 23).Text = "", 0, txtDato(i + 23).Text)), 4), FORMATO_NUM_2) & " AS impfac_hpr, "
            sExpresion = sExpresion & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
            sExpresion = sExpresion & "FROM siscfg.sgpms pms "
            sExpresion = sExpresion & "WHERE pms.codusr='" & gsCodUsr & "' "
            sExpresion = sExpresion & "AND pms.codsis='" & gsCodSis & "' AND pms.codmdl='frmMTCbCie' "
            sExpresion = sExpresion & "AND pms.codemp<>'" & gsCodEmp & "' "
            sExpresion = sExpresion & "AND EXISTS(SELECT * FROM cotcbmes tcb WHERE tcb.codemp=pms.codemp "
            sExpresion = sExpresion & "AND pdoano='" & gsAnoAct & "' "
            sExpresion = sExpresion & "AND mespvs='" & Format(i, "00") & "') "
            sExpresion = sExpresion & "ORDER BY codemp"
            uocnnMain.Execute sExpresion
          ElseIf (pvRegTotal <> 0 And Format(i, "00") = gsMesAct) Then
            sExpresion = "UPDATE cotcbmes tcb, siscfg.sgpms pms SET "
            sExpresion = sExpresion & "tcb.imptcb_cpr=" & Format(Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 4), FORMATO_NUM_2) & ", "
            sExpresion = sExpresion & "tcb.imptcb_vta=" & Format(Round(CDbl(IIf(txtDato(i + 11).Text = "", 0, txtDato(i + 11).Text)), 4), FORMATO_NUM_2) & ", "
            sExpresion = sExpresion & "tcb.impfac_hpr=" & Format(Round(CDbl(IIf(txtDato(i + 23).Text = "", 0, txtDato(i + 23).Text)), 4), FORMATO_NUM_2) & ", "
            sExpresion = sExpresion & "tcb.usrmdf='" & gsAbvUsr & "', "
            sExpresion = sExpresion & "tcb.fyhmdf='" & Format(Now, s_FmtFeHoMysql_0) & "' "
            sExpresion = sExpresion & "WHERE pms.codusr='" & gsCodUsr & "' "
            sExpresion = sExpresion & "AND pms.codsis='" & gsCodSis & "' "
            sExpresion = sExpresion & "AND pms.codmdl='frmMTCbCie' "
            sExpresion = sExpresion & "AND pms.codemp<>'" & gsCodEmp & "' "
            sExpresion = sExpresion & "AND tcb.codemp=pms.codemp "
            sExpresion = sExpresion & "AND tcb.pdoano='" & gsAnoAct & "' "
            sExpresion = sExpresion & "AND tcb.mespvs='" & Format(i, "00") & "'"
            uocnnMain.Execute sExpresion
          End If
        End If
      End If
    Next i
  End With

End Sub
']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property

Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
   cmdDeshacer.Enabled = IIf(pbNuevo = True, False, True)
End Property
