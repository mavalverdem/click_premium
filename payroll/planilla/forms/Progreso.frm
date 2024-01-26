VERSION 5.00
Begin VB.Form Progreso 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Progreso"
   ClientHeight    =   540
   ClientLeft      =   3960
   ClientTop       =   3180
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   120
   End
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Width           =   6375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FFFFFF&
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
         Height          =   350
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   6105
         TabIndex        =   1
         Top             =   200
         Width           =   6135
         Begin VB.Label label 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   120
            TabIndex        =   2
            Top             =   75
            Width           =   5895
         End
      End
   End
End
Attribute VB_Name = "Progreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private unidad As Long
Private limite As Long
Private Progreso As Long
Private defProgBarHwnd  As Long
Private Detener As Boolean
Dim habilitarproceso As Boolean
Private Sub Form_Load()
    Dim i As Long
    habilitarproceso = True
    label.Caption = labelprogreso
    Picture1.DrawMode = 10
    Picture1.FillStyle = 0
    unidad = 1
    limite = 100
    Progreso = 1
    Timer1.Interval = IntervalodeTiempo
    'Timer1.Interval = 100
End Sub
Private Sub MostrarPorcentaje(limite As Long, Progreso As Long)
    Dim msg As String
    If Progreso <= limite Then
      If Progreso > Picture1.ScaleWidth Then
         Progreso = Picture1.ScaleWidth
      End If
      Picture1.Cls
      Picture1.ScaleWidth = limite
      msg = Format$(CLng((Progreso / Picture1.ScaleWidth) * 100)) + "%"
      Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(msg)) \ 2
      Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight(msg)) \ 2
      Picture1.Print msg
      Picture1.Line (0, 0)-(Progreso, Picture1.ScaleHeight), Picture1.ForeColor, BF
      DoEvents
    End If
End Sub
Private Sub Timer1_Timer()
   Progreso = Progreso + unidad
   MostrarPorcentaje limite, Progreso
   If habilitarproceso = True Then
   End If
   habilitarproceso = False
   If Progreso >= limite Then
     Timer1.Interval = 0
     'frame.Caption = "Completado....."
     Unload Me
   End If
End Sub

