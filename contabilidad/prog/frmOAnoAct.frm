VERSION 5.00
Begin VB.Form frmOAnoAct 
   Caption         =   "[Entidad]"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstAnoAct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "frmOAnoAct.frx":0000
      Left            =   1740
      List            =   "frmOAnoAct.frx":0097
      TabIndex        =   1
      Top             =   840
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1380
      ScaleHeight     =   690
      ScaleWidth      =   1815
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1815
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
         Left            =   800
         Picture         =   "frmOAnoAct.frx":01C1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
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
         Left            =   60
         Picture         =   "frmOAnoAct.frx":030B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   780
      ScaleHeight     =   2115
      ScaleWidth      =   2775
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2835
   End
End
Attribute VB_Name = "frmOAnoAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private psAnoAct As String

Private Sub Form_Load()
   psAnoAct = gsAnoAct

   lstAnoAct.ListIndex = (lstAnoAct.ListCount - 1) - (Val(psAnoAct) - 2002)
End Sub

Private Sub lstAnoAct_Click()
   psAnoAct = lstAnoAct.Text
End Sub

Private Sub cmdAceptar_Click()
   gsAnoAct = psAnoAct
   frmMain.lblVar(2) = gsAnoAct
   gsRutBDC = Left(gsRutBDC, Len(gsRutBDC) - 5) & gsAnoAct & "\"
   gsRutBDS = Left(gsRutBDS, Len(gsRutBDC) - 5) & gsAnoAct & "\"
  
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

