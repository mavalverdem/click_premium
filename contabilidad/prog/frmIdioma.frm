VERSION 5.00
Begin VB.Form frmIdioma 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libros Oficiales"
   ClientHeight    =   2595
   ClientLeft      =   4305
   ClientTop       =   6195
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Inicio de sesión"
   Begin VB.CommandButton cmdIdioma 
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   1
      Left            =   1800
      Picture         =   "frmIdioma.frx":0000
      TabIndex        =   0
      Top             =   840
      Width           =   1250
   End
   Begin VB.CommandButton cmdIdioma 
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   2
      Left            =   3360
      Picture         =   "frmIdioma.frx":0624
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1250
   End
End
Attribute VB_Name = "frmIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdIdioma_Click(Index As Integer)
  gsIdioma = Index
  Unload Me
End Sub
Private Sub Form_Load()
  gsIdioma = NvlUsr_Adm
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If gsIdioma = NvlUsr_Adm Then End
End Sub
