VERSION 5.00
Begin VB.Form fInicio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "inicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTiempo 
      Interval        =   1
      Left            =   4230
      Top             =   570
   End
   Begin VB.Label lblAdvertencia 
      BackStyle       =   0  'Transparent
      Caption         =   $"inicio.frx":000C
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   885
      Left            =   240
      TabIndex        =   5
      Top             =   4485
      Width           =   4485
   End
   Begin VB.Label lblAutoriza 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autoriza el uso de este Producto a "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4410
      TabIndex        =   4
      Top             =   195
      Width           =   3105
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© 1998-2008 System Corporation"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4110
      TabIndex        =   3
      Top             =   3900
      Width           =   3015
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 2.0"
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
      Height          =   195
      Left            =   4110
      TabIndex        =   2
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label lblPlataforma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Servidor para Windows"
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
      Height          =   195
      Left            =   4110
      TabIndex        =   1
      Top             =   3390
      Width           =   2625
   End
   Begin VB.Label lblSoftware 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Personal y Planillas"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   1110
      TabIndex        =   0
      Top             =   2055
      Width           =   5415
   End
End
Attribute VB_Name = "fInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private n_Index As Integer                  ' Contador del bucle de tiempo
Private s_Archivo As String                 ' Nombre de archivo de imagen, icono u otro
Private Sub Form_Click()
n_Index = 100
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
n_Index = 100
End Sub
Private Sub Form_Load()

n_Index = 0
' Verifico que exista el Logo del Sistema
Me.Picture = LoadPicture()
s_Archivo = gdl_Procedure.ps_PathImagen & "logo sysmavm.jpg"
If dir$(s_Archivo, vbNormal) <> "" Then
    Me.Picture = LoadPicture(s_Archivo)
End If

ps_Licencia = " MLV Contadores S.A.C."
lblAutoriza = "Autoriza el uso de este Producto a " & ps_Licencia
lblSoftware = ps_NomSistema
Me.Refresh

End Sub
Private Sub tmrTiempo_Timer()

n_Index = n_Index + 1
If n_Index > 23 Then
    tmrTiempo.Interval = 0
    Unload Me
End If

End Sub

