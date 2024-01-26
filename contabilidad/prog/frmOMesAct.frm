VERSION 5.00
Begin VB.Form frmOMesAct 
   Caption         =   "[Entidad]"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1403
      ScaleHeight     =   690
      ScaleWidth      =   1575
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5100
      Width           =   1575
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
         Picture         =   "frmOMesAct.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmOMesAct.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   4635
      Left            =   780
      ScaleHeight     =   4575
      ScaleWidth      =   2775
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   2835
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Cierre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   13
         Left            =   840
         TabIndex        =   13
         Top             =   4140
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Apertura"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Diciembre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   12
         Left            =   840
         TabIndex        =   12
         Top             =   3840
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Noviembre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   11
         Top             =   3540
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Octubre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   10
         Top             =   3240
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Setiembre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   9
         Left            =   840
         TabIndex        =   9
         Top             =   2940
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "A&gosto"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   8
         Left            =   840
         TabIndex        =   8
         Top             =   2640
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "J&ulio"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   7
         Top             =   2340
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Junio"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   6
         Top             =   2040
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Mayo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   5
         Top             =   1740
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "A&bril"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Marzo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   3
         Top             =   1140
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Febrero"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   1200
      End
      Begin VB.OptionButton optMesAct 
         Caption         =   "&Enero"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   1
         Top             =   540
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmOMesAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private psMesAct As String

Private Sub Form_Load()
   psMesAct = gsMesAct
  
  '[ Cargo los mensajes de botones
  Dim nIndex As Integer
  ReDim aLabel(0, 0)
  For nIndex = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      optMesAct(nIndex).Caption = Choose(nIndex + 1, "&Apertura", "&Enero", "&Febrero", "&Marzo", "A&bril", "Ma&yo", "&Junio", "J&ulio", "A&gosto", "&Setiembre", "&Octubre", "&Noviembre", "&Diciembre", "&Cierre")
    Else
      optMesAct(nIndex).Caption = Choose(nIndex + 1, "&Opening", "&January", "&February", "&March", "A&pril", "Ma&y", "J&une", "Ju&ly", "Au&gust", "&September", "Oc&tober", "&November", "&December", "&Closing")
    End If
  Next nIndex
  CaptionBotones Me, True, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
   
   optMesAct(Val(psMesAct)).Value = True
End Sub

Private Sub optMesAct_Click(Index As Integer)
   psMesAct = gfCeros(Str(Index), 2, 0, "0")
End Sub

Private Sub cmdAceptar_Click()
   gsMesAct = psMesAct
   
   frmMain.lblVar(3) = gsMesAct
'[Propio del Proyecto.
   gpCamposSaldos
   gpCieMes
']
  
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

