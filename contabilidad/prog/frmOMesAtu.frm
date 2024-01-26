VERSION 5.00
Begin VB.Form frmOMesAtu 
   Caption         =   "[Entidad]"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1380
      ScaleHeight     =   690
      ScaleWidth      =   1575
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4260
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
         Picture         =   "frmOMesAtu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
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
         Picture         =   "frmOMesAtu.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   3795
      Left            =   780
      ScaleHeight     =   3735
      ScaleWidth      =   2775
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2835
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Diciembre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   12
         Top             =   3420
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Noviembre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   11
         Top             =   3120
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Octubre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   9
         Left            =   840
         TabIndex        =   10
         Top             =   2820
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Setiembre"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   8
         Left            =   840
         TabIndex        =   9
         Top             =   2520
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Agosto"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   8
         Top             =   2220
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Julio"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   7
         Top             =   1920
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Junio"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   6
         Top             =   1620
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Mayo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   5
         Top             =   1320
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Abril"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   4
         Top             =   1020
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Marzo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Febrero"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   420
         Width           =   1200
      End
      Begin VB.OptionButton optMesAtu 
         Caption         =   "Enero"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   120
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmOMesAtu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pocnnBDS As ADODB.Connection
Private porstCoCfg As ADODB.Recordset

Private psMesAtu As String

Private Sub Form_Load()
   Set pocnnBDS = New Connection
   Set porstCoCfg = New Recordset
   
   With pocnnBDS
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   
   With porstCoCfg
      .ActiveConnection = pocnnBDS
      .Source = "SELECT MesAtu "
      .Source = .Source & "FROM COCfg "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
      .Open
   End With
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(0, 0)
  For nElemento = 0 To 11
    If gsIdioma = NvlUsr_Sup Then
      optMesAtu(nElemento).Caption = Choose(nElemento + 1, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    Else
      optMesAtu(nElemento).Caption = Choose(nElemento + 1, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    End If
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, True, False, True, aLabel
 ']
   
   psMesAtu = porstCoCfg!MesAtu
   optMesAtu(Val(psMesAtu) - 1).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   porstCoCfg.Close
   pocnnBDS.Close
   Set porstCoCfg = Nothing
   Set pocnnBDS = Nothing
End Sub

Private Sub optMesAtu_Click(Index As Integer)
   psMesAtu = gfCeros(Str(Index), 2, 1, "0")
End Sub

Private Sub cmdGrabar_Click()
   porstCoCfg!MesAtu = psMesAtu
   porstCoCfg.Update
  
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

