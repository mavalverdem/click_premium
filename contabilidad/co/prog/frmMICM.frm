VERSION 5.00
Begin VB.Form frmMICM 
   Caption         =   "[Entidad]"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   4290
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   3735
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   240
      Width           =   3795
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
         MaxLength       =   14
         TabIndex        =   11
         Top             =   4440
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   10
         Top             =   4080
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   9
         Top             =   3720
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   8
         Top             =   3360
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   1
         Top             =   840
         Width           =   1575
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
         MaxLength       =   14
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   540
         TabIndex        =   18
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Left            =   540
         TabIndex        =   20
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label5 
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
         Left            =   540
         TabIndex        =   19
         Top             =   900
         Width           =   570
      End
      Begin VB.Label Label1 
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
         Left            =   540
         TabIndex        =   22
         Top             =   1980
         Width           =   390
      End
      Begin VB.Label Label2 
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
         Left            =   540
         TabIndex        =   23
         Top             =   2340
         Width           =   375
      End
      Begin VB.Label Label6 
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
         Left            =   540
         TabIndex        =   21
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label Label7 
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
         Left            =   540
         TabIndex        =   25
         Top             =   3060
         Width           =   525
      End
      Begin VB.Label Label8 
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
         Left            =   540
         TabIndex        =   26
         Top             =   3420
         Width           =   720
      End
      Begin VB.Label Label9 
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
         Left            =   540
         TabIndex        =   24
         Top             =   2700
         Width           =   315
      End
      Begin VB.Label Label10 
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
         Left            =   540
         TabIndex        =   28
         Top             =   4140
         Width           =   765
      End
      Begin VB.Label Label11 
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
         Left            =   540
         TabIndex        =   29
         Top             =   4500
         Width           =   705
      End
      Begin VB.Label Label12 
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
         Left            =   540
         TabIndex        =   27
         Top             =   3780
         Width           =   585
      End
      Begin VB.Label Label16 
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
         Left            =   540
         TabIndex        =   17
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Indice"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   720
      ScaleHeight     =   690
      ScaleWidth      =   2895
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5640
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
         Picture         =   "frmMICM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmMICM.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmMICM.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmMICM.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   700
      End
   End
End
Attribute VB_Name = "frmMICM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As adodb.Connection
Public uorstMain As adodb.Recordset
'Public dvValTCb As String
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer
Private pbNuevo As Boolean

Private Sub Form_Load()
Dim dnContador As Integer
   Me.KeyPreview = True
   pbNuevo = True
   psConnStrgSele = "SELECT MesICM, ImpInd," _
                  & "  UsrCre, FyHCre, UsrMdf, FyHMdf " _
                  & "FROM CoICM "
   psConnStrgOrde = "ORDER BY 1"
   
   Set uocnnMain = New adodb.Connection
   Set uorstMain = New adodb.Recordset
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
      .Properties("Unique Table").Value = "CoICM"
   End With

   mostrardatos
   
   If pbNuevo Then
'      cmdRetroceder.Enabled = False
'      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False

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

'Private Sub cmdRetroceder_Click()
'   Dim dnContador As Integer
''   gpTUe_Retroceder frmMCCoGrd.uorstMain, Me 'Cambiar Formulario de Grid.
'   With TxtDato
'      For dnContador = 0 To .Count - 1
'         If TxtDato(0).SetFocus Then
'            Exit Sub
'         End If
'         If TxtDato(dnContador).SetFocus Then
'            TxtDato(dnContador - 1).SetFocus
'            Exit Sub
'         End If
'      Next
'   End With
'End Sub

'Private Sub cmdAvanzar_Click()
'   Dim dnContador As Integer
''   gpTUe_Avanzar frmMCCoGrd.uorstMain, Me 'Cambiar Formulario de Grid.
'   With TxtDato
'      For dnContador = 0 To .Count - 1
'         If TxtDato(25).SetFocus Then
'            Exit Sub
'         End If
'         If TxtDato(dnContador).SetFocus Then
''            TxtDato(dnContador + 1).SetFocus
''            Exit Sub
''         End If
''      Next
''   End With
'End Sub

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
  'uocnnMain.BeginTrans  'INICIA TRANSACCION.
  uorstMain.CancelUpdate
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
   Dim var As Long
   Dim ent As Long
   var = Len(txtDato.Item(Index).Text)
   ent = Int(CDbl(Val(txtDato.Item(Index).Text)))
'   If var > 6 Then
'      Cancel = True
'      MsgBox "Solo acepta 6 digitos", vbOKOnly, "Advertencia"
'      txtDato.Item(Index).SetFocus
'   Else
   If ent > 9999999 Then
      Cancel = True
      MsgBox Choose(gsIdioma, "Solo acepta 7 enteros", "Only accept 7 integers"), vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
      txtDato.Item(Index).SetFocus
   Else
      If CDec(Val(txtDato.Item(Index).Text)) > 999999 Then
         Cancel = True
         MsgBox Choose(gsIdioma, "Solo acepta 6 decimales", "Only accept 6 decimals"), vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
         txtDato.Item(Index).SetFocus
      Else
         txtDato.Item(Index).Text = Format(Round(CDbl(Val(txtDato.Item(Index).Text)), 6), FORMATO_NUM_3)
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

  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
End Sub

'[Propios del formulario
Private Sub mostrardatos()
   Dim i As Integer
   Dim v As String
   Dim pvRegistroActual As Variant
   With uorstMain
      If Not .EOF Then
         For i = 1 To 12
            If i < 10 Then
               v = "0" + Trim(Str(i))
            Else
               v = Trim(Str(i))
            End If
            pvRegistroActual = .Bookmark
            .MoveFirst
            .Find "MesICM = '" & v & "'"
            If .EOF Then
               'RSet dvValTCb = Str(Round(0, 3))
               'txtDato(i - 1).Text = dvValTCb
               'txtDato(i + 11).Text = dvValTCb
               txtDato(i - 1).Text = Format(Round(0, 6), FORMATO_NUM_3)
               .Bookmark = pvRegistroActual
            Else
               'RSet dvValTCb = Str(Round(CDbl(uorstMain!ImpTCb_Cpr), 3))
               'txtDato(i - 1) = dvValTCb
               'txtDato(i + 11) = dvValTCb
               txtDato(i - 1).Text = Format(Round(CDec(uorstMain!ImpInd), 6), FORMATO_NUM_3)
               .Bookmark = pvRegistroActual
            End If
            'txtDato(i - 1).SelAlignment = 1
            'txtDato(i + 11).SelAlignment = 1
         Next i
      Else
         For i = 1 To 12
            txtDato(i - 1).Text = Format(Round(0, 6), FORMATO_NUM_3)
         Next i
         uocnnMain.BeginTrans
      End If
   End With
End Sub

Private Sub guardardatos()
   Dim i, dnContador As Integer
   Dim v As String
   Dim pvRegistroActual, pvRegTotal As Variant
   Dim dvFeCre, dvFeMdf
   If uorstMain.RecordCount() > 0 Then
      uocnnMain.BeginTrans
   End If
   With uorstMain
      pvRegTotal = .RecordCount()
      For i = 1 To 12
         If i < 10 Then
            v = "0" + Trim(Str(i))
         Else
            v = Trim(Str(i))
         End If
         If pvRegTotal = 0 Then
            pvRegistroActual = 1
            .AddNew
            !MesICM = v
            !ImpInd = Format(Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 6), FORMATO_NUM_3)
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            .MoveFirst
            .Find "MesICM = '" & v & "'"
            If Not .EOF And (!ImpInd <> Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 6)) Then
               !UsrMdf = gsAbvUsr
               !ImpInd = Format(Round(CDbl(IIf(txtDato(i - 1).Text = "", 0, txtDato(i - 1).Text)), 6), FORMATO_NUM_3)
            End If
            .Update
         End If
      Next i
'      .Update
'      .Bookmark = pvRegistroActual
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

