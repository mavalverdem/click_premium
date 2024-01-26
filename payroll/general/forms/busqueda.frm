VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3210
   ClientLeft      =   3675
   ClientTop       =   3465
   ClientWidth     =   5340
   Icon            =   "busqueda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5340
      _Version        =   65536
      _ExtentX        =   9419
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   4455
         TabIndex        =   7
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "busqueda.frx":000C
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4200
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   2685
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   5310
      _Version        =   65536
      _ExtentX        =   9366
      _ExtentY        =   4736
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.TextBox txtValor 
         Height          =   300
         Left            =   2100
         TabIndex        =   1
         Top             =   495
         Width           =   3015
      End
      Begin VB.ListBox lstCampos 
         Height          =   2010
         ItemData        =   "busqueda.frx":0028
         Left            =   255
         List            =   "busqueda.frx":002A
         TabIndex        =   4
         Top             =   510
         Width           =   1755
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   1530
         Index           =   0
         Left            =   2715
         TabIndex        =   2
         Top             =   945
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2963
         _ExtentY        =   2699
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "busqueda.frx":002C
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Valor :"
         Height          =   195
         Index           =   1
         Left            =   2115
         TabIndex        =   0
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Campos :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   660
      End
   End
End
Attribute VB_Name = "fBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                     ' Declarar variable antes de usarla

Private s_TitleWindow As String                     ' Titulo del Formulario
Private i As Integer                                ' Contador temporal para los bucles
Private Sub BuscaCriterio(rs_Datos As ADODB.Recordset, tb As TDBGrid)
  Dim s_Condicion As String, nPosicion As Integer
  
  ' Capturo el numero de columna del campo seleccionado
  For i = 0 To tb.Columns.Count - 1
    If tb.Columns(i).Caption = lstCampos.List(lstCampos.ListIndex) Then Exit For
  Next i

  s_Condicion = tb.Columns(i).DataField & " LIKE '" & gdl_Funcion.aTexto(txtValor) & "*'"
  ' Verifico que el Dato Ingresado sea Correcto
  Select Case rs_Datos(tb.Columns(i).DataField).Type
   Case adDBTimeStamp
      If Not IsDate(Trim$(txtValor)) Then Beep: MsgBox "Valor de Búsqueda Erróneo", vbExclamation: Exit Sub
      s_Condicion = tb.Columns(i).DataField & " LIKE #" & gdl_Funcion.aTexto(txtValor) & "#"
   Case adSmallInt, adInteger, adSingle, adDouble, adBinary, adNumeric, adUnsignedInt, adUnsignedBigInt, adUnsignedSmallInt
      If Not IsNumeric(Trim$(txtValor)) Then Beep: MsgBox "Valor de Búsqueda Erróneo", vbExclamation: Exit Sub
      s_Condicion = tb.Columns(i).DataField & " LIKE " & gdl_Funcion.aTexto(txtValor)
  End Select
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Busco el Criterio Seleccionado
  nPosicion = IIf(rs_Datos.Bookmark >= 1 And rs_Datos.Bookmark < rs_Datos.RecordCount, adSearchForward, adSearchBackward)
  rs_Datos.Find s_Condicion, 1, nPosicion
  ' Verifico si la Busqueda fue exitosa
  If Not rs_Datos.EOF Then
    If (rs_Datos.Bookmark = 1 Or rs_Datos.Bookmark = rs_Datos.RecordCount) And (Mid(UCase(rs_Datos(tb.Columns(i).DataField)), 1, Len(txtValor)) <> UCase(txtValor)) Then
      Beep
      MsgBox "No se Encontró " & lblTitle & " !!!", vbInformation
    End If
  Else
    rs_Datos.MoveFirst
  End If
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal
    
End Sub
Private Sub cmdAction_Click(Index As Integer)

  If Index = 0 Then
    If lstCampos.SelCount = 0 Then
      Beep
      MsgBox "Debe Seleccionar Campo", vbExclamation: Exit Sub
    ElseIf txtValor = "" Then
      Beep
      MsgBox "Ingrese Valor a Buscar", vbExclamation: Exit Sub
    End If
    ' Busco el Criterio Ingresado
    BuscaCriterio go_dcaBusqueda.Recordset, go_tdbBusqueda
  Else
    ' Cierro el formulario
    Unload Me
  End If

End Sub
Private Sub Form_Load()

  s_TitleWindow = "Búsqueda"
  lblTitle = Trim(Mid$(go_dcaBusqueda.Caption, InStrRev(go_dcaBusqueda.Caption, " ", -1, vbTextCompare)))
  
  ' Configuro la Visualización del Formulario, Controles del ToolBar
  ReDim aElemento(2, 2)
  ' Icono y título de Formulario
  aElemento(2, 1) = "busqueda": aElemento(2, 2) = s_TitleWindow
  
  ' Cargo los Graficos a los Controles
  For i = 0 To 1
      aElemento(i, 1) = Choose(i + 1, "procebus", "cancelar")
      aElemento(i, 2) = Choose(i + 1, "Realizar Busqueda", "Cancelar Busqueda")
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  For i = 0 To gn_ColBusqueda - 1
      If go_tdbBusqueda.Columns(i).Visible Then lstCampos.AddItem go_tdbBusqueda.Columns(i).Caption
  Next i
  
  ' Posiciono el cursor
  Me.Tag = 1: lstCampos.ListIndex = 0: Me.Tag = 0
  cmdAction(1).Cancel = True
   
End Sub
Private Sub lstCampos_Click()
  ' Paso el foco al Ingreso del Valor
  If Me.Tag = 0 Then txtValor.SetFocus
End Sub
Private Sub txtValor_GotFocus()
  gdl_Procedure.MarcaGet txtValor
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click (0)
End Sub
