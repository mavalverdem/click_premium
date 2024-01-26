VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form frmMSeg 
   Caption         =   "[Entidad]"
   ClientHeight    =   6525
   ClientLeft      =   2205
   ClientTop       =   1200
   ClientWidth     =   7425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   7425
   Visible         =   0   'False
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7425
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   7425
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
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
         Index           =   1
         Left            =   2655
         Picture         =   "frmmseg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
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
         Left            =   5850
         Picture         =   "frmmseg.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame frmCuadro 
      Caption         =   " Permisos por Opciones "
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
      Height          =   5865
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      Begin VB.CommandButton cmdLlaveAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6795
         Picture         =   "frmmseg.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   375
         Width           =   255
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
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin MSOutl.Outline outEmpresa 
         Height          =   4485
         Left            =   135
         TabIndex        =   2
         Top             =   1290
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   7911
         _StockProps     =   77
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
         MouseIcon       =   "frmmseg.frx":03F6
         PicturePlus     =   "frmmseg.frx":0412
         PictureMinus    =   "frmmseg.frx":050C
         PictureLeaf     =   "frmmseg.frx":0606
         PictureOpen     =   "frmmseg.frx":0B28
         PictureClosed   =   "frmmseg.frx":0C22
      End
      Begin MSOutl.Outline outOpciones 
         Height          =   4485
         Left            =   3675
         TabIndex        =   10
         Top             =   1290
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   7911
         _StockProps     =   77
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
         MouseIcon       =   "frmmseg.frx":1144
         PicturePlus     =   "frmmseg.frx":1160
         PictureMinus    =   "frmmseg.frx":125A
         PictureLeaf     =   "frmmseg.frx":1354
         PictureOpen     =   "frmmseg.frx":1876
         PictureClosed   =   "frmmseg.frx":1970
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   3675
         TabIndex        =   11
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblLlaveDeta 
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
         Height          =   315
         Index           =   0
         Left            =   3075
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
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
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   420
         Width           =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         X1              =   135
         X2              =   7035
         Y1              =   840
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmMSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public uocnnMain As ADODB.Connection
'[Propio del formulario.
Public uorstSGUsr As ADODB.Recordset
Private aEmpresa() As String, aOpciones() As String
Private pbValidada As Boolean
']
Public Sub cmdImprimir_Click(Index As Integer)
 '[Datos del formulario de impresión.  'Cambiar.
   frmLPms.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLPms.Show vbModal
 ']
End Sub
Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2
      txtLlave(Index).SetFocus
   End Select
   ppAyuBus AYULLA, Index
End Sub
Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
   If txtLlave(0).Text <> "" Then ppAyuDet AYULLA, 0
 ']
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub
Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   Set uocnnMain = New ADODB.Connection
   Set uorstSGUsr = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With uorstSGUsr
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodUsr, NomUsr "
      .Source = .Source & "FROM SGUsr"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Usuario :", "Empresas", "Opciones")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "User :", "Companies", "Options")
  Next nElemento
  frmCuadro.Caption = Choose(gsIdioma, "Permisos por Opciones", "Permissions by Options")
  CaptionBotones Me, False, False, False, False, False, False, False, True, False, False, False, False, True, aLabel
 '
 pbValidada = False
 txtLlave(0).Text = gsCodUsr
 CargaEmpOpc
 CargaEmpresaUsuario txtLlave(0).Text
End Sub
Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   uocnnMain.Close
   Set uorstSGUsr = Nothing
   Set uocnnMain = Nothing
End Sub
Private Sub cmdSalir_Click()
   Unload Me
End Sub
'[Código propio del formulario.
']
Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Usr_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
   Else
'      Select Case tnIndex
'      Case 0                           'Cambiar (añadir índices).
'         modAyuBus.Dro_Cod "Length(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
'         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'      End Select
   End If
End Sub
Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex                 'Cambiar.
      Case 0
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With uorstSGUsr
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodUsr='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !NomUsr
            End If
         End With
      End Select
   Else
'      Select Case tnIndex                 'Cambiar.
'      Case 0
'         If txtDato(tnIndex).Text = "" Then
'            lblDatoDeta(tnIndex).Caption = ""
'            Exit Function
'         End If
'         With frmTVtaGrd.uorstCODro
'            If .RecordCount > 0 Then .MoveFirst
'            .Find "CodDro='" & txtDato(tnIndex).Text & "'"
'            If .EOF Then
'               MsgBox TEXT_8006, vbExclamation
'               ppAyuDet = True
'            Else
'               lblDatoDeta(tnIndex).Caption = " " & !DetDro
'            End If
'         End With
'      End Select
   End If
End Function
Private Sub CargaEmpOpc()
Static nContador As Integer
Static rstConfigura As ADODB.Recordset
Static nOpciones As Integer

Set rstConfigura = New ADODB.Recordset

With rstConfigura
  .ActiveConnection = uocnnMain
  .Source = "SELECT CodEmp, RazEmp FROM TGEmp ORDER BY CodEmp, RazEmp"
  .CursorType = adOpenKeyset
  .LockType = adLockReadOnly
  .Open
  .MoveLast
  .MoveFirst
End With

ReDim aEmpresa(1)
outEmpresa.AddItem Choose(gsIdioma, "Empresas de Contabilidad", "Accounting Companies")
outEmpresa.Indent(0) = 0
outEmpresa.PictureType(0) = outOpen
aEmpresa(outEmpresa.ListCount - 1) = "empresa"
' Cargo las empresas
While Not rstConfigura.EOF
  outEmpresa.ListIndex = -1
  outEmpresa.AddItem rstConfigura!codemp & " - " & rstConfigura!RazEmp
  ' Redimensiono el array de empresas
  ReDim Preserve aEmpresa(UBound(aEmpresa, 1) + 1)
  aEmpresa(outEmpresa.ListCount - 1) = rstConfigura!codemp
  rstConfigura.MoveNext
  DoEvents
Wend
rstConfigura.Close

ReDim aOpciones(1)
outOpciones.AddItem Choose(gsIdioma, "Sistema de Contabilidad", "Accounting System")
outOpciones.Indent(0) = 0
outOpciones.PictureType(0) = outOpen
aOpciones(outOpciones.ListCount) = "sistema"
' Cargo todas la opciones del sistema en el Outline
For nContador = 1 To 5
    outOpciones.ListIndex = -1
    If gsIdioma = NvlUsr_Sup Then
      outOpciones.AddItem Choose(nContador, "Transacciones", "Reportes", "Procesos", "Tablas", "Utilitarios")
    Else
      outOpciones.AddItem Choose(nContador, "Transactions", "Reports", "Processes", "Tables", "Tools")
    End If
    nOpciones = outOpciones.ListCount - 1
    ' Redimensiono el array de las opciones
    ReDim Preserve aOpciones(UBound(aOpciones, 1) + 1)
    aOpciones(outOpciones.ListCount) = "modulo"
    ' Cargo opciones del sistema
    With rstConfigura
      If .State = adStateOpen Then .Close
      .Source = "SELECT CodMdl, Orden, " & Choose(gsIdioma, "DetMdl", "DetMdlx") & " AS DetMdl "
      .Source = .Source & "FROM SGMdl "
      .Source = .Source & "WHERE codsis='" & gsCodSis & "' "
      .Source = .Source & "AND Opcion='" & nContador & "' "
      .Source = .Source & "ORDER BY Orden, CodMdl"
      .Open
      .MoveLast
      .MoveFirst
    End With

    While Not rstConfigura.EOF
      ' Redimensiono el array de las opciones
      ReDim Preserve aOpciones(UBound(aOpciones, 1) + 1)
      aOpciones(outOpciones.ListCount) = rstConfigura!CodMdl
      outOpciones.ListIndex = nOpciones
      outOpciones.AddItem Mid$(rstConfigura!DetMdl, 6)
      rstConfigura.MoveNext
   Wend
Next nContador
rstConfigura.Close
Set rstConfigura = Nothing

' Visualizo y desactivo opciones
For nContador = 1 To outOpciones.ListCount - 1
  If outOpciones.Indent(nContador) = 1 Then
    If outOpciones.HasSubItems(nContador) Then
      outOpciones.PictureType(nContador) = outOpen
    Else
      outOpciones.RemoveItem nContador
    End If
  Else
    outOpciones.PictureType(nContador) = outClosed
  End If
Next nContador

End Sub

Private Sub CargaEmpresaUsuario(sUsuario As String)
 
Static rstConfigura As ADODB.Recordset
Static nContador As Integer, nEmpresa As Integer

Set rstConfigura = New ADODB.Recordset
With rstConfigura
  .ActiveConnection = uocnnMain
  .Source = "SELECT DISTINCT CodEmp "
  .Source = .Source & "FROM SGPms "
  .Source = .Source & "WHERE CodSis='" & gsCodSis & "' "
  .Source = .Source & "AND CodUsr='" & sUsuario & "' "
  .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodEmp, '')<>'' "
  .Source = .Source & "ORDER BY CodEmp"
  .CursorType = adOpenKeyset
  .LockType = adLockReadOnly
  .Open
End With
    
nEmpresa = 0
' Cargo todas las empresas contables por el usuario
For nContador = 1 To outEmpresa.ListCount - 1
  If Not (rstConfigura.BOF And rstConfigura.EOF) Then rstConfigura.MoveFirst
  rstConfigura.Find "CodEmp='" & aEmpresa(nContador) & "'"
  If Not rstConfigura.EOF Then
    outEmpresa.PictureType(nContador) = outLeaf
    nEmpresa = IIf(nEmpresa > 0, nEmpresa, nContador)
  Else
    outEmpresa.PictureType(nContador) = outClosed
  End If
Next nContador
outEmpresa.ListIndex = nEmpresa
rstConfigura.Close
Set rstConfigura = Nothing
CargaOpcionesUsuario aEmpresa(outEmpresa.ListIndex), sUsuario
 
End Sub

Private Sub CargaOpcionesUsuario(sEmpresa As String, sUsuario As String)
 
Static rstConfigura As ADODB.Recordset
Static nContador As Integer

Set rstConfigura = New ADODB.Recordset
With rstConfigura
  .ActiveConnection = uocnnMain
  .Source = "SELECT DISTINCT CodMdl FROM SGPms "
  .Source = .Source & "WHERE CodSis='" & gsCodSis & "' "
  .Source = .Source & "AND CodUsr='" & sUsuario & "' "
  .Source = .Source & "AND CodEmp='" & sEmpresa & "' "
  .Source = .Source & "ORDER BY CodMdl"
  .CursorType = adOpenKeyset
  .LockType = adLockReadOnly
  .Open
End With
    
' Cargo todas las opciones del sistema de contabilidad por el empresa, usuario
For nContador = 1 To outOpciones.ListCount - 1
  If outOpciones.Indent(nContador) = 2 Then
    If Not (rstConfigura.EOF And rstConfigura.BOF) Then rstConfigura.MoveFirst
    rstConfigura.Find "CodMdl='" & aOpciones(nContador) & "'"
    If Not rstConfigura.EOF Then
      outOpciones.PictureType(nContador) = outLeaf
    Else
      outOpciones.PictureType(nContador) = outClosed
    End If
  End If
Next nContador
outOpciones.ListIndex = 1
rstConfigura.Close
Set rstConfigura = Nothing
 
End Sub

Private Sub outEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
Static nListIndex As Integer

If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
  nListIndex = outEmpresa.ListIndex
  If (nListIndex > 0 And nListIndex < outEmpresa.ListCount - 1) Then
      If KeyCode = vbKeyDown Then nListIndex = nListIndex + 1 Else nListIndex = nListIndex - 1
  End If
  CargaOpcionesUsuario aEmpresa(nListIndex), txtLlave(0).Text
End If

End Sub

Private Sub outEmpresa_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
  CargaOpcionesUsuario aEmpresa(outEmpresa.ListIndex), txtLlave(0).Text
End If

End Sub

Private Sub outEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
  CargaOpcionesUsuario aEmpresa(outEmpresa.ListIndex), txtLlave(0).Text
End If

End Sub

Private Sub outEmpresa_PictureDblClick(ListIndex As Integer)
Static sSentencia As String
   
If outEmpresa.Indent(ListIndex) = 1 Then
  ' Ubico en el indice actualizar
  outEmpresa.ListIndex = ListIndex
  ' Asigno o Desasigno la Opcion
  uocnnMain.BeginTrans            'INICIA TRANSACCION.
  If outEmpresa.PictureType(ListIndex) = outClosed Then
    ' Elimino primero las opciones
    sSentencia = "DELETE FROM SGPms WHERE codsis='" & gsCodSis & "' AND CodUsr='" & txtLlave(0).Text & "' AND CodEmp='" & aEmpresa(ListIndex) & "'"
    uocnnMain.Execute sSentencia
    ' Agrego las opciones
    outEmpresa.PictureType(ListIndex) = outLeaf
    sSentencia = "INSERT INTO SGPms(CodUsr, CodEmp, CodMdl, CodSis, IndPms01, IndPms02, IndPms03, IndPms04, IndPms05, "
    sSentencia = sSentencia & "IndPms06, IndPms07, IndPms08, IndPms09, IndPms10, UsrCre, FyHCre, UsrMdf, FyHMdf) "
    sSentencia = sSentencia & "SELECT '" & txtLlave(0).Text & "', '" & aEmpresa(ListIndex) & "', a.CodMdl, "
    sSentencia = sSentencia & "a.codsis, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, '" & gsAbvUsr & "', "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & ", "
    sSentencia = sSentencia & "Null, Null "
    sSentencia = sSentencia & "FROM SGMdl a "
    sSentencia = sSentencia & "WHERE a.codsis='" & gsCodSis & "' "
    sSentencia = sSentencia & "ORDER BY a.Opcion, a.Orden, a.CodMdl"
  Else
    outEmpresa.PictureType(ListIndex) = outClosed
    sSentencia = "DELETE FROM SGPms WHERE codsis='" & gsCodSis & "' AND CodUsr='" & txtLlave(0).Text & "' AND CodEmp='" & aEmpresa(ListIndex) & "'"
  End If
  uocnnMain.Execute sSentencia
  uocnnMain.CommitTrans
  ' Refresco el control
  CargaOpcionesUsuario aEmpresa(ListIndex), txtLlave(0).Text
End If

End Sub

Private Sub outOpciones_PictureClick(ListIndex As Integer)
Static sSentencia As String
   
If outOpciones.Indent(ListIndex) > 1 Then
  ' Asigno o Desasigno la Opcion
  If outOpciones.PictureType(ListIndex) = outClosed Then
    outOpciones.PictureType(ListIndex) = outLeaf
    sSentencia = "INSERT INTO SGpms(CodUsr, CodEmp, CodMdl, CodSis, IndPms01, IndPms02, IndPms03, IndPms04, IndPms05, "
    sSentencia = sSentencia & "IndPms06, IndPms07, IndPms08, IndPms09, IndPms10, UsrCre, FyHCre, UsrMdf, FyHMdf) "
    sSentencia = sSentencia & "VALUES('" & txtLlave(0).Text & "', '" & aEmpresa(outEmpresa.ListIndex) & "', '" & aOpciones(ListIndex) & "', '"
    sSentencia = sSentencia & gsCodSis & "', 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, '" & gsAbvUsr & "', "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & ", "
    sSentencia = sSentencia & "Null, Null)"
  Else
    outOpciones.PictureType(ListIndex) = outClosed
    sSentencia = "DELETE FROM SGPms WHERE codsis='" & gsCodSis & "' AND CodUsr='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND CodEmp='" & aEmpresa(outEmpresa.ListIndex) & "' AND CodMdl='" & aOpciones(ListIndex) & "'"
  End If
  uocnnMain.BeginTrans            'INICIA TRANSACCION.
  uocnnMain.Execute sSentencia
  uocnnMain.CommitTrans
End If
End Sub
Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub
Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYULLA, Index
   End If
End Sub
Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtLlave(Index).SetFocus 'Cambiar.
End Sub
Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
   Case 0
      Cancel = ppAyuDet(AYULLA, Index)
      If Cancel Then Exit Sub
   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 Then
      With uorstSGUsr
         If Not (.BOF And .EOF) Then
            .MoveFirst
            .Find "CodUsr='" & txtLlave(0).Text & "'"
            If .EOF Then
               MsgBox Choose(gsIdioma, "Dato no Existe, Verificar", "Data doesn't Exist, Verify"), vbExclamation
            End If
         End If
      End With
      pbValidada = False
      ' Cargo las empresas del usuario
      CargaEmpresaUsuario txtLlave(0).Text
   Else
      MsgBox Choose(gsIdioma, "Dato no Existe, Verificar", "Data doesn't Exist, Verify"), vbExclamation
      pbValidada = True
   End If
End Sub
