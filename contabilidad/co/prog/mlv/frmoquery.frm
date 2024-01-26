VERSION 5.00
Begin VB.Form frmOQuery 
   Caption         =   "[Entidad]"
   ClientHeight    =   4050
   ClientLeft      =   165
   ClientTop       =   1350
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame fraUbicacion 
      Caption         =   " Carpeta "
      ForeColor       =   &H00800000&
      Height          =   3330
      Index           =   0
      Left            =   4905
      TabIndex        =   16
      Top             =   90
      Width           =   2535
      Begin VB.DirListBox dlbDirectorio 
         Height          =   1440
         Left            =   150
         TabIndex        =   19
         Top             =   690
         Width           =   2235
      End
      Begin VB.FileListBox flbArchivo 
         Height          =   870
         Left            =   150
         Pattern         =   "*.txt"
         TabIndex        =   21
         Top             =   2355
         Width           =   2235
      End
      Begin VB.DriveListBox drvUnidad 
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   400
         Width           =   2235
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Directorio :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   200
         Width           =   765
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Archivos :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   20
         Top             =   2150
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   350
      Left            =   3855
      TabIndex        =   0
      Top             =   3615
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   350
      Left            =   2175
      TabIndex        =   1
      Top             =   3615
      Width           =   1215
   End
   Begin VB.Frame fraParametro 
      Caption         =   " Parámetro "
      ForeColor       =   &H00800000&
      Height          =   2160
      Left            =   165
      TabIndex        =   5
      Top             =   1260
      Width           =   4650
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   1
         Left            =   780
         TabIndex        =   10
         Top             =   645
         Width           =   435
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   0
         Left            =   780
         TabIndex        =   7
         Top             =   300
         Width           =   435
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   4290
         Picture         =   "frmoquery.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   645
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   4290
         Picture         =   "frmoquery.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   255
      End
      Begin VB.ComboBox cmbEjercicio 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   1245
      End
      Begin VB.ComboBox cmbEjercicio 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label lblDatoDeta 
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   645
         Width           =   3090
      End
      Begin VB.Label lblDatoDeta 
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   300
         Width           =   3090
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         Caption         =   "Inicio Ejercicio"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   885
         TabIndex        =   12
         Top             =   1170
         Width           =   1350
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
         Caption         =   "Fin Ejercicio"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   2490
         TabIndex        =   14
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   690
         Width           =   675
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   345
         Width           =   675
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   " Opción "
      ForeColor       =   &H00800000&
      Height          =   945
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   1680
      Begin VB.OptionButton optParametro 
         Caption         =   "Configuración"
         ForeColor       =   &H80000002&
         Height          =   250
         Index           =   1
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   570
         Width           =   1410
      End
      Begin VB.OptionButton optParametro 
         Caption         =   "Contabilidad"
         ForeColor       =   &H8000000D&
         Height          =   250
         Index           =   0
         Left            =   105
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmOQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pocnnMain As ADODB.Connection
Private porstTGEMP As ADODB.Recordset

Private Sub cmdAceptar_Click()
   Dim pofsoQuery As FileSystemObject
   Dim potxsQuery As TextStream
   Dim psLinea As String, psCadena As String
   Dim psArchivo As String, psCatalogo As String
   Dim pnContador As Integer
   Dim pocnQuery As ADODB.Connection
   Dim porsClone As ADODB.Recordset

   On Error GoTo Err

  If Not flbArchivo.FileName <> "" Then MsgBox Choose(gsIdioma, "Seleccione archivo a procesar; Verificar", "You select file to process; Verify"), vbExclamation: flbArchivo.SetFocus: Exit Sub
  If optParametro(0).Value Then
    If Not (Right(cmbEjercicio(1).Text, 4) >= Right(cmbEjercicio(0).Text, 4)) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: cmbEjercicio(1).SetFocus: Exit Sub
  End If
  
  If MsgBox(Choose(gsIdioma, "Desea procesar la actualización de información ?", "Do you want to process the update to information?"), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    Set pocnQuery = New ADODB.Connection
    Set porsClone = New ADODB.Recordset
    Set porsClone = porstTGEMP.Clone()
    
    ' Aperturo el archivo query
    psArchivo = flbArchivo.Path & "\" & flbArchivo.FileName
    Set pofsoQuery = New FileSystemObject
    Set potxsQuery = pofsoQuery.OpenTextFile(psArchivo, ForReading, False, TristateFalse)
    Do While Not potxsQuery.AtEndOfStream
      psLinea = potxsQuery.ReadLine
      If Left(psLinea, 1) <> "#" Then
        psCadena = psCadena & psLinea
        If Right(Trim(psLinea), 1) = ";" Then
          ' Ejecuto la sentencia
          If optParametro(0).Value Then
            ' Ubico en el registro deseado de empresas
'            porsClone.MoveFirst
'            porsClone.Find "CodEmp='" & txtDato(0).Text & "'"
          
'            Do While Not porsClone.EOF
'              If Not (porsClone!codemp > txtDato(1).Text) Then
'                For pnContador = Val(Right(cmbEjercicio(0).Text, 4)) To Val(Right(cmbEjercicio(1).Text, 4))
'                  psCatalogo = "c" & Trim$(porsClone!codemp) & Trim$(pnContador)
                  psCatalogo = gsNomBDS
                  If ValidadConexion(psCatalogo) Then
                    With pocnQuery
                      If .State = adStateOpen Then .Close
                      .ConnectionTimeout = 15
                      .CursorLocation = adUseClient
                      .ConnectionString = CONNSTRG & psCatalogo
                      .Open
                    End With
                    pocnQuery.BeginTrans                  'INICIA TRANSACCION.
                    pocnQuery.Execute psCadena
                    pocnQuery.CommitTrans                 'CONFIRMA TRANSACCION.
                  End If
Contabilidad:
'                Next pnContador
'              End If
'              porsClone.MoveNext
'            Loop
          Else
            pocnnMain.BeginTrans                  'INICIA TRANSACCION.
            pocnnMain.Execute psCadena
            pocnnMain.CommitTrans                 'CONFIRMA TRANSACCION.
          End If
Configuracion:
          psCadena = ""
        End If
      End If
    Loop
    potxsQuery.Close
    Set pofsoQuery = Nothing
    Set potxsQuery = Nothing
    
    MsgBox TEXT_8008, vbInformation
    ' Elimino el archivo de secuencias
    If Dir$(psArchivo, vbNormal) <> "" Then Kill psArchivo
    flbArchivo.Refresh
    cmdAceptar.Enabled = True
  End If
  GoTo Fin

Err:
  If Err.Number = -2147217900 Then
    If optParametro(0).Value Then
      Resume cnnQuery
    Else
      Resume cnnMain
    End If
  End If
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  GoTo Fin

cnnMain:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  If MsgBox(Choose(gsIdioma, "¡Información actualizada!", "Up-to-date Information") & Chr(13) & Choose(gsIdioma, " Continua actualización de información ?", " Continue update of information?"), vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    GoTo Configuracion
  End If
  GoTo Archivo
  
cnnQuery:
  pocnQuery.RollbackTrans              'RESTAURA TRANSACCION.
  pocnQuery.Close
  If MsgBox(Choose(gsIdioma, "¡Información actualizada!", "Up-to-date Information") & Chr(13) & Choose(gsIdioma, " Continua actualización de información ?", " Do you Continue update of information?"), vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    GoTo Contabilidad
  End If
  porsClone.Close
  Set porsClone = Nothing
  Set pocnQuery = Nothing
  
Archivo:
  potxsQuery.Close
  Set pofsoQuery = Nothing
  Set potxsQuery = Nothing

Fin:
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dlbDirectorio_Change()
  flbArchivo.Path = dlbDirectorio.Path
  flbArchivo.Refresh
End Sub

Private Sub drvUnidad_Change()
  dlbDirectorio.Path = drvUnidad.Drive
  dlbDirectorio.Refresh
End Sub

Private Sub Form_Load()
  Dim n_Contador As Integer
  
  Me.Caption = Choose(gsIdioma, "Actualización de Sistema", " Update to System")
  drvUnidad.Drive = gsRutSis
  dlbDirectorio.Path = gsRutSis
  flbArchivo.Path = dlbDirectorio.Path
  flbArchivo.Pattern = "*.sql"
  
  ' Configuro los controles de año y mes
  For n_Contador = (Val(gsAnoAct) - 9) To Val(gsAnoAct)
    cmbEjercicio(0).AddItem Choose(gsIdioma, "Año ", "Year ") & n_Contador
    cmbEjercicio(1).AddItem Choose(gsIdioma, "Año ", "Year ") & n_Contador
  Next n_Contador
  cmbEjercicio(0).ListIndex = 9
  cmbEjercicio(1).ListIndex = 9
  
  
  Set pocnnMain = New ADODB.Connection
  Set porstTGEMP = New ADODB.Recordset
   
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDC
    .Open
  End With
  With porstTGEMP
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodEmp, RazEmp FROM TGEmp"
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  With txtDato
    For n_Contador = 0 To 1
      .Item(n_Contador).DataField = "CodEmp"
      .Item(n_Contador).MaxLength = porstTGEMP.Fields(.Item(n_Contador).DataField).DefinedSize
    Next n_Contador
  End With
   
  'Límites de rangos.
  With porstTGEMP
    .MoveLast
    txtDato(1).Text = !codemp
    .MoveFirst
    txtDato(0).Text = !codemp
  End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
  If txtDato(0).Text <> "" Then ppAyuDet 0
  If txtDato(1).Text <> "" Then ppAyuDet 1
  ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(6, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Inicio :", "Fin :", "Inicio Ejercicio", "Fin Ejercicio", "Directorio:", "Archivos:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Beginning :", "End :", "Beginning Period", "End Period", "Directory:", "Files:")
  Next nElemento
  fraOpcion.Caption = Choose(gsIdioma, "Opción", "Option")
  optParametro(0).Caption = Choose(gsIdioma, "Contabilidad", "Accounting")
  optParametro(1).Caption = Choose(gsIdioma, "Configuración", "Configuration")
  fraParametro.Caption = Choose(gsIdioma, "Parámetro", "Parameter")
  fraUbicacion(0).Caption = Choose(gsIdioma, "Carpeta", "Location")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
  optParametro(0).Value = True
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   porstTGEMP.Close
   pocnnMain.Close
   Set porstTGEMP = Nothing
   Set pocnnMain = Nothing
End Sub
Private Sub optParametro_Click(Index As Integer)
  fraParametro.Enabled = (Index = 0)
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub
Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  Select Case Index
   Case 0, 1      'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
    End If
  End Select

  'Busca el dato en su tabla principal.
  Select Case Index
   Case 0, 1                             'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
      
  Exit Sub
Err:
  gpErrores

End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  
  Select Case tnIndex
   Case 0, 1                             'Cambiar (añadir índices).
    modAyuBus.Emp_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select

End Sub
Private Function ppAyuDet(tnIndex As Integer)
  
  Select Case tnIndex                 'Cambiar.
   Case 0, 1
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstTGEMP
      .MoveFirst
      .Find "CodEmp='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & !RazEmp
      End If
    End With
  End Select

End Function
