VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRTp54Ret 
   Caption         =   "[título]"
   ClientHeight    =   4170
   ClientLeft      =   2250
   ClientTop       =   2385
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5745
   Begin TabDlg.SSTab tabProceso 
      Height          =   2835
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5001
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   494
      TabCaption(0)   =   "Configuración de Parámetros"
      TabPicture(0)   =   "frmRTp54Ret.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmRegistro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmUbicacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboTpoMon"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.ComboBox cboTpoMon 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1395
         Width           =   1260
      End
      Begin VB.Frame frmUbicacion 
         Caption         =   " Carpeta "
         ForeColor       =   &H00000080&
         Height          =   2325
         Left            =   2880
         TabIndex        =   3
         Top             =   350
         Width           =   2535
         Begin VB.DriveListBox drvUnidad 
            Height          =   315
            Left            =   150
            TabIndex        =   11
            Top             =   400
            Width           =   2235
         End
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Left            =   150
            TabIndex        =   4
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   200
            Width           =   765
         End
      End
      Begin VB.Frame frmRegistro 
         Caption         =   " Registros "
         ForeColor       =   &H00000080&
         Height          =   900
         Left            =   150
         TabIndex        =   0
         Top             =   345
         Width           =   2500
         Begin VB.CheckBox chkInformacion 
            Caption         =   "&Percepciones"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   2
            Top             =   510
            Value           =   1  'Checked
            Width           =   2200
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "&Retenciones"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   250
            Value           =   1  'Checked
            Width           =   2200
         End
      End
      Begin VB.Label lblTexto 
         Caption         =   "Moneda"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   555
         TabIndex        =   13
         Top             =   1455
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   400
      Left            =   1380
      TabIndex        =   9
      Top             =   3675
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   400
      Left            =   3060
      TabIndex        =   8
      Top             =   3675
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Archivo:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1785
   End
End
Attribute VB_Name = "frmRTp54Ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
  On Error GoTo Err
  
  If MsgBox("¿ Estás Seguro de Generar archivo de información ? ", vbQuestion + vbYesNo) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    pgbProgreso(0).Value = 0: pgbProgreso(0).Min = 0
    ' Genero los archivos de información
    ppGenArchivo
    
    MsgBox TEXT_8008, vbInformation
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    cmdSalir.SetFocus
  End If
  Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Function pfOutApoRet(s_Expresion As String) As String

  s_Expresion = Trim$(s_Expresion)
  If s_Expresion <> "" Then
    ' saco los enters de la cadena de caracteres
    While InStr(s_Expresion, Chr(13)) <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(13)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(13)) + 1)
    Wend
    ' saco los retornos de la cadena de caracteres
    While InStr(s_Expresion, Chr(10)) <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(10)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(10)) + 1)
    Wend
    ' saco los apostrofes de la cadena de caracteres
    While InStr(s_Expresion, "'") <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "'") - 1) & "´" & Mid$(s_Expresion, InStr(s_Expresion, "'") + 1)
    Wend
    ' saco los rayas de la cadena de caracteres
    While InStr(s_Expresion, "|") <> 0
      s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "|") - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, "|") + 1)
    Wend
  End If
  pfOutApoRet = Trim$(s_Expresion)

End Function

Private Sub ppGenArchivo()
  
  Dim sSentencia As String, sLinea As String
  Dim nContador As Integer, nArchivo As Integer
  Dim nRegistro As Double, nNumRegistros As Long
  Dim sArchivo As String, nSecuencia As Integer
  Dim sCaracter  As String, sRegistro As String

  Dim pocnnMain As ADODB.Connection
  Dim porstTmp As ADODB.Recordset

  ' Seteo y activo la coneccion
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  
  ' Seteo el recordset temporal
  Set porstTmp = New ADODB.Recordset
  sCaracter = " "
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkInformacion.Count - 1
    ' Verifico que se haya seleccionado
    If chkInformacion(nContador).Value = vbChecked Then
      ' Obtengo el archivo de texto libre
      nArchivo = FreeFile
      
      ' Generacion de la tabla de seleccion
      sSentencia = "SELECT DISTINCT prv.codaux, prv.refdoc, aux.rucaux, prv.feedoc, prv.codtdc, prv.serdoc, prv.nrodoc, "
      If nContador = 0 Then
        sSentencia = sSentencia & "prv.imptot_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & " AS cimporte "
        
      Else
        sSentencia = sSentencia & "ROUND(prv.impoi1_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "+" & "prv.impoi2_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & ", 2) AS cimporte "
        
      End If
      sSentencia = sSentencia & "FROM " & Choose(nContador + 1, "covtadoc", "cocprdoc") & " prv "
      sSentencia = sSentencia & "LEFT JOIN tgaux aux ON prv.codemp=aux.codemp AND prv.codaux=aux.codaux "
      sSentencia = sSentencia & "WHERE prv.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND prv.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND prv.mespvs='" & gsMesAct & "' "
      sSentencia = sSentencia & "AND prv.categoriadoc>=" & CategoriaDocumento.RetencionIva & " "
      sSentencia = sSentencia & "ORDER BY prv.codaux, prv.feedoc"
      ' Abro el recordset temporal
      With porstTmp
        If .State = adStateOpen Then .Close
        .ActiveConnection = pocnnMain
        .Source = sSentencia
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
      End With
      If Not (porstTmp.BOF And porstTmp.EOF) Then
        ' Barro todo el recordset y lo grabo en el archivo
        lblProgreso(0).Caption = Choose(gsIdioma, "Exportando Archivo:", "Exporting File:") & Mid(Trim(chkInformacion(nContador).Caption), 2)
        nNumRegistros = porstTmp.RecordCount
        pgbProgreso(0).Max = nNumRegistros
        pgbProgreso(0).Value = pgbProgreso(0).Min
        nRegistro = 0
        ' Nombre del archivo de texto a generar
        sArchivo = dlbDirectorio.Path & "\" & gsRUCEmp & gsAnoAct & gsMesAct & Choose(nContador + 1, "r", "p") & ".txt"
        ' Elimino archivo de texto si existe
        If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
        If Dir$(sArchivo, vbNormal) = "" Then
          Open sArchivo For Output Access Write Lock Read Write As #nArchivo
          porstTmp.MoveFirst
          While Not porstTmp.EOF
            nRegistro = nRegistro + 1
            ' Diseño y grabro la linea en el archivo
            sLinea = ""
            sRegistro = gfPadR(Left(IIf(IsNull(porstTmp!refdoc), "", porstTmp!refdoc), 3), 3, sCaracter)
            sLinea = sLinea & sRegistro
            sRegistro = gfPadL(porstTmp!RucAux, 13, sCaracter)
            sLinea = sLinea & sRegistro
            sLinea = sLinea & Format(porstTmp!FeEDoc, "dd/mm/yyyy")
            sRegistro = gfPadL(porstTmp!SerDoc, 4, sCaracter)
            sLinea = sLinea & sRegistro
            sRegistro = gfPadR(porstTmp!NroDoc, 12, sCaracter)
            sLinea = sLinea & sRegistro
            sRegistro = Format(CDec(porstTmp!cimporte), "############0.00")
            If Right(sArchivo, 5) = "r.txt" Then
               sRegistro = gfPadL(sRegistro, 14, sCaracter)
            Else
               sRegistro = gfPadL(sRegistro, 16, sCaracter)
            End If
            sLinea = sLinea & sRegistro
            Print #nArchivo, sLinea
            pgbProgreso(0).Value = nRegistro
            DoEvents
            porstTmp.MoveNext
          Wend
          Close #nArchivo
        End If
      End If
      porstTmp.Close
    End If
  Next nContador
  ' Cierro y saco de memoria los objetos
  Set porstTmp = Nothing
  pocnnMain.Close
  Set pocnnMain = Nothing

End Sub

Private Sub drvUnidad_Change()
  dlbDirectorio.Path = drvUnidad.Drive
  dlbDirectorio.Refresh
End Sub

Private Sub Form_Activate()
   
  drvUnidad.Drive = gsRutSis
  dlbDirectorio.Path = gsRutSis
  
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  cmdSalir.SetFocus

End Sub

Private Sub Form_Load()
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Directorio :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Directory :", "Currency :")
  Next nElemento
  tabProceso.TabCaption(0) = Choose(gsIdioma, "Configuración de Parámetros", "Configuration of Parameters")
  frmUbicacion.Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  frmRegistro.Caption = Choose(gsIdioma, " Registros ", " Registers ")
  chkInformacion(0).Caption = Choose(gsIdioma, "&Retenciones", "&Withholding")
  chkInformacion(1).Caption = Choose(gsIdioma, "&Percepciones", "&Perceptions")
  lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Archivo:", "Processing File:")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
End Sub
