VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fPrnTituloPlanilla 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3225
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "prntituloplla.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6420
   Begin TabDlg.SSTab tabRegister 
      Height          =   2040
      Left            =   75
      TabIndex        =   9
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   3598
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabMaxWidth     =   3052
      BackColor       =   -2147483644
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "prntituloplla.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmCuadro(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin Threed.SSFrame frmCuadro 
         Height          =   900
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   690
         Width           =   5925
         _Version        =   65536
         _ExtentX        =   10451
         _ExtentY        =   1587
         _StockProps     =   14
         Caption         =   " Rango de Pagina "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtPagina 
            Height          =   300
            Index           =   0
            Left            =   1485
            TabIndex        =   3
            Top             =   375
            Width           =   900
         End
         Begin VB.TextBox txtPagina 
            Height          =   300
            Index           =   1
            Left            =   4185
            TabIndex        =   5
            Top             =   375
            Width           =   900
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Desde :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   2
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Hasta :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   3075
            TabIndex        =   4
            Top             =   420
            Width           =   1005
         End
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   200
         TabIndex        =   0
         Top             =   225
         Width           =   375
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
      Begin Threed.SSCommand cmdCancel 
         Height          =   360
         Left            =   5790
         TabIndex        =   10
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
         Picture         =   "prntituloplla.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   5400
         TabIndex        =   11
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prntituloplla.frx":0044
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
         Left            =   285
         TabIndex        =   7
         Top             =   120
         Width           =   4800
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   2715
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
   End
End
Attribute VB_Name = "fPrnTituloPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private n_PapelSize As Integer, n_PapelPos As Integer   ' Tamaño y orientación del papel
Private n_FontSize As Integer                           ' Tamaño de caracteres
Private n_LonSize As Integer, n_LenSize As Integer      ' Dimensiones especiales
'[
Private Function GeneraTitulo(ByVal s_Cadena As String, ByVal s_Expresion As String, ByVal n_Posicion As Integer, ByVal n_Longitud As Integer, ByVal s_Caracter As String, ByVal s_Tipo As String) As String
  Dim nLenCadena As Integer, nLen As Integer
  Dim nPosInicio As Integer
  
  nPosInicio = IIf(Left(s_Expresion, 1) = "|" Or Left(s_Expresion, 1) = "[", 2, 1)
  If s_Tipo = "D" Then
    s_Expresion = gdl_Funcion.PadR(s_Expresion, n_Longitud, Chr(32))
  Else
    s_Expresion = gdl_Funcion.PadL(s_Expresion, n_Longitud, Chr(32))
  End If
  s_Expresion = Replace(s_Expresion, "[", "", 1, 1, vbTextCompare)
  ' Caracter de separación de campos
  nLenCadena = Len(s_Cadena)
  nLen = nLenCadena - n_Posicion
  If n_Posicion >= nLenCadena Then
    s_Cadena = s_Cadena & String((n_Posicion - nLenCadena), Chr(32))
  Else
    n_Longitud = n_Longitud - nLen
    n_Longitud = IIf(n_Longitud > 0, n_Longitud, 0)
    If s_Tipo = "C" Or s_Tipo = "G" Then
      s_Expresion = Right(s_Expresion, n_Longitud)
    ElseIf s_Tipo = "D" Then
      s_Expresion = Left(s_Expresion, n_Longitud)
    End If
    s_Caracter = IIf(n_Longitud > 0, s_Caracter, "")
  End If
  n_Longitud = n_Longitud - (nPosInicio - 1)
  n_Longitud = IIf(n_Longitud > 0, n_Longitud, 0)
  s_Expresion = Choose(nPosInicio, "", s_Caracter) & IIf(s_Tipo = "D", Left(s_Expresion, n_Longitud), Right(s_Expresion, n_Longitud))
  s_Cadena = s_Cadena & s_Expresion
  GeneraTitulo = s_Cadena
    
End Function
Private Sub ImprimeTitulo(ByVal a_Cabecera, ByVal s_RegPatronal As String, ByVal s_Direccion As String, ByVal s_TituloReporte As String, ByVal s_Periodo As String, s_Pagina As String)
  Dim nDetalle As Integer, nFontSize As Integer
  Dim nColIni As Integer, nColFin As Integer
  Dim nLongitud As Integer, nFilaPrn As Double
  
  Printer.Font.Bold = True
  Printer.Font.Underline = False
  Printer.Font.Italic = False
  Printer.CurrentY = 0
  Printer.CurrentX = 0
  nColIni = Choose(n_LonSize, Choose(n_LenSize, 160, 150, 140), Choose(n_LenSize, 169, 162, 157), Choose(n_LenSize, 165, 157, 150))
  nColFin = Choose(n_LonSize, Choose(n_LenSize, 179, 173, 168), Choose(n_LenSize, 182, 179, 176), Choose(n_LenSize, 181, 177, 173))
  
  nFontSize = n_FontSize + 4
  Printer.FontSize = nFontSize
  Printer.Print ps_NomEmpresa;
  
  nFontSize = n_FontSize + 2
  Printer.FontSize = nFontSize
  Printer.CurrentX = nColIni: Printer.Print "Reg. Patronal :";
  Printer.Font.Bold = False
  Printer.CurrentX = nColFin: Printer.Print s_RegPatronal
  
  nFontSize = n_FontSize + 4
  Printer.FontSize = nFontSize
  Printer.Font.Bold = True
  Printer.CurrentX = 0: Printer.Print s_Direccion;
  
  nFontSize = n_FontSize + 2
  Printer.FontSize = nFontSize
  Printer.CurrentX = nColIni: Printer.Print "R.U.C. :";
  Printer.Font.Bold = False
  Printer.CurrentX = nColFin: Printer.Print ps_RucEmpresa
   
  nFontSize = n_FontSize + 3
  Printer.FontSize = nFontSize
  Printer.Font.Bold = True
  Printer.Print
  nLongitud = Choose(n_LonSize, Choose(n_LenSize, 150, 122, 103), Choose(n_LenSize, 212, 174, 147), Choose(n_LenSize, 188, 146, 123))
  Printer.CurrentX = 0: Printer.Print gdl_Funcion.PadC(s_TituloReporte, nLongitud, Chr(32))
'  Printer.CurrentX = 0: Printer.Print "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
'  Printer.CurrentX = 0: Printer.Print "        10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220       230       240       250"
  nLongitud = nLongitud - 20
  Printer.CurrentX = 0: Printer.Print gdl_Funcion.PadC(s_Periodo, nLongitud, Chr(32));
  'CurrentY = 4.766457
  '1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
  '        10        20        30        40        50        60        70        80        90       100       110       120       130       140       150       160       170       180       190       200       210       220       230       240       250
  
  nFontSize = n_FontSize + 2
  Printer.FontSize = nFontSize
  Printer.CurrentX = nColIni: Printer.Print "Pagina :";
  Printer.Font.Bold = False
  Printer.CurrentX = nColFin: Printer.Print s_Pagina
  
  nFontSize = n_FontSize + 1
  Printer.FontSize = nFontSize
  Printer.Font.Bold = True
  Printer.Print String(272, "-")
  
  nFontSize = n_FontSize
  Printer.FontSize = nFontSize
  nFilaPrn = Printer.CurrentY
  For nDetalle = 1 To UBound(a_Cabecera)
    If Trim(a_Cabecera(nDetalle)) <> "x" Then
      Printer.Print Mid(a_Cabecera(nDetalle), 2)
    End If
  Next nDetalle
  nFontSize = n_FontSize + 1
  Printer.FontSize = nFontSize
  Printer.Print String(272, "-")

End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  Dim sDireccion As String, sRegPatronal As String
  Dim sTituloReporte As String, sPeriodo As String
  Dim s_OldMessage As String
  Dim s_Registro As String, a_sCabecera(7) As String
  Dim nContador As Long
  
  ' Realizo las validaciones de los campos a actualizar
  If Not IsNumeric(txtPagina(0).Text) Then Beep: MsgBox "Número de pagina inicial es invalido", vbExclamation: txtPagina(0).SetFocus: Exit Sub
  If CInt(txtPagina(0).Text) <= 0 Then Beep: MsgBox "Número de pagina inicial es invalido", vbExclamation: txtPagina(0).SetFocus: Exit Sub
  If Not IsNumeric(txtPagina(1).Text) Then Beep: MsgBox "Número de pagina final es invalido", vbExclamation: txtPagina(1).SetFocus: Exit Sub
  If CInt(txtPagina(1).Text) <= 0 Then Beep: MsgBox "Número de pagina final es invalido", vbExclamation: txtPagina(1).SetFocus: Exit Sub
  If CInt(txtPagina(1).Text) < CInt(txtPagina(0).Text) Then Beep: MsgBox "Número de pagina final debe ser mayor e igual que inicial", vbExclamation: txtPagina(1).SetFocus: Exit Sub
  If ps_RucEmpresa = "" Then Beep: MsgBox "Número de RUC invalido; Verifique", vbExclamation: Exit Sub
  
  ' Obtengo los datos de la empresa
  sDireccion = "": sRegPatronal = ""
  s_Sql = "SELECT via.abrevia, prm.direccionvia, prm.numerodir, zon.abrezona, prm.direccionzona, prm.ubigeodir, prm.regpatronal "
  s_Sql = s_Sql & "FROM plcfgempresa prm "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON prm.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON prm.codzona=zon.codzona "
  s_Sql = s_Sql & "WHERE prm.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    sRegPatronal = gdl_Funcion.aTexto(porstRecordset!regpatronal)
    sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
    sDireccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
    sDireccion = gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!direccionzona) & " - " & sDireccion
  End If
  porstRecordset.Close
  sPeriodo = "MES DE "
  sTituloReporte = "P L A N I L L A  D E  P A G O  D E  R E M U N E R A C I O N E S"
    
  Beep
  If MsgBox("¿ Estás Seguro de Imprimir " & lblTitle & " ?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    fMenu.panPercent.Visible = True
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    
    ' Obtengo la información de los titulos
    s_Sql = "SELECT pll.fila, pll.columna, pll.posicion, pll.tipo, pll.alias, pll.descripcion, pll.longitud, "
    s_Sql = s_Sql & "pll.subrayado, pll.sizefont, pll.sizepapel, pll.posipapel, pll.imprimecab, pll.despll "
    s_Sql = s_Sql & "FROM plplanilla pll "
    s_Sql = s_Sql & "WHERE pll.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pll.codpll='" & fPlanillaGnral.dcaRegistro.Recordset!codpll & "' "
    s_Sql = s_Sql & "ORDER BY pll.fila, pll.columna"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    For n_Index = 1 To UBound(a_sCabecera): a_sCabecera(n_Index) = "x": Next n_Index
    ' Si hay registros de configuración
    If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
      ' Inicializo las variables
      n_PapelSize = Choose(CInt(porstRecordset!sizepapel) + 1, vbPRPSA4, vbPRPSA3, vbPRPSFanfoldUS)
      n_PapelPos = Choose(CInt(porstRecordset!posipapel), vbPRORPortrait, vbPRORLandscape)
      n_FontSize = CInt(porstRecordset!sizefont)
      n_Index = 0
      While Not porstRecordset.EOF
        s_Registro = UCase(Left(gdl_Funcion.aTexto(porstRecordset("descripcion")), CInt(porstRecordset("longitud"))))
        n_Index = CInt(porstRecordset("fila"))
        a_sCabecera(n_Index) = GeneraTitulo(a_sCabecera(n_Index), s_Registro, CInt(porstRecordset("posicion")), CInt(porstRecordset("longitud")), "|", Trim(porstRecordset("tipo")))
        porstRecordset.MoveNext
      Wend
    End If
    ' Dimensiones A4  6 :192, 8 :152 y 10 :122
    ' Dimensiones A3  6 :272, 8 :212 y 10 :172
    ' Dimensiones Continuo  6 :230, 8 :180 y 10 :150
    n_LonSize = IIf(n_PapelSize = vbPRPSA4, 1, IIf(n_PapelSize = vbPRPSA3, 2, 3))
    n_LenSize = IIf(n_FontSize = 6, 1, IIf(n_FontSize = 8, 2, 3))
    
    ' Parametros iniciales de impresión
    Printer.ScaleMode = vbCharacters
    Printer.PaperSize = n_PapelSize
    Printer.Orientation = n_PapelPos
    Printer.PrintQuality = vbPRPQMedium
    Printer.Font = "Courier New"
    Printer.ScaleWidth = 200
    Printer.ScaleHeight = 100
    Printer.ScaleLeft = -2
    Printer.ScaleTop = -2

    ' Imprimo los titulos
    For nContador = CInt(txtPagina(0).Text) To CInt(txtPagina(1).Text)
      ImprimeTitulo a_sCabecera, sRegPatronal, sDireccion, sTituloReporte, sPeriodo, Format(nContador, "00000")
      Printer.NewPage
      ' Incremento el porcentaje
      fMenu.panPercent.FloodPercent = ((nContador * 100) \ CInt(txtPagina(1).Text))
    Next nContador
    Printer.EndDoc
    
    fMenu.panPercent.FloodPercent = 0
    fMenu.panPercent.Visible = False
    MuestraMensaje s_OldMessage
    ' Coloco el puntero en normal
    gdl_Procedure.PunteroNormal
  End If

End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 3700: Me.Width = 6510
  Me.Left = 4580: Me.Top = 3500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Impresión de Titulos de Planilla"
  lblTitle = "Titulo de Planilla"
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "reporte": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "imprimir"
  aElemento(0, 2) = "Imprime Cabecera de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdUpdate, aElemento
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Carga los datos en el formulario
  lblTexto.Caption = " " & fPlanillaGnral.dcaRegistro.Recordset!codpll & " - " & fPlanillaGnral.dcaRegistro.Recordset!despll & " "
  gdl_Procedure.EditText "AT", txtPagina(0), CInt(0), s_MdoData_Ins, False, 5
  gdl_Procedure.EditText "AT", txtPagina(1), CInt(0), s_MdoData_Ins, False, 5
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub txtPagina_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPagina(Index)
End Sub
Private Sub txtPagina_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPagina_Validate(Index As Integer, Cancel As Boolean)
  txtPagina(Index).Text = IIf(Not IsNumeric(txtPagina(Index).Text), 0, txtPagina(Index).Text)
  txtPagina(Index).Text = FormatNumber(CInt(txtPagina(Index).Text), 0)
End Sub
