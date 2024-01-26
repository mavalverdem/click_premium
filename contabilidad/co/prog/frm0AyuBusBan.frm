VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm0AyuBusBan 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8505
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8505
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
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
         Left            =   6240
         Picture         =   "frm0AyuBusBan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   7750
         Picture         =   "frm0AyuBusBan.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
      Begin VB.Frame fraBuscar 
         Caption         =   "&Buscar por [Columna]"
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
         TabIndex        =   3
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   200
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Re&frescar"
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
         Left            =   6960
         Picture         =   "frm0AyuBusBan.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm0AyuBusBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain As ADODB.Recordset
Public ubBDConfiguracion As Boolean
Public usConnStrgSele As String, usConnStrgOrde As String
Public unArribaFormulario As Integer, _
       unIzquierdaFormulario As Integer, _
       unAltoFormulario As Integer, _
       unAnchoFormulario As Integer
Public unElementos As Integer
Public uaTitulos As Variant, uaAncho As Variant, _
       uaFormato As Variant, uaAlineamiento As Variant, _
       uaOrden As Variant
Public uvDato1Posicion As Integer, uvDato1Previo As Variant, uvDato1 As Variant
Public uvDato2Posicion As Integer, uvDato2 As Variant, uvDato3 As Variant, uvDato4 As Variant, uvDato5 As Variant, uvDato6 As Variant
Public uvDato3Posicion As Integer
Public uvDato4Posicion As Integer
Public uvDato5Posicion As Integer
Public uvDato6Posicion As Integer
Public usCriterio As String
Private pnColumnaOrd As Integer

Private Sub Form_Load()
   Me.Top = unArribaFormulario
   Me.Left = unIzquierdaFormulario
   Me.Height = unAltoFormulario
   Me.Width = unAnchoFormulario
 
 '[Recordsets                          'Cambiar.
   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
'     .ConnectionString = CONNSTRG & IIf(ubBDConfiguracion, gsRutBDC & gsNomBDC, gsRutBDS & gsNomBDS)
      .ConnectionString = CONNSTRG & IIf(ubBDConfiguracion, gsNomBDC, gsNomBDS)
      .Open
   End With
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele & usConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
      
      If .RecordCount <> 0 Then
         .Find usCriterio
         If .EOF Then
            .MoveFirst
         End If
      End If
   End With
 ']
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  Me.Caption = Choose(gsIdioma, "Ayuda", "Help")
  CaptionBotones Me, True, False, False, False, False, True, False, False, False, False, False, False, True, aLabel
  ']
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain
End Sub

Private Sub Form_Activate()
   DatosGrid
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
   If vbKeyReturn Then
      dgrMain_KeyUp KeyCode, Shift
   End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   cmdSalir.Left = Me.Width - 840
   cmdRefrescar.Left = cmdSalir.Left - 790
   cmdAceptar.Left = cmdRefrescar.Left - 720
   fraBuscar.Width = cmdAceptar.Left - fraBuscar.Left - 50
   txtBuscar.Width = fraBuscar.Width - 240
   dgrMain.Height = Me.Height - 450 - picOpciones.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   uorstMain.Close
   uocnnMain.Close
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo Err
   
   uvDato1 = dgrMain.Columns.Item(uvDato1Posicion).Value
   uvDato2 = IIf(IsNull(dgrMain.Columns.Item(uvDato2Posicion).Value), "", dgrMain.Columns.Item(uvDato2Posicion).Value)
   
   If ayudaban = True Then
        uvDato3 = IIf(IsNull(dgrMain.Columns.Item(uvDato3Posicion).Value), "", dgrMain.Columns.Item(uvDato3Posicion).Value)
        uvDato4 = IIf(IsNull(dgrMain.Columns.Item(uvDato4Posicion).Value), "", dgrMain.Columns.Item(uvDato4Posicion).Value)
        uvDato5 = IIf(IsNull(dgrMain.Columns.Item(uvDato5Posicion).Value), "", dgrMain.Columns.Item(uvDato5Posicion).Value)
        uvDato6 = IIf(IsNull(dgrMain.Columns.Item(uvDato6Posicion).Value), "", dgrMain.Columns.Item(uvDato6Posicion).Value)
   End If

   Unload Me
   
   Exit Sub
Err:
  If Err.Number = 13 Then  '13=El tipo no concide. Aparece si uvDato2 no tiene valor.
     Resume Next
  Else
     MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  End If
End Sub

Private Sub cmdRefrescar_Click()
   uorstMain.Requery
   DatosGrid
   dgrMain.SetFocus
End Sub

Private Sub cmdSalir_Click()
   uvDato1 = uvDato1Previo
   Unload Me
End Sub

Private Sub dgrMain_DblClick()
   cmdAceptar_Click
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

'  usConnStrgOrde = "ORDER BY " & dgrMain.Columns(pnColumnaOrd).DataField
   usConnStrgOrde = "ORDER BY " & IIf(uaOrden(pnColumnaOrd) = "", dgrMain.Columns(pnColumnaOrd).DataField, uaOrden(pnColumnaOrd))
   With uorstMain
      .Close
      .Source = usConnStrgSele & usConnStrgOrde
      .Open
   End With
   Set dgrMain.DataSource = uorstMain
   DatosGrid

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyHome
      IrPrimero_Click
   Case vbKeyEnd
      IrUltimo_Click
   Case vbKeyReturn
      cmdAceptar_Click
   End Select
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With uorstMain
      dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
      Select Case VarType(.Fields(pnColumnaOrd))
      Case vbString
         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
      Case vbDouble
         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'     Case vbDate
'         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
      End Select
      .Find dsCriterio, , , 1
      If .EOF = True Then
         .Bookmark = dvRegistroActual
      End If
   End With
']ARREGLAR.
   
   Exit Sub
Err:
  If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
     uorstMain.Bookmark = dvRegistroActual
  Else
     MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  End If
End Sub

Private Sub IrPrimero_Click()
   On Error GoTo Err

   uorstMain.MoveFirst

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub IrUltimo_Click()
   On Error GoTo Err

   uorstMain.MoveLast

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub DatosGrid()             'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         If dnNum < unElementos Then
            .Item(dnNum).Caption = uaTitulos(dnNum)
            .Item(dnNum).Width = uaAncho(dnNum)
            .Item(dnNum).NumberFormat = uaFormato(dnNum)
            .Item(dnNum).Alignment = uaAlineamiento(dnNum)
         Else
            .Item(dnNum).Visible = False
         End If
      Next
   End With
End Sub

Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property


