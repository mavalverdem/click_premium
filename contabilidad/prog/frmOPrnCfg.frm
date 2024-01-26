VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmOPrnCfg 
   Caption         =   "Configuración de Impresora"
   ClientHeight    =   3510
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraImpresora 
      Caption         =   "Impresora"
      ForeColor       =   &H80000002&
      Height          =   1755
      Left            =   120
      TabIndex        =   9
      Top             =   1620
      Width           =   3615
      Begin VB.CommandButton cmdPrnCfg 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   2100
         TabIndex        =   6
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblDrvPrn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Driver"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblUbiPrn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblNomPrn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label lblOriPrn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orientac."
         Height          =   315
         Left            =   960
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTexto 
         Caption         =   "Nombre:"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTexto 
         Caption         =   "Dispositivo:"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.Frame fraParametro 
      Caption         =   "Parámetros"
      ForeColor       =   &H80000002&
      Height          =   1455
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   3195
      Begin VB.TextBox txtMargenIzq 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtCopias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Text            =   " "
         Top             =   660
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53673985
         CurrentDate     =   37041
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Margen Izquierdo:"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Emisión :"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Copias:"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame fraDestino 
      Caption         =   "Destino"
      ForeColor       =   &H80000002&
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2235
      Begin VB.OptionButton optDispositivo 
         Caption         =   "Impresora Gráfica"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optDispositivo 
         Caption         =   "Impresora Matricial"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1755
      End
   End
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   4440
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2820
      Width           =   1275
   End
End
Attribute VB_Name = "frmOPrnCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   lblNomPrn.Caption = Printer.DeviceName
   lblUbiPrn.Caption = Printer.Port
   lblDrvPrn.Caption = Printer.DriverName
   lblOriPrn.Caption = Printer.Orientation
   cdgMain.Flags = cdlPDPrintSetup
  
  '[ Cargo los mensajes de botones
  Me.Caption = Choose(gsIdioma, "Configuración de Impresora", "Print Setup")
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Fecha de Emisión :", "Nº de Copias :", "Margen Izquierdo :", "Nombre :", "Dispositivo :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Issue Date :", "Nº Copies :", "Left Margin :", "Name :", "Device :")
  Next nElemento
  fraDestino.Caption = Choose(gsIdioma, " Destino ", " Destination ")
  optDispositivo(1).Caption = Choose(gsIdioma, "Impresora Gráfica", "Graphic Printer")
  optDispositivo(0).Caption = Choose(gsIdioma, "Impresora Matricial", "Dot Matrix Printer")
  fraParametro.Caption = Choose(gsIdioma, " Párametros ", " Parameters ")
  fraImpresora.Caption = Choose(gsIdioma, " Impresora ", " Printer ")
  cmdPrnCfg.Caption = Choose(gsIdioma, "&Modificar", "&Modify")
  CaptionBotones Me, True, False, False, False, False, False, False, False, False, False, False, False, False, aLabel
 ']
   
End Sub

Private Sub cmdAceptar_Click()
    Me.Hide
End Sub

Private Sub cmdPrnCfg_Click()
   On Error GoTo Err
    
   cdgMain.ShowPrinter

   lblNomPrn.Caption = Printer.DeviceName
   lblUbiPrn.Caption = Printer.Port
   lblDrvPrn.Caption = Printer.DriverName
   lblOriPrn.Caption = Printer.Orientation
   
   cmdAceptar.SetFocus
   
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub
Private Sub txtCopias_GotFocus()
   txtCopias.SelStart = 0
   txtCopias.SelLength = txtCopias.MaxLength
End Sub

Private Sub txtCopias_LostFocus()
   If Val(txtCopias.Text) <= 0 Then txtCopias.Text = 1
End Sub

Private Sub txtMargenIzq_GotFocus()
   txtMargenIzq.SelStart = 0
   txtMargenIzq.SelLength = txtMargenIzq.MaxLength
End Sub

Private Sub txtMargenIzq_LostFocus()
   If Val(txtMargenIzq.Text) <= 0 Then txtMargenIzq.Text = 0
End Sub

Public Sub ConfiguraPrn(tnLugar As Integer, toReporte As Form)
   With toReporte
      Select Case tnLugar
      Case 0
'[ARREGLAR. Añadido para escoger el tipo de impresora desde la ventana de cada reporte (Raul 12/1/04).
.usDEstino = IIf(.optTipoImpresion(1).Value, PRN_DEST_GRAF, PRN_DEST_MATR)
']
         Select Case .usDEstino
         Case PRN_DEST_GRAF
            optDispositivo(0) = True
            optDispositivo(1) = False
         Case PRN_DEST_MATR
            optDispositivo(0) = False
            optDispositivo(1) = True
         End Select
      
         dtpFecha = .udFecha
         txtCopias = .unCopias
         txtMargenIzq = .unMargenIzquierdo / 100
   
      Case 1
         .usDEstino = IIf(optDispositivo(0).Value, PRN_DEST_GRAF, PRN_DEST_MATR)
'[ARREGLAR. Añadido para escoger el tipo de impresora desde la ventana de cada reporte.
.optTipoImpresion(0).Value = optDispositivo(1).Value
.optTipoImpresion(1).Value = optDispositivo(0).Value
']
         frmMain.rptMain.CopiesToPrinter = CInt(txtCopias)
         frmMain.rptMain.PrinterName = lblNomPrn
         frmMain.rptMain.PrinterDriver = lblDrvPrn
         frmMain.rptMain.PrinterPort = lblUbiPrn

         .udFecha = dtpFecha.Value
         .unCopias = Val(txtCopias)
         .unMargenIzquierdo = Val(txtMargenIzq.Text) * 100
      End Select
   End With
End Sub

Public Sub OrientacionPrn(tnLugar As Integer, toReporte As Form)
   With toReporte
      Select Case tnLugar
      Case 0
         toReporte.usOrientacionOri = IIf(Printer.Orientation = 1, PRN_ORIE_VERT, PRN_ORIE_HORI)
         If Printer.Orientation = 1 And toReporte.usOrientacionRpt = PRN_ORIE_HORI Or _
            Printer.Orientation = 2 And toReporte.usOrientacionRpt = PRN_ORIE_VERT Then
            SendKeys "%{" & toReporte.usOrientacionRpt & "}+{ENTER}"
            cdgMain.ShowPrinter
         End If
   
      Case 1
         If Printer.Orientation = 1 And toReporte.usOrientacionOri = PRN_ORIE_HORI Or _
            Printer.Orientation = 2 And toReporte.usOrientacionOri = PRN_ORIE_VERT Then
            SendKeys "%{" & toReporte.usOrientacionOri & "}+{ENTER}"
            cdgMain.ShowPrinter
         End If
      End Select
   End With
End Sub


