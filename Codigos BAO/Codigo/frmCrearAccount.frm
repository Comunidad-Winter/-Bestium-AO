VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   0  'None
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCrearAccount.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5580
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckRegla 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   200
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox CheckRegistro 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   200
      Left            =   480
      TabIndex        =   5
      Top             =   3870
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.TextBox Remail 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   310
      Left            =   360
      TabIndex        =   4
      Top             =   3090
      Width           =   3015
   End
   Begin VB.TextBox Mail 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   310
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox repass 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   310
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1700
      Width           =   3015
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   310
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Nombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   310
      Left            =   360
      TabIndex        =   0
      Top             =   450
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   720
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Image Command2 
      Height          =   525
      Left            =   360
      Top             =   4800
      Width           =   1260
   End
   Begin VB.Image Command1 
      Height          =   525
      Left            =   2160
      Top             =   4800
      Width           =   1260
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

With frmMensaje

If CheckRegla.value = Unchecked Then
.Show vbModal
.lblmensaje.Caption = "Debe aceptar el reglamento."
Exit Sub
End If

If Pass <> repass Then
.Show vbModal
.lblmensaje.Caption = "Las passwords que tipeó no coinciden."
Exit Sub
End If

If Remail <> Mail Then
.Show vbModal
.lblmensaje.Caption = "Los mails que tipeó no coinciden."
Exit Sub
End If

If Not CheckMailString(Mail) Then
.Show vbModal
.lblmensaje.Caption = "Dirección de mail invalida."
Exit Sub
End If

If CheckRegistro.value = Checked Then
Call ShellExecute(0, "Open", "http://tierras-perdidas.com/f/register.php?ID=&", "", App.Path, SW_SHOWNORMAL)
End If

If nombre = "" Or Pass = "" Or repass = "" Or Mail = "" Then
.Show vbModal
.lblmensaje.Caption = "Completá todos los campos."
Exit Sub
End If

End With

frmCrearCuenta.Show vbModal

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmCrearAccount.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_Main.jpg")
Command2.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BAtrasN.jpg")
Command1.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BCrearN.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BAtrasN.jpg")
Command1.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BCrearN.jpg")
End Sub

Private Sub command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BAtrasA.jpg")
End Sub

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BAtrasI.jpg")
End Sub

Private Sub command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BCrearA.jpg")
End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Picture = LoadPicture(DirInterfaces & "Principal\CrearCuenta_BCrearI.jpg")
End Sub

Private Sub Image1_Click()
Call ShellExecute(0, "Open", "http://tierras-perdidas.com/f/forumdisplay.php?s=f0210bdfe8050acc2839bf8832422f1f&f=115", "", App.Path, SW_SHOWNORMAL)
End Sub
