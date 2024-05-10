VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox OcultarPass 
      Height          =   195
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.TextBox txtRespuesta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H0000C0C0&
      Height          =   295
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtPregunta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H0000C0C0&
      Height          =   295
      Left            =   330
      TabIndex        =   0
      Top             =   860
      Width           =   3375
   End
   Begin VB.Image Aceptar 
      Height          =   570
      Left            =   240
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Image Salir 
      Height          =   585
      Left            =   2520
      Top             =   3120
      Width           =   1155
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Picture = LoadPicture(DirInterfaces & "\Principal\Pregunta.jpg")
Aceptar.Picture = LoadPicture(DirInterfaces & "\Principal\Preguntaentrar.jpg")
Salir.Picture = LoadPicture(DirInterfaces & "\Principal\Preguntasalir.jpg")

End Sub

Private Sub aceptar_click()

If txtPregunta.Text = "" Or txtRespuesta.Text = "" Then MsgBox "Rellená todos los campos."

Call SendData("NACCNT" & frmCrearAccount.nombre & "," & frmCrearAccount.Pass & "," & frmCrearAccount.Mail & "," & txtPregunta & "," & txtRespuesta)

Unload Me

End Sub

Private Sub OcultarPass_Click()

If OcultarPass.value = Unchecked Then
txtRespuesta.PasswordChar = ""

Else
txtRespuesta.PasswordChar = "*"

End If

End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Aceptar.Picture = LoadPicture(DirInterfaces & "\Principal\Preguntaentrar.jpg")
Salir.Picture = LoadPicture(DirInterfaces & "\Principal\Preguntasalir.jpg")
End Sub

Private Sub aceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Aceptar.Picture = LoadPicture(DirInterfaces & "Principal\preguntaEntrarA.jpg")
End Sub

Private Sub Salir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Salir.Picture = LoadPicture(DirInterfaces & "Principal\preguntasalirA.jpg")
End Sub
